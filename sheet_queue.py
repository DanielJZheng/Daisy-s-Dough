from __future__ import annotations

import asyncio
import logging
from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from typing import Any

from gspread.exceptions import APIError

logger = logging.getLogger("daisys-dough.queue")


@dataclass
class _Job:
    action: Callable[[], Any]
    on_success: Callable[[], Awaitable[None]]
    on_failure: Callable[[Exception], Awaitable[None]]
    attempts: int = 0


class SheetWriteQueue:
    """Run Google Sheets writes one at a time; wait and retry on 429."""

    def __init__(
        self,
        *,
        min_interval: float = 1.5,
        max_attempts: int = 15,
        base_backoff: float = 8.0,
    ) -> None:
        self._min_interval = min_interval
        self._max_attempts = max_attempts
        self._base_backoff = base_backoff
        self._queue: asyncio.Queue[_Job | None] = asyncio.Queue()
        self._worker: asyncio.Task[None] | None = None

    def start(self) -> None:
        if self._worker is None:
            self._worker = asyncio.create_task(self._worker_loop())

    async def stop(self) -> None:
        if self._worker is not None:
            await self._queue.put(None)
            await self._worker
            self._worker = None

    async def submit(
        self,
        action: Callable[[], Any],
        *,
        on_success: Callable[[], Awaitable[None]],
        on_failure: Callable[[Exception], Awaitable[None]],
    ) -> None:
        self.start()
        await self._queue.put(
            _Job(action=action, on_success=on_success, on_failure=on_failure)
        )

    def waiting(self) -> int:
        return self._queue.qsize()

    async def _worker_loop(self) -> None:
        while True:
            job = await self._queue.get()
            if job is None:
                break
            while True:
                try:
                    await asyncio.to_thread(job.action)
                except APIError as error:
                    if "429" in str(error) and job.attempts < self._max_attempts:
                        job.attempts += 1
                        delay = min(
                            self._base_backoff * (2 ** (job.attempts - 1)),
                            120.0,
                        )
                        logger.warning(
                            "Sheets read quota (429) — hold %ss, retry %s/%s (%s queued)",
                            delay,
                            job.attempts,
                            self._max_attempts,
                            self.waiting(),
                        )
                        await asyncio.sleep(delay)
                        continue
                    await job.on_failure(error)
                    break
                except Exception as error:
                    await job.on_failure(error)
                    break
                else:
                    await job.on_success()
                    break
            await asyncio.sleep(self._min_interval)
