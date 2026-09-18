"""Whether this worker should keep taking jobs.

A worker whose iSpring add-in is broken, or whose browser is signed out,
fails every job it touches. Left alone it will empty the queue into the
failed pile in twenty minutes: a hundred materials marked failed because of
one thing wrong on one machine.

So the worker counts consecutive failures and, past a limit, declares itself
unhealthy and stops asking for work. The queue stays intact, one job is lost
instead of a hundred, and everything resumes when the machine is put right.

The state is a small JSON file so it survives a restart — a machine that
reboots into the same broken state must not start pulling jobs again.
"""

from __future__ import annotations

import json
import logging
import os
import time
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger(__name__)

STAGE = "worker_health"
STATE_NAME = "worker-health.json"
# Two failures can be bad luck. Three in a row is the machine.
DEFAULT_LIMIT = 3
# Stages that mean the machine is broken, not the deck.
MACHINE_STAGES = {"powerpoint_open", "ispring_publish", "worker_health"}


@dataclass
class Health:
    consecutive_failures: int = 0
    unhealthy_since: float | None = None
    reason: str = ""
    last_stage: str = ""
    history: list[str] = field(default_factory=list)

    @property
    def healthy(self) -> bool:
        return self.unhealthy_since is None

    def as_dict(self) -> dict:
        return {
            "consecutive_failures": self.consecutive_failures,
            "unhealthy_since": self.unhealthy_since,
            "reason": self.reason,
            "last_stage": self.last_stage,
            "history": self.history[-10:],
        }


def state_path(root: str | Path) -> Path:
    return Path(root) / STATE_NAME


def load(root: str | Path) -> Health:
    path = state_path(root)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return Health()
    return Health(
        consecutive_failures=int(data.get("consecutive_failures") or 0),
        unhealthy_since=data.get("unhealthy_since"),
        reason=str(data.get("reason") or ""),
        last_stage=str(data.get("last_stage") or ""),
        history=list(data.get("history") or []),
    )


def save(root: str | Path, health: Health) -> None:
    path = state_path(root)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(health.as_dict(), indent=2), encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        logger.warning(
            "could not write the health file",
            extra={"stage": STAGE, "file": str(path), "error": str(exc)},
        )


def record_success(root: str | Path) -> Health:
    """A job went through. The machine works; forget the past failures."""
    health = load(root)
    if health.consecutive_failures or not health.healthy:
        logger.info(
            "worker is healthy again",
            extra={"stage": STAGE, "after_failures": health.consecutive_failures},
        )
    health.consecutive_failures = 0
    health.unhealthy_since = None
    health.reason = ""
    health.last_stage = ""
    save(root, health)
    return health


def record_failure(
    root: str | Path,
    *,
    stage: str = "",
    message: str = "",
    limit: int = DEFAULT_LIMIT,
) -> Health:
    """A job failed. Past the limit, stop taking work.

    Only failures that point at the machine count. A password-protected deck
    or a dead download URL is that job's problem, and holding the whole
    worker back for it would be wrong.
    """
    health = load(root)
    if stage and stage not in MACHINE_STAGES:
        logger.info(
            "job failed for its own reasons; the worker stays healthy",
            extra={"stage": STAGE, "failed_stage": stage},
        )
        health.consecutive_failures = 0
        save(root, health)
        return health

    health.consecutive_failures += 1
    health.last_stage = stage
    health.history.append(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {stage}: {message[:120]}")
    if health.consecutive_failures >= limit and health.healthy:
        health.unhealthy_since = time.time()
        health.reason = f"{health.consecutive_failures} failures in a row at {stage or 'unknown'}"
        logger.error(
            "worker is unhealthy and will stop taking jobs; "
            "fix the machine, then clear the health file",
            extra={
                "stage": STAGE,
                "reason": health.reason,
                "file": str(state_path(root)),
            },
        )
    else:
        logger.warning(
            "job failed",
            extra={
                "stage": STAGE,
                "consecutive": health.consecutive_failures,
                "limit": limit,
                "failed_stage": stage,
            },
        )
    save(root, health)
    return health


def should_take_work(root: str | Path) -> bool:
    """The question the job loop asks before requesting the next material."""
    if os.environ.get("WORKER_IGNORE_HEALTH", "").strip().lower() in {"1", "true", "yes"}:
        return True
    health = load(root)
    if health.healthy:
        return True
    logger.error(
        "not asking for work: this worker is marked unhealthy",
        extra={"stage": STAGE, "reason": health.reason},
    )
    return False


def clear(root: str | Path) -> None:
    """Put the worker back in service after the machine has been fixed."""
    save(root, Health())
    logger.info("worker health cleared", extra={"stage": STAGE})
