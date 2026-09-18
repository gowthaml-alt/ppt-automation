from pathlib import Path

from worker.health import (
    DEFAULT_LIMIT,
    clear,
    load,
    record_failure,
    record_success,
    should_take_work,
)


def test_a_new_worker_takes_work(tmp_path: Path):
    assert should_take_work(tmp_path) is True


def test_three_machine_failures_stop_the_worker(tmp_path: Path):
    for _ in range(DEFAULT_LIMIT):
        record_failure(tmp_path, stage="ispring_publish", message="no ribbon tab")
    assert should_take_work(tmp_path) is False
    assert "3 failures in a row" in load(tmp_path).reason


def test_a_bad_deck_does_not_stop_the_worker(tmp_path: Path):
    for _ in range(DEFAULT_LIMIT + 2):
        record_failure(tmp_path, stage="download", message="password protected")
    # The deck's problem, not the machine's.
    assert should_take_work(tmp_path) is True


def test_a_success_clears_the_count(tmp_path: Path):
    record_failure(tmp_path, stage="ispring_publish", message="one")
    record_failure(tmp_path, stage="ispring_publish", message="two")
    record_success(tmp_path)
    assert load(tmp_path).consecutive_failures == 0
    record_failure(tmp_path, stage="ispring_publish", message="three")
    assert should_take_work(tmp_path) is True


def test_the_state_survives_a_restart(tmp_path: Path):
    for _ in range(DEFAULT_LIMIT):
        record_failure(tmp_path, stage="powerpoint_open", message="add-in gone")
    # A reboot reads the same file back: a broken machine must not resume.
    assert should_take_work(tmp_path) is False
    clear(tmp_path)
    assert should_take_work(tmp_path) is True
