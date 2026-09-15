from pathlib import Path

import pytest

from config.settings import Settings
from publisher.ispring import get_publisher
from utils.exceptions import ISpringNotConfiguredError
from validator.html5 import validate_ispring_output


def _settings(tmp_path: Path, adapter: str) -> Settings:
    return Settings(
        app_env="development",
        backend_base_url="https://backend.example.com",
        temp_root=str(tmp_path / "jobs"),
        log_root=str(tmp_path / "logs"),
        ispring_adapter=adapter,
    )


def test_not_configured_is_the_default_and_raises(tmp_path):
    publisher = get_publisher(_settings(tmp_path, "not_configured"))
    assert publisher.is_available().installed is False
    with pytest.raises(ISpringNotConfiguredError) as exc:
        publisher.publish(tmp_path / "source.pptx", tmp_path / "out", timeout_s=1)
    assert "probe_ispring.py" in str(exc.value)


@pytest.mark.parametrize("adapter", ["vba", "uia", "cli"])
def test_unverified_adapters_do_not_invent_an_api(tmp_path, adapter):
    publisher = get_publisher(_settings(tmp_path, adapter))
    with pytest.raises(ISpringNotConfiguredError):
        publisher.publish(tmp_path / "source.pptx", tmp_path / "out", timeout_s=1)


def test_fake_copies_the_fixture_package(tmp_path):
    output = tmp_path / "out"
    publisher = get_publisher(_settings(tmp_path, "fake"))
    result = publisher.publish(tmp_path / "source.pptx", output, timeout_s=1)
    assert result == output
    validate_ispring_output(output)
