import pytest

from utils.paths import (
    assert_within,
    build_job_paths,
    create_job_dirs,
    free_disk_gb,
    iter_files,
)


def test_build_job_paths_uses_the_queue_id_only(tmp_path):
    paths = build_job_paths(tmp_path, 101)
    assert paths.root == tmp_path / "101"
    assert paths.source_pptx == paths.input_dir / "source.pptx"
    assert paths.partial_pptx == paths.input_dir / "source.pptx.part"
    assert paths.working_pptx == paths.working_dir / "source.pptx"


def test_build_job_paths_rejects_a_non_integer(tmp_path):
    with pytest.raises(TypeError):
        build_job_paths(tmp_path, "../../etc")


def test_create_job_dirs_creates_every_subdirectory(tmp_path):
    paths = build_job_paths(tmp_path, 7)
    create_job_dirs(paths)
    assert paths.output_dir.is_dir()


def test_assert_within_rejects_an_escaping_path(tmp_path):
    with pytest.raises(ValueError, match="outside"):
        assert_within(tmp_path, tmp_path / ".." / "escaped.txt")


def test_free_disk_gb_is_positive(tmp_path):
    assert free_disk_gb(tmp_path) > 0


def test_iter_files_yields_posix_relative_paths(tmp_path):
    (tmp_path / "res" / "deep").mkdir(parents=True)
    (tmp_path / "index.html").write_text("x")
    (tmp_path / "res" / "deep" / "app.js").write_text("y")
    found = {rel for _, rel in iter_files(tmp_path)}
    assert found == {"index.html", "res/deep/app.js"}
