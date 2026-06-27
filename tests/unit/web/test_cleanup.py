import os
import time
from datetime import timedelta
from pathlib import Path

import pytest

from xlsliberator.web.cleanup import CleanupSafetyError, cleanup_old_jobs


def test_cleanup_deletes_only_old_job_dirs(tmp_path: Path) -> None:
    old = tmp_path / "jobs" / "old"
    new = tmp_path / "jobs" / "new"
    old.mkdir(parents=True)
    new.mkdir()
    old_time = time.time() - 72 * 3600
    os.utime(old, (old_time, old_time))

    deleted = cleanup_old_jobs(tmp_path, timedelta(hours=24))

    assert deleted == [old]
    assert not old.exists()
    assert new.exists()


@pytest.mark.parametrize("path", [Path("/"), Path.home(), Path.cwd()])
def test_cleanup_refuses_unsafe_dirs(path: Path) -> None:
    with pytest.raises(CleanupSafetyError):
        cleanup_old_jobs(path, timedelta(hours=24))
