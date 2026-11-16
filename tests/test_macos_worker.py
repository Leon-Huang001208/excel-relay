import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from excel_relay.core import macos_excel


class WindLogOwnershipTests(unittest.TestCase):
    def test_only_worker_open_log_can_supply_freshness(self):
        worker = Path("/tmp/excel_20261116_103823.log")
        foreground = Path("/tmp/excel_20261115_221147.log")
        apps = [{"pid": 7208}, {"pid": 24738}]
        with patch.object(macos_excel, "_wind_logs_for_pid", side_effect=lambda pid: {
            7208: {foreground}, 24738: {worker}
        }[pid], create=True):
            self.assertEqual(macos_excel._owned_wind_logs(24738, apps), {worker})

    def test_process_log_path_case_can_differ_from_vendor_symlink(self):
        with tempfile.TemporaryDirectory() as directory:
            canonical = Path(directory) / "Add-Ins"
            canonical.mkdir()
            alias = Path(directory) / "Add-ins"
            if not alias.exists():
                self.skipTest("test volume is case-sensitive")
            name = "excel_20261116_103823.log"
            response = subprocess.CompletedProcess([], 0, stdout=f"n{alias / name}\n", stderr="")
            with patch.object(macos_excel, "_WIND_LOG_DIR", canonical), patch.object(
                macos_excel.subprocess, "run", return_value=response
            ):
                self.assertEqual(macos_excel._wind_logs_for_pid(42), {canonical.resolve() / name})

    def test_shared_log_cannot_supply_freshness(self):
        shared = Path("/tmp/excel_20261116_103823.log")
        apps = [{"pid": 7208}, {"pid": 24738}]
        with patch.object(macos_excel, "_wind_logs_for_pid", return_value={shared}, create=True):
            with self.assertRaisesRegex(RuntimeError, "归属"):
                macos_excel._owned_wind_logs(24738, apps)

    def test_only_complete_appended_responses_in_worker_log_count(self):
        response = b"post_matrix_request response size=1 request: probe\n"
        with tempfile.TemporaryDirectory() as directory:
            worker = Path(directory) / "excel_20261116_103823.log"
            worker.write_bytes(response)
            initial = macos_excel._wind_log_offsets({worker})
            foreground = worker.with_name("excel_20261115_221147.log")
            foreground.write_bytes(response)
            self.assertEqual(macos_excel._wind_response_count(initial), 0)
            with worker.open("ab") as stream:
                stream.write(b"post_matrix_request request: probe\n" + response + response[:-1])
            self.assertEqual(macos_excel._wind_response_count(initial), 1)

    def test_wait_requires_an_additional_quiet_response_wave(self):
        counts = iter([2, 3, 4, 4, 4])
        ticks = iter([0.0, 0.0, 0.1, 0.2, 0.7, 1.1])
        with patch.object(macos_excel, "_wind_response_count", side_effect=lambda _: next(counts)), patch.object(
            macos_excel.time, "monotonic", side_effect=lambda: next(ticks)
        ), patch.object(macos_excel.time, "sleep"):
            self.assertEqual(macos_excel._wait_for_wind_wave({}, 2, 2.0), 4)

    def test_no_additional_response_fails_closed_at_shared_deadline(self):
        ticks = iter([0.0, 0.1, 0.3, 0.5])
        with patch.object(macos_excel, "_wind_response_count", return_value=1), patch.object(
            macos_excel.time, "monotonic", side_effect=lambda: next(ticks)
        ), patch.object(macos_excel.time, "sleep"):
            with self.assertRaisesRegex(RuntimeError, "Wind"):
                macos_excel._wait_for_wind_wave({}, 1, 0.5)


class WorkerRecoveryTests(unittest.TestCase):
    def test_malformed_or_uncertain_worker_records_block(self):
        with tempfile.TemporaryDirectory() as directory:
            marker = Path(directory) / "worker.json"
            for worker in (None, {"pid": 12}, {"pid": "12", "started": 1}):
                with self.subTest(worker=worker):
                    marker.write_text(json.dumps({"worker": worker, "path": "/tmp/staged.xlsx"}))
                    with self.assertRaises(macos_excel.ExcelBusyError):
                        macos_excel._check_previous_worker(marker)
                    self.assertTrue(marker.is_file())

    def test_reused_pid_does_not_match_finished_worker(self):
        with tempfile.TemporaryDirectory() as directory:
            marker = Path(directory) / "worker.json"
            marker.write_text(json.dumps({"worker": {"pid": 12, "started": 1}, "path": "/tmp/staged.xlsx"}))
            with patch.object(macos_excel, "_run_automation", return_value=[{"pid": 12, "started": 2}]):
                macos_excel._check_previous_worker(marker)
            self.assertFalse(marker.exists())

    @unittest.skipUnless(sys.platform == "darwin", "macOS legacy lock")
    def test_legacy_flock_stays_held_until_recovery_marker_is_cleared(self):
        import fcntl

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "excel.lock"
            marker = path.with_suffix(".worker.json")
            marker.write_text(json.dumps({"worker": None, "path": "/tmp/staged.xlsx"}))
            with path.open("wb") as lock:
                fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
                guard = macos_excel._retain_worker_lock(marker, lock.fileno())
            try:
                with path.open("rb") as contender:
                    with self.assertRaises(BlockingIOError):
                        fcntl.flock(contender, fcntl.LOCK_EX | fcntl.LOCK_NB)
            finally:
                marker.unlink()
                guard.wait(timeout=5)
            with path.open("rb") as contender:
                fcntl.flock(contender, fcntl.LOCK_EX | fcntl.LOCK_NB)


if __name__ == "__main__":
    unittest.main()
