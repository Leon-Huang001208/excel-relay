import hashlib
import tempfile
import unittest
from contextlib import contextmanager, nullcontext
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from excel_relay.core import macos_excel
from excel_relay.core.catalog import CascadingCatalog, CatalogManager
from excel_relay.core.refresher import HeadlessExcelRefresher, RefreshError


class MacRefreshTest(unittest.TestCase):
    def setUp(self):
        platform = patch("excel_relay.core.refresher.sys", SimpleNamespace(platform="darwin"))
        platform.start()
        self.addCleanup(platform.stop)
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.source = self.root / "底稿.xlsx"
        book = Workbook()
        book.active.title = "日期"
        book.active["A2"] = 46339
        book.create_sheet("海外")["C2"] = 0
        book.save(self.source)
        book.close()
        manager, _ = CatalogManager.init_project(self.root)
        entry = manager.register_path(self.source)[0]
        self.entry = manager.set_source_inputs(entry.source_id, ["截止日=日期!A2"])
        self.catalog = CascadingCatalog(manager)
        self.source_hash = hashlib.sha256(self.source.read_bytes()).digest()

    def test_busy_gate_is_before_input_preprocessing(self):
        busy_error = type("ExcelBusyError", (RuntimeError,), {})

        @contextmanager
        def busy():
            raise busy_error("后台 Excel 忙碌")
            yield

        with (
            patch.object(macos_excel, "ExcelBusyError", busy_error, create=True),
            patch.object(macos_excel, "excel_refresh_lock", busy),
            patch("excel_relay.core.refresher.ExcelInspector.refresh_input_values") as read,
            patch.object(macos_excel, "refresh_xlsx") as run,
            self.assertRaisesRegex(RefreshError, "忙碌") as raised,
        ):
            HeadlessExcelRefresher.refresh(self.entry.source_id, catalog_mgr=self.catalog)

        read.assert_not_called()
        run.assert_not_called()
        self.assertEqual(type(raised.exception).__name__, "RefreshBusyError")

    def test_reads_fresh_outputs_inside_lock_and_reuses_normalized_artifact(self):
        from excel_relay.core.contracts import configure_source_contract, read_contract_outputs

        configure_source_contract(
            self.catalog,
            self.entry.source_id,
            {
                "date_inputs": ["截止日"],
                "outputs": [{"name": "恒指周涨跌幅", "cell": "海外!C2", "unit": "%"}],
            },
        )
        active = False
        next_output = 0.87059619276008704

        @contextmanager
        def gate():
            nonlocal active
            self.assertFalse(active)
            active = True
            try:
                yield
            finally:
                active = False

        def run(path, resolved, timeout):
            self.assertTrue(active)
            self.assertEqual([(item.name, value) for item, value in resolved], [("截止日", 46325)])
            book = load_workbook(path)
            book["海外"]["C2"] = next_output
            book.save(path)
            book.close()

        def read(path, contract):
            self.assertTrue(active)
            self.assertNotEqual(path, self.source)
            return read_contract_outputs(path, contract)

        with (
            patch.object(macos_excel, "excel_refresh_lock", gate),
            patch.object(macos_excel, "refresh_xlsx", run),
            patch("excel_relay.core.contracts.read_contract_outputs", side_effect=read),
        ):
            path, _, artifact, outputs = HeadlessExcelRefresher.refresh_with_outputs(
                self.entry.source_id, {"截止日": "2026-10-30"}, catalog_mgr=self.catalog
            )
            self.assertAlmostEqual(outputs[0]["value"], next_output)
            self.assertEqual(artifact.params, {"截止日": 46325})
            next_output = 1.25
            same_path, _, newer, outputs = HeadlessExcelRefresher.refresh_with_outputs(
                self.entry.source_id, {"日期!A2": 46325}, catalog_mgr=self.catalog
            )
            self.assertEqual(same_path, path)
            self.assertEqual(newer.artifact_id, artifact.artifact_id)
            self.assertEqual(outputs[0]["value"], next_output)
            self.assertEqual(len(self.catalog.list_artifacts(self.entry.source_id)), 1)
            legacy = HeadlessExcelRefresher.refresh(
                self.entry.source_id, {"截止日": 46325}, catalog_mgr=self.catalog
            )
            self.assertEqual(len(legacy), 3)

        self.assertEqual(hashlib.sha256(self.source.read_bytes()).digest(), self.source_hash)

    def test_invalid_date_and_unregistered_parameter_do_not_launch_excel(self):
        from excel_relay.core.contracts import configure_source_contract

        configure_source_contract(self.catalog, self.entry.source_id, {"date_inputs": ["截止日"]})
        for params, message in (({"截止日": "2026-02-30"}, "日期"), ({"其他": "2026-10-30"}, "白名单")):
            with (
                self.subTest(params=params),
                patch.object(macos_excel, "excel_refresh_lock", nullcontext),
                patch.object(macos_excel, "refresh_xlsx") as run,
                self.assertRaisesRegex(RefreshError, message),
            ):
                HeadlessExcelRefresher.refresh(
                    self.entry.source_id, params, catalog_mgr=self.catalog
                )
            run.assert_not_called()

    def test_worker_cleanup_uncertainty_preserves_stage(self):
        cleanup_error = type("WorkerCleanupError", (RuntimeError,), {})
        target = self.root / "交付.xlsx"
        target.write_bytes(b"previous result")
        seen = []

        def uncertain_cleanup(path, resolved, timeout):
            seen.append(path)
            raise cleanup_error("后台 Excel 是否退出尚未确认")

        with (
            patch.object(macos_excel, "WorkerCleanupError", cleanup_error, create=True),
            patch.object(macos_excel, "excel_refresh_lock", nullcontext),
            patch.object(macos_excel, "refresh_xlsx", uncertain_cleanup),
            self.assertRaisesRegex(RefreshError, "尚未确认") as raised,
        ):
            HeadlessExcelRefresher.refresh(
                self.entry.source_id, output_path=target, catalog_mgr=self.catalog
            )

        self.assertTrue(seen[0].is_file())
        self.assertIn(str(seen[0]), str(raised.exception))
        self.assertEqual(target.read_bytes(), b"previous result")
        self.assertEqual(self.catalog.list_artifacts(self.entry.source_id), [])

    def test_confirmed_timeout_removes_stage_and_keeps_previous_result(self):
        target = self.root / "交付.xlsx"
        target.write_bytes(b"previous result")
        seen = []

        def timeout(path, resolved, timeout):
            seen.append(path)
            raise TimeoutError("后台 Excel 刷新超时")

        with (
            patch.object(macos_excel, "excel_refresh_lock", nullcontext),
            patch.object(macos_excel, "refresh_xlsx", timeout),
            self.assertRaisesRegex(RefreshError, "超时"),
        ):
            HeadlessExcelRefresher.refresh(
                self.entry.source_id, output_path=target, catalog_mgr=self.catalog
            )

        self.assertFalse(seen[0].exists())
        self.assertEqual(target.read_bytes(), b"previous result")


if __name__ == "__main__":
    unittest.main()
