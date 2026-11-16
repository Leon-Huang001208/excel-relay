import asyncio
import inspect
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import threading
import unittest
from unittest.mock import patch

from openpyxl import Workbook
from typer.testing import CliRunner

from excel_relay.core import CascadingCatalog, CatalogManager, HeadlessExcelRefresher
from excel_relay.core.models import RunArtifact
from excel_relay.interface import cli, mcp


class ContractFixture:
    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.manager, _ = CatalogManager.init_project(self.root)
        self.catalog = CascadingCatalog(self.manager)
        workbook = Workbook()
        workbook.active.title = "报表"
        workbook.active["A1"] = 46339
        workbook.active["B1"] = 0.870596
        path = self.root / "样例.xlsx"
        workbook.save(path)
        workbook.close()
        self.source = self.manager.register_path(path)[0].source_id
        self.catalog.set_source_inputs(self.source, ["截止日=报表!A1"])
        self.spec = {
            "date_inputs": ["截止日"],
            "outputs": [{"name": "涨幅", "cell": "报表!B1", "unit": "%"}],
        }
        self.values = [{**self.spec["outputs"][0], "sheet": "报表", "status": "ok", "value": 0.870596}]
        self.artifact = RunArtifact(
            artifact_id="result",
            source_id=self.source,
            file_path=str(self.root / "result.xlsx"),
            params={"截止日": 46325},
            params_hash="example",
            created_at="2026-11-16T09:00:00",
            keep_formulas=True,
        )


class CliInterfaceTests(ContractFixture, unittest.TestCase):
    def test_set_contract_registers_project_metrics(self):
        definition = self.root / "指标.json"
        definition.write_text(json.dumps(self.spec, ensure_ascii=False), encoding="utf-8")
        with patch.object(cli, "get_cascading_catalog", return_value=self.catalog):
            result = CliRunner().invoke(
                cli.app, ["set", self.source, "--contract", str(definition), "--json"]
            )
        self.assertEqual(result.exit_code, 0, result.output)
        payload = json.loads(result.output)
        self.assertEqual(payload["contract"], self.spec)
        self.assertTrue((self.manager.catalog_dir / "contracts.json").is_file())

    def test_run_json_returns_artifact_and_outputs(self):
        with (
            patch.object(cli, "get_cascading_catalog", return_value=self.catalog),
            patch.object(HeadlessExcelRefresher, "refresh", side_effect=AssertionError("legacy refresh")),
            patch.object(
                HeadlessExcelRefresher, "refresh_with_outputs", create=True,
                return_value=(self.artifact.file_path, 2.0, self.artifact, self.values),
            ) as refresh,
        ):
            result = CliRunner().invoke(
                cli.app, ["run", self.source, "--param", "截止日=2026-10-30", "--json"]
            )
        self.assertEqual(result.exit_code, 0, result.output)
        payload = json.loads(result.output)
        self.assertEqual(payload["result_file_path"], self.artifact.file_path)
        self.assertEqual(payload["outputs"], self.values)
        self.assertEqual(payload["refreshed_at"], self.artifact.created_at)
        self.assertEqual(refresh.call_args.kwargs["params"], {"截止日": "2026-10-30"})

    def test_run_busy_has_distinct_json_status(self):
        from excel_relay.core.refresher import RefreshBusyError

        with (
            patch.object(cli, "get_cascading_catalog", return_value=self.catalog),
            patch.object(
                HeadlessExcelRefresher, "refresh_with_outputs", create=True,
                side_effect=RefreshBusyError("已有刷新在运行"),
            ),
        ):
            result = CliRunner().invoke(cli.app, ["run", self.source, "--json"])
        self.assertNotEqual(result.exit_code, 0)
        self.assertEqual(json.loads(result.output)["status"], "busy")


class McpInterfaceTests(ContractFixture, unittest.IsolatedAsyncioTestCase):
    async def test_configure_contract_is_visible_in_detail(self):
        with patch.object(mcp, "get_cascading_catalog", return_value=self.catalog):
            content, _ = await mcp.mcp.call_tool(
                "configure_source_contract",
                {"source_id": self.source, **self.spec, "workspace_path": str(self.root)},
            )
            detail = json.loads(mcp.get_source_detail(self.source, workspace_path=str(self.root)))
        payload = json.loads(content[0].text)
        self.assertEqual(payload["contract"], self.spec)
        self.assertEqual(detail["contract"], self.spec)

    async def test_refresh_can_return_busy_during_another_request(self):
        self.assertTrue(inspect.iscoroutinefunction(mcp.refresh_data_source))
        from excel_relay.core.refresher import RefreshBusyError

        started = threading.Event()
        release = threading.Event()
        listener = threading.get_ident()
        catalog_threads = []

        def catalog_lookup(*_args, **_kwargs):
            catalog_threads.append(threading.get_ident())
            return self.catalog

        def refresh(**kwargs):
            if kwargs["params"]["截止日"] == 46325:
                started.set()
                if not release.wait(4):
                    raise RuntimeError("test refresh was not released")
                return self.artifact.file_path, 2.0, self.artifact, self.values
            raise RefreshBusyError("已有刷新在运行")

        with (
            patch.object(mcp, "get_cascading_catalog", side_effect=catalog_lookup),
            patch.object(HeadlessExcelRefresher, "refresh_with_outputs", create=True, side_effect=refresh),
        ):
            first = asyncio.create_task(mcp.refresh_data_source(self.source, {"截止日": 46325}))
            try:
                self.assertTrue(await asyncio.to_thread(started.wait, 2))
                busy = await asyncio.wait_for(
                    mcp.refresh_data_source(self.source, {"截止日": 46326}), timeout=1
                )
                self.assertEqual(json.loads(busy)["status"], "busy")
            finally:
                release.set()
                result = json.loads(await first)
        self.assertEqual(result["outputs"], self.values)
        self.assertEqual(result["refreshed_at"], self.artifact.created_at)
        self.assertEqual(len(catalog_threads), 2)
        self.assertTrue(all(thread != listener for thread in catalog_threads))

    async def test_refresh_without_contract_preserves_existing_result_shape(self):
        self.assertTrue(inspect.iscoroutinefunction(mcp.refresh_data_source))
        with (
            patch.object(mcp, "get_cascading_catalog", return_value=self.catalog),
            patch.object(
                HeadlessExcelRefresher, "refresh_with_outputs", create=True,
                return_value=(self.artifact.file_path, 2.0, self.artifact, None),
            ),
        ):
            result = json.loads(await mcp.refresh_data_source(self.source))
        self.assertEqual(result["artifact_id"], self.artifact.artifact_id)
        self.assertNotIn("outputs", result)
        self.assertNotIn("refreshed_at", result)
