import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.utils.datetime import MAC_EPOCH

from excel_relay.core.catalog import CascadingCatalog, CatalogManager
from excel_relay.core.contracts import (
    SourceContract,
    configure_source_contract,
    get_source_contract,
    normalize_contract_params,
    read_contract_outputs,
)
from excel_relay.core.refresher import HeadlessExcelRefresher, RefreshError


def make_workbook(path, *, mac_epoch=False):
    book = Workbook()
    book.active.title = "日期"
    if mac_epoch:
        book.epoch = MAC_EPOCH
    book.active["A2"] = 46339
    book.active["B2"] = "HSI.HI"
    results = book.create_sheet("海外")
    results["C2"] = "=0.87059619276008704"
    results["C3"] = "#DIV/0!"
    results["C4"].number_format = "0.00"
    results["D2"] = 0
    book.save(path)
    book.close()


def make_saved_copy(source, copy):
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with ZipFile(source) as original, ZipFile(copy, "w") as refreshed:
        for name in original.namelist():
            data = original.read(name)
            if name == "xl/worksheets/sheet2.xml":
                root = ElementTree.fromstring(data)
                cell = root.find(f".//{ns}c[@r='C2']")
                cell.find(f"{ns}v").text = "0.87059619276008704"
                refreshed.writestr(name, ElementTree.tostring(root))
            else:
                refreshed.writestr(name, data)


class SourceContractTest(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.workbook = self.root / "底稿.xlsx"
        make_workbook(self.workbook)
        manager, _ = CatalogManager.init_project(self.root)
        entry = manager.register_path(self.workbook)[0]
        self.entry = manager.set_source_inputs(entry.source_id, ["截止日=日期!A2", "标的=日期!B2"])
        self.catalog = CascadingCatalog(manager)
        self.spec = {
            "date_inputs": ["日期!$A$2"],
            "outputs": [{"name": "恒指周涨跌幅", "cell": "海外!$C$2", "unit": "%"}],
        }

    def configure(self):
        return configure_source_contract(self.catalog, self.entry.source_id, self.spec)

    def test_contract_is_project_sidecar_and_survives_catalog_updates(self):
        catalog_file = self.root / ".excel-relay/catalog.json"
        before = catalog_file.read_bytes()
        contract = self.configure()
        assert catalog_file.read_bytes() == before
        sidecar = self.root / ".excel-relay/contracts.json"
        self.assertEqual(json.loads(sidecar.read_text())[self.entry.source_id], contract.model_dump())
        self.assertEqual(contract.date_inputs, ["截止日"])
        self.assertEqual(contract.outputs[0].cell, "海外!C2")
        older_catalog = CatalogManager(base_dir=self.root)
        older_catalog.register_path(self.workbook)
        self.assertEqual(get_source_contract(self.catalog, self.entry), contract)
        self.assertEqual(json.loads(sidecar.read_text())[self.entry.source_id], contract.model_dump())

    def test_second_registration_keeps_other_source_contracts(self):
        first = self.configure()
        other = self.root / "另一个底稿.xlsx"
        make_workbook(other)
        other_entry = self.catalog.primary.register_path(other)[0]
        configure_source_contract(self.catalog, other_entry.source_id, {"outputs": self.spec["outputs"]})
        self.assertEqual(get_source_contract(self.catalog, self.entry), first)

    def test_rejects_contract_for_unregistered_source(self):
        with self.assertRaises(ValueError):
            configure_source_contract(self.catalog, "unknown", self.spec)

    def test_rejects_missing_date_input_or_nonqualified_output(self):
        for update in (
            {"date_inputs": ["未登记"], "outputs": self.spec["outputs"]},
            {"date_inputs": [], "outputs": [{"name": "收益", "cell": "A1", "unit": "%"}]},
            {"date_inputs": [], "outputs": [{"name": "收益", "cell": "海外!A1:B2", "unit": "%"}]},
            {"date_inputs": [], "outputs": [{"name": "收益", "cell": "海外!XFE1048576", "unit": "%"}]},
        ):
            with self.subTest(update=update), self.assertRaises(ValueError):
                configure_source_contract(self.catalog, self.entry.source_id, update)

    def test_registered_date_is_normalized_in_both_excel_epochs(self):
        contract = self.configure()
        mac_book = self.root / "mac.xlsx"
        make_workbook(mac_book, mac_epoch=True)
        for path, expected in ((self.workbook, 46325), (mac_book, 44863)):
            with self.subTest(path=path):
                for name in ("截止日", "日期!$A$2"):
                    params = {name: "2026-10-30", "标的": "HSI.HI"}
                    converted = normalize_contract_params(path, self.entry, params, contract)
                    self.assertEqual(converted, {name: expected, "标的": "HSI.HI"})
                    self.assertEqual(params[name], "2026-10-30")
                    self.assertEqual(normalize_contract_params(path, self.entry, {name: expected}, contract),
                                     {name: expected})

    def test_invalid_date_stays_an_error_and_unknown_parameter_stays_unregistered(self):
        contract = self.configure()
        for invalid in ("2026-02-30", "2026-1-01", True):
            with self.subTest(invalid=invalid), self.assertRaisesRegex(ValueError, "YYYY-MM-DD"):
                normalize_contract_params(self.workbook, self.entry, {"截止日": invalid}, contract)
        params = normalize_contract_params(self.workbook, self.entry, {"额外输入": "2026-10-30"}, contract)
        self.assertEqual(params, {"额外输入": "2026-10-30"})
        with self.assertRaises(RefreshError):
            HeadlessExcelRefresher.validate_and_resolve_params(params, self.entry.inputs)

    def test_outputs_use_exact_cached_value_and_report_per_cell_failures(self):
        source_hash = hashlib.sha256(self.workbook.read_bytes()).digest()
        saved = self.root / "本次结果.xlsx"
        make_saved_copy(self.workbook, saved)
        contract = SourceContract.model_validate({
            "outputs": [
                {"name": "收益", "cell": "海外!C2", "unit": "%"},
                {"name": "报错", "cell": "海外!C3", "unit": ""},
                {"name": "空白", "cell": "海外!C4", "unit": ""},
                {"name": "越界", "cell": "海外!C8", "unit": ""},
                {"name": "缺页", "cell": "已删除!A1", "unit": ""},
                {"name": "零", "cell": "海外!D2", "unit": ""},
            ],
        })
        outcomes = read_contract_outputs(saved, contract)
        self.assertEqual(outcomes[0], {"name": "收益", "cell": "海外!C2", "unit": "%",
                                       "status": "ok", "value": 0.87059619276008704})
        for result, reason in zip(outcomes[1:5], ("#DIV/0!", "空白", "缺失", "工作表")):
            self.assertEqual(result["status"], "error")
            self.assertIn(reason, result["error"])
        self.assertEqual(outcomes[-1]["value"], 0)
        self.assertEqual(hashlib.sha256(self.workbook.read_bytes()).digest(), source_hash)
