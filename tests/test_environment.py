import contextlib
import io
import json
import os
from pathlib import Path
import runpy
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd
from dotenv import dotenv_values

from src.config import PROJECT_ROOT, env_choice, env_mapping, env_path, env_seconds, env_text, mapped_file


class ConfigurationTests(unittest.TestCase):
    def test_required_and_optional_values(self):
        with patch.dict(os.environ, {"EMPTY": ""}, clear=True):
            for name in ("MISSING", "EMPTY"):
                with self.subTest(name=name), self.assertRaisesRegex(ValueError, name):
                    env_text(name)
            self.assertEqual(env_text("EMPTY", allow_empty=True), "")

    def test_site_validation(self):
        with patch.dict(os.environ, {"SITE": "BA"}, clear=True):
            self.assertEqual(env_choice("SITE", ("BA", "RS")), "BA")
            with self.assertRaisesRegex(ValueError, "SITE"):
                env_choice("SITE", ("AL",))

    def test_mapping_validation(self):
        for value in ("invalid", "[]", "{}", '{"BA": 1}', '{"": "file.xlsx"}'):
            with self.subTest(value=value), patch.dict(os.environ, {"FILES": value}, clear=True):
                with self.assertRaisesRegex(ValueError, "FILES"):
                    env_mapping("FILES")
        with patch.dict(os.environ, {"FILES": '{"BA":"control.xlsx"}'}, clear=True):
            self.assertEqual(mapped_file("FILES", "BA"), "control.xlsx")
            with self.assertRaisesRegex(ValueError, "RS"):
                mapped_file("FILES", "RS")

    def test_wait_validation(self):
        for value in ("-1", "nan", "inf", "text"):
            with self.subTest(value=value), patch.dict(os.environ, {"WAIT": value}, clear=True):
                with self.assertRaisesRegex(ValueError, "WAIT"):
                    env_seconds("WAIT")
        with patch.dict(os.environ, {"WAIT": "0"}, clear=True):
            self.assertEqual(env_seconds("WAIT"), 0)

    def test_relative_paths_and_templates(self):
        with patch.dict(os.environ, {"FILE": "outputs\\{site}-{date}.xlsx"}, clear=True):
            self.assertEqual(
                env_path("FILE", site="BA", date="2026-01-01"),
                str(PROJECT_ROOT / "outputs" / "BA-2026-01-01.xlsx"),
            )
        with patch.dict(os.environ, {"FILE": "{unknown}.xlsx"}, clear=True):
            with self.assertRaisesRegex(ValueError, "FILE"):
                env_text("FILE", site="BA")

    def test_dotenv_loading_from_other_working_directory_and_precedence(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "src").mkdir()
            (root / "src" / "config.py").write_text(
                (PROJECT_ROOT / "src" / "config.py").read_text(encoding="utf-8"), encoding="utf-8",
            )
            (root / ".env").write_text(
                (PROJECT_ROOT / ".env.example").read_text(encoding="utf-8"), encoding="utf-8",
            )
            previous = Path.cwd()
            try:
                os.chdir(root / "src")
                with patch.dict(os.environ, {"PERFIS_RS_SITE": "AC"}, clear=True):
                    runpy.run_path(str(root / "src" / "config.py"))
                    self.assertEqual(os.environ["PERFIS_RS_SITE"], "AC")
                    self.assertEqual(os.environ["PERFIS_RS_LOCATION_COLUMN"], "ID Automação")
            finally:
                os.chdir(previous)

    def test_template_has_optional_subarea(self):
        example = dotenv_values(PROJECT_ROOT / ".env.example")
        self.assertEqual(example["PERFIS_RS_SUBAREA_COLUMN"], "")


class ScriptTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.values = dict(dotenv_values(PROJECT_ROOT / ".env.example"))

    def execute(self, script, overrides):
        with patch.dict(os.environ, {**self.values, **overrides}, clear=True), \
                patch("webbrowser.open"), patch("time.sleep"), \
                patch("builtins.input", return_value=""), \
                patch("pyshorteners.Shortener") as shortener, \
                contextlib.redirect_stdout(io.StringIO()):
            shortener.return_value.tinyurl.short.return_value = "https://example.com/short"
            namespace = runpy.run_module(f"src.{Path(script).stem}", run_name="__main__")
            for value in namespace.values():
                if isinstance(value, pd.ExcelFile):
                    value.close()
            return shortener.return_value.tinyurl.short.call_args_list

    def test_pa_check_custom_paths_and_columns(self):
        yael = self.root / "yael.xlsx"
        taise = self.root / "taise.xlsx"
        output = self.root / "comparison.xlsx"
        pd.DataFrame({"Plant": ["A"], "Point": ["P1"], "Chemical": ["Arsênio"]}).to_excel(yael, index=False)
        pd.DataFrame({"Plant": ["A"], "Point": ["P1"], "Chemical": ["Chumbo"]}).to_excel(taise, index=False)
        self.execute("pa_check.py", {
            "PA_CHECK_YAEL_FILE": str(yael), "PA_CHECK_TAISE_FILE": str(taise),
            "PA_CHECK_OUTPUT_FILE": str(output),
            "PA_CHECK_SITE_COLUMN": "Plant", "PA_CHECK_POINT_COLUMN": "Point",
            "PA_CHECK_PARAMETER_COLUMN": "Chemical",
        })
        result = pd.read_excel(output, sheet_name="PA Yael")
        self.assertEqual(result.loc[0, "key"], "A-P1-Arsênio")
        self.assertIn("não foram solicitados", result.loc[0, "Parâmetros ausentes"])

    def test_perfis_rs_all_sites_and_qc(self):
        for site in ("BA", "AL", "AC", "Tupy", "RS"):
            with self.subTest(site=site):
                if site == "BA":
                    path, plant = "https://example.com/PE1/", "PE-1"
                elif site == "AL":
                    path, plant = "https://example.com/Perfil Construtivo/Plant/", "Plant"
                elif site == "AC":
                    path, plant = "https://example.com/Custom AC/01. REFMAT/", "REFMAT"
                elif site == "RS":
                    path, plant = "https://example.com/Custom RS/Plant/Sub/", "Plant"
                else:
                    path, plant = "https://example.com/", "example.com"
                control = self.root / f"{site}-control.xlsx"
                links = self.root / f"{site}-links.xlsx"
                output = self.root / f"{site}-out.xlsx"
                qc = self.root / f"{site}-qc.xlsx"
                pd.DataFrame({
                    "Location": ["P1", "MISSING"], "Plant": [plant, plant],
                    "Area": ["Sub", "Sub"],
                }).to_excel(control, index=False)
                pd.DataFrame({
                    "Filename": ["P1.pdf", "EXTRA.pdf"], "Path": [path, path],
                }).to_excel(links, index=False)
                calls = self.execute("perfis_rs.py", {
                    "PERFIS_RS_SITE": site, "PERFIS_RS_INPUT_DIR": str(self.root),
                    "PERFIS_RS_OUTPUT_DIR": str(self.root),
                    "PERFIS_RS_CONTROL_FILES": json.dumps({site: control.name}),
                    "PERFIS_RS_LINKS_FILE": "{site}-links.xlsx",
                    "PERFIS_RS_LOCATION_COLUMN": "Location",
                    "PERFIS_RS_SITE_COLUMN": "Plant", "PERFIS_RS_SUBAREA_COLUMN": "Area",
                    "PERFIS_RS_LINK_NAME_COLUMN": "Filename", "PERFIS_RS_LINK_PATH_COLUMN": "Path",
                    "PERFIS_RS_RS_PATH_MARKER": "Custom RS",
                    "PERFIS_RS_AC_PATH_MARKER": "Custom AC/",
                    "PERFIS_RS_AL_PATH_MARKER": "Perfil Construtivo/",
                    "PERFIS_RS_OUTPUT_FILE": "{site}-out.xlsx", "PERFIS_RS_QC_FILE": "{site}-qc.xlsx",
                })
                self.assertEqual(len(calls), 1)
                self.assertEqual(calls[0].args[0], path.replace(" ", "%20") + "P1.pdf")
                result = pd.read_excel(output)
                self.assertEqual(len(result), 2)
                self.assertEqual(result.loc[0, "link"], '<a href="https://example.com/short">Disponível</a>')
                self.assertTrue(pd.isna(result.loc[1, "link"]))
                self.assertEqual(pd.read_excel(qc, sheet_name=0).loc[0, "Location"], "MISSING")
                self.assertEqual(pd.read_excel(qc, sheet_name=1).loc[0, "Name"], "EXTRA.pdf")

    def test_rs_requires_subarea_before_reading_files(self):
        with self.assertRaisesRegex(ValueError, "PERFIS_RS_SUBAREA_COLUMN"):
            self.execute("perfis_rs.py", {"PERFIS_RS_SITE": "RS", "PERFIS_RS_SUBAREA_COLUMN": ""})

    def test_configurable_retry_intervals(self):
        pd.DataFrame({
            "ID Automação": ["P1"], "Site": ["PE-1"],
        }).to_excel(self.root / "control.xlsx", index=False)
        pd.DataFrame({
            "Name": ["P1.pdf"], "Folder Path": ["https://example.com/PE1/"],
        }).to_excel(self.root / "links.xlsx", index=False)
        with patch("time.sleep") as sleep:
            with patch.dict(os.environ, {
                **self.values, "PERFIS_RS_INPUT_DIR": str(self.root),
                "PERFIS_RS_OUTPUT_DIR": str(self.root),
                "PERFIS_RS_CONTROL_FILES": '{"BA":"control.xlsx"}',
                "PERFIS_RS_LINKS_FILE": "links.xlsx",
                "LINK_RETRY_WAIT_1": "0.25", "LINK_RETRY_WAIT_2": "0.5",
            }, clear=True), patch("webbrowser.open"), \
                    patch("pyshorteners.Shortener") as shortener, \
                    contextlib.redirect_stdout(io.StringIO()):
                shortener.return_value.tinyurl.short.side_effect = [
                    ConnectionError("first"), ConnectionError("second"), "https://example.com/short",
                ]
                runpy.run_module("src.perfis_rs", run_name="__main__")
            self.assertEqual([call.args[0] for call in sleep.call_args_list], [0.25, 0.5])

    def test_perfis_custom_columns_and_fallback(self):
        control = self.root / "control.xlsx"
        links = self.root / "links.xlsx"
        pd.DataFrame({
            "Location": ["P1"], "Plant": ["AL"], "Global": ["g1"], "Object": [1],
        }).to_excel(control, index=False)
        pd.DataFrame({
            "Filename": ["P1.pdf"], "Path": ["https://example.com/AL/"],
        }).to_excel(links, index=False)
        calls = self.execute("perfis.py", {
            "PERFIS_SITE": "AL", "PERFIS_INPUT_DIR": str(self.root), "PERFIS_OUTPUT_DIR": str(self.root),
            "PERFIS_CONTROL_FILES": '{"AL":"control.xlsx"}', "PERFIS_LINKS_FILE": "links.xlsx",
            "PERFIS_FILE_COLUMN": "File", "PERFIS_LOCATION_COLUMN": "Location",
            "PERFIS_SITE_COLUMN": "Plant", "PERFIS_GLOBAL_ID_COLUMN": "Global",
            "PERFIS_OBJECT_ID_COLUMN": "Object", "PERFIS_LINK_NAME_COLUMN": "Filename",
            "PERFIS_LINK_PATH_COLUMN": "Path", "PERFIS_QC_FILE": "qc.xlsx",
            "PERFIS_OUTPUT_FILE": "out.xlsx",
        })
        self.assertEqual(len(calls), 1)
        self.assertEqual(pd.read_excel(self.root / "out.xlsx").loc[0, "SYS_LOC_CODE"], "P1")
        self.assertTrue(pd.read_excel(self.root / "qc.xlsx").empty)

    def test_main_custom_files_sheets_and_plan_columns(self):
        plan = self.root / "plan.xlsx"
        edd = self.root / "edd.xlsx"
        checklist = self.root / "checklist.xlsx"
        with pd.ExcelWriter(plan) as writer:
            pd.DataFrame({
                "Group": ["Metais"], "Points": ["P1,P2"], "Parameters": ["Arsênio,Chumbo"],
            }).to_excel(writer, sheet_name="PE-1_PA", index=False)
        with pd.ExcelWriter(edd) as writer:
            pd.DataFrame({
                "sys_sample_code": ["S1", "S3"], "sys_loc_code": ["P1", "P3"],
                "sample_date": ["2026-01-01", "2026-01-01"],
            }).to_excel(writer, sheet_name="Samples", index=False)
            pd.DataFrame({
                "sys_sample_code": ["S1", "S3"], "chemical_name": ["Arsênio", "Chumbo"],
            }).to_excel(writer, sheet_name="Results", index=False)
            for name in ("SMP", "Tests", "Batch"):
                pd.DataFrame({"unused": []}).to_excel(writer, sheet_name=name, index=False)
        pd.DataFrame([
            ["info", "info", "info", "info"],
            ["info", "info", "info", "info"],
            ["info", "info", "info", "info"],
            ["Poço", "Data", "Poço foi amostrado?", "Status do Poço"],
            ["P1", "2026-01-01", "Sim", "OK"],
            ["P2", "2026-01-01", "Não", "Seco"],
        ]).to_excel(checklist, index=False)
        (self.root / "checklists").mkdir()
        checklist.rename(self.root / "checklists" / checklist.name)
        self.execute("main.py", {
            "MAIN_PLAN_FILE": str(plan), "MAIN_EQUIS_DIR": str(self.root),
            "MAIN_CHECKLIST_DIR": str(self.root / "checklists"),
            "MAIN_PLAN_OUTPUT_FILE": str(self.root / "plan-out.xlsx"),
            "MAIN_REPORT_OUTPUT_FILE": str(self.root / "report.xlsx"),
            "MAIN_EQUIS_FILES": '{"PE-1_PA":"edd.xlsx"}',
            "MAIN_CHECKLIST_FILES": '{"PE-1_PA":"checklist.xlsx"}',
            "MAIN_GROUP_COLUMN": "Group", "MAIN_POINTS_COLUMN": "Points",
            "MAIN_PARAMETERS_COLUMN": "Parameters",
            "MAIN_SAMPLE_SHEET": "Samples", "MAIN_SMP_SHEET": "SMP",
            "MAIN_TEST_SHEET": "Tests", "MAIN_RESULT_SHEET": "Results", "MAIN_BATCH_SHEET": "Batch",
        })
        self.assertEqual(len(pd.read_excel(self.root / "plan-out.xlsx")), 4)
        missing = pd.read_excel(self.root / "report.xlsx", sheet_name="Análises ausentes")
        self.assertIn("Chumbo", missing.loc[missing["Ponto"].eq("P1"), "Parâmetro"].tolist())
        self.assertEqual(set(missing.loc[missing["Ponto"].eq("P2"), "Motivo"]), {"Seco"})


if __name__ == "__main__":
    unittest.main()
