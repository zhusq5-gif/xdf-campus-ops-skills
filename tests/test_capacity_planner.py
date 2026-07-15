from __future__ import annotations

import importlib.util
import json
import tempfile
import unittest
import zipfile
from datetime import date
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / ".agents/skills/xdf-plan-campus-capacity/scripts/plan_capacity.py"
SPEC = importlib.util.spec_from_file_location("plan_capacity", SCRIPT)
assert SPEC and SPEC.loader
PLAN_CAPACITY = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PLAN_CAPACITY)

INPUT = ROOT / "evals/fixtures/synthetic-capacity-input.xlsx"
CONFIG = ROOT / "config/capacity-planning.json"
TEMPLATE = ROOT / ".agents/skills/xdf-plan-campus-capacity/assets/capacity-output-template.xlsx"


class CapacityPlannerTest(unittest.TestCase):
    def run_case(
        self,
        input_path: Path,
        output_dir: Path,
        *,
        backtest: bool = True,
        reference_date: date | None = None,
    ):
        return PLAN_CAPACITY.run_planning(
            input_path,
            CONFIG,
            TEMPLATE,
            output_dir,
            backtest=backtest,
            reference_date=reference_date,
        )

    def mutate_input(self, folder: Path, name: str, mutator) -> Path:
        target = folder / name
        workbook = load_workbook(INPUT)
        mutator(workbook)
        workbook.save(target)
        workbook.close()
        return target

    def test_golden_capacity_numbers_and_human_confirmation(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "result"
            result = self.run_case(INPUT, output)

            self.assertEqual("ready", result["status"])
            a_row = next(
                row
                for row in result["forecast"]
                if row["scenario"] == "基准" and row["class_type"] == "素养A" and row["month"] == "2026-09"
            )
            self.assertEqual(24, a_row["projected_students"])
            self.assertEqual(3, a_row["planned_classes"])

            b_teacher = next(
                row
                for row in result["teacher_gaps"]
                if row["scenario"] == "基准" and row["class_type"] == "素养B" and row["month"] == "2026-10"
            )
            b_room = next(
                row
                for row in result["room_gaps"]
                if row["scenario"] == "基准" and row["room_type"] == "活动教室" and row["month"] == "2026-10"
            )
            self.assertEqual(4, b_teacher["gap"])
            self.assertEqual(2, b_room["gap"])

            payload = json.loads((output / "message-payload.json").read_text(encoding="utf-8"))
            self.assertTrue(payload["requires_human_confirmation"])
            self.assertFalse(payload["auto_send"])
            self.assertTrue(payload["actions"])
            self.assertTrue(all(action["evidence_keys"] for action in payload["actions"]))

    def test_same_input_and_config_produce_same_payload(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.run_case(INPUT, root / "one")
            self.run_case(INPUT, root / "two")
            first = (root / "one/message-payload.json").read_bytes()
            second = (root / "two/message-payload.json").read_bytes()
            self.assertEqual(first, second)

    def test_class_boundaries_determine_class_count(self):
        rule = {"min_class_size": 6, "target_class_size": 10, "max_class_size": 12}
        self.assertEqual((0, 0.0, "低于开班下限"), PLAN_CAPACITY.plan_classes(5, rule))
        self.assertEqual((1, 11.0, "可开班"), PLAN_CAPACITY.plan_classes(11, rule))
        self.assertEqual((2, 6.5, "可开班"), PLAN_CAPACITY.plan_classes(13, rule))
        self.assertEqual((3, 8.0, "可开班"), PLAN_CAPACITY.plan_classes(24, rule))

    def test_missing_column_blocks_recommendations(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = self.mutate_input(root, "missing-column.xlsx", lambda wb: setattr(wb["在读学员"]["A1"], "value", "未知列"))
            result = self.run_case(path, root / "out")
            self.assertEqual("blocked", result["status"])
            self.assertIn("MISSING_COLUMN", {issue["code"] for issue in result["issues"]})
            self.assertEqual([], result["recommendations"])

    def test_duplicate_student_blocks_recommendations(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)

            def duplicate(workbook):
                sheet = workbook["在读学员"]
                sheet.append([cell.value for cell in sheet[2]])

            path = self.mutate_input(root, "duplicate-student.xlsx", duplicate)
            result = self.run_case(path, root / "out")
            self.assertEqual("blocked", result["status"])
            self.assertIn("DUPLICATE_STUDENT", {issue["code"] for issue in result["issues"]})
            self.assertEqual([], result["recommendations"])

    def test_out_of_range_renewal_rate_blocks_recommendations(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = self.mutate_input(root, "bad-rate.xlsx", lambda wb: setattr(wb["续费率"]["D2"], "value", 1.2))
            result = self.run_case(path, root / "out")
            self.assertEqual("blocked", result["status"])
            self.assertIn("RATE_OUT_OF_RANGE", {issue["code"] for issue in result["issues"]})
            self.assertEqual([], result["recommendations"])

    def test_excel_formula_blocks_recommendations(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = self.mutate_input(root, "formula.xlsx", lambda wb: setattr(wb["在读学员"]["A2"], "value", "=1+1"))
            result = self.run_case(path, root / "out")
            self.assertEqual("blocked", result["status"])
            self.assertTrue({"EXCEL_FORMULA", "FORMULA_INJECTION"} & {issue["code"] for issue in result["issues"]})
            self.assertEqual([], result["recommendations"])

    def test_teacher_and_room_time_conflicts_block(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)

            def conflict(workbook):
                teacher = workbook["教师供给"]
                teacher.append([cell.value for cell in teacher[2]])
                room = workbook["教室供给"]
                room.append([cell.value for cell in room[2]])

            path = self.mutate_input(root, "conflicts.xlsx", conflict)
            result = self.run_case(path, root / "out")
            codes = {issue["code"] for issue in result["issues"]}
            self.assertEqual("blocked", result["status"])
            self.assertIn("TEACHER_TIME_CONFLICT", codes)
            self.assertIn("ROOM_TIME_CONFLICT", codes)
            self.assertEqual([], result["recommendations"])

    def test_stale_data_blocks_live_decision(self):
        with tempfile.TemporaryDirectory() as directory:
            result = self.run_case(
                INPUT,
                Path(directory) / "out",
                backtest=False,
                reference_date=date(2026, 10, 1),
            )
            self.assertEqual("blocked", result["status"])
            self.assertIn("STALE_DATA", {issue["code"] for issue in result["issues"]})
            self.assertEqual([], result["recommendations"])

    def test_start_month_before_recruitment_month_blocks(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = self.mutate_input(root, "date-order.xlsx", lambda wb: setattr(wb["招生计划"]["G2"], "value", "2026-07"))
            result = self.run_case(path, root / "out")
            self.assertEqual("blocked", result["status"])
            self.assertIn("DATE_ORDER_CONFLICT", {issue["code"] for issue in result["issues"]})
            self.assertEqual([], result["recommendations"])

    def test_direct_identifier_blocks(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            direct_identifier = "13" + "800000000"
            path = self.mutate_input(root, "direct-identifier.xlsx", lambda wb: setattr(wb["在读学员"]["A2"], "value", direct_identifier))
            result = self.run_case(path, root / "out")
            self.assertEqual("blocked", result["status"])
            self.assertIn("DIRECT_IDENTIFIER", {issue["code"] for issue in result["issues"]})
            self.assertEqual([], result["recommendations"])

    def test_external_link_part_blocks(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root / "external-link.xlsx"
            with zipfile.ZipFile(INPUT) as source, zipfile.ZipFile(target, "w") as destination:
                for item in source.infolist():
                    destination.writestr(item, source.read(item.filename))
                destination.writestr("xl/externalLinks/externalLink1.xml", "<externalLink/>")
            result = self.run_case(target, root / "out")
            self.assertEqual("blocked", result["status"])
            self.assertIn("EXTERNAL_LINK", {issue["code"] for issue in result["issues"]})
            self.assertEqual([], result["recommendations"])


if __name__ == "__main__":
    unittest.main()
