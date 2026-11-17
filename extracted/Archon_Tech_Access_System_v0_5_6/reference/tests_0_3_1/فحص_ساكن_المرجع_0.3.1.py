#!/usr/bin/env python3
"""Static packaging and business-rule smoke checks. This does NOT compile or run Access/VBA."""
from __future__ import annotations
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "access_0_3_1" / "Archon_Access_Full_Package"
SRC = ROOT / "src"
BUILD = (SRC / "modArchonBuild.bas").read_text(encoding="utf-8")
FINANCE = (SRC / "modArchonFinance.bas").read_text(encoding="utf-8")
UI = (SRC / "modArchonUI.bas").read_text(encoding="utf-8")
REPORTS = (SRC / "modArchonReports.bas").read_text(encoding="utf-8")


def schema_fields() -> dict[str, set[str]]:
    fields: dict[str, set[str]] = {}
    for table, ddl in re.findall(r'"CREATE TABLE (\w+) \((.*?)\)"', BUILD):
        names: set[str] = set()
        for part in ddl.split(","):
            m = re.match(r"\s*(\w+)\s+", part)
            if m:
                names.add(m.group(1).lower())
        fields[table.lower()] = names
    return fields


def report_query_strings() -> list[tuple[str, str]]:
    lines = REPORTS.splitlines()
    result: list[tuple[str, str]] = []
    i = 0
    while i < len(lines):
        match = re.search(r'AT_ReplaceQuery\s+db,\s*"([^"]+)"', lines[i])
        if not match:
            i += 1
            continue
        name = match.group(1)
        chunks: list[str] = []
        j = i
        while j < len(lines):
            if j > i and (not lines[j].strip() or re.search(r'AT_ReplaceQuery\s+db,\s*"', lines[j])):
                break
            literals = re.findall(r'"((?:[^"]|"")*)"', lines[j])
            if j == i:
                literals = literals[1:]  # Skip the query name.
            chunks.extend(literals)
            j += 1
        sql = "".join(part.replace('""', '"') for part in chunks)
        result.append((name, sql))
        i = j
    return result


class SourceValidation(unittest.TestCase):
    def test_vba_modules_have_balanced_procedure_counts(self):
        for path in SRC.glob("*.bas"):
            text = path.read_text(encoding="utf-8")
            starts = len(re.findall(r"(?im)^\s*(?:(?:Public|Private)\s+)?(?:Sub|Function)\s+\w+", text))
            ends = len(re.findall(r"(?im)^\s*End\s+(?:Sub|Function)\s*$", text))
            self.assertEqual(starts, ends, f"Procedure block mismatch in {path.name}")
            self.assertIn("Option Explicit", text)

    def test_schema_relation_targets_exist(self):
        tables = schema_fields()
        self.assertGreaterEqual(len(tables), 35)
        links = re.findall(r'"(rel[^|]+)\|([^|]+)\|([^|]+)\|([^|]+)\|([^"|]+)"', BUILD)
        self.assertGreaterEqual(len(links), 40)
        for rel, child, child_field, parent, parent_field in links:
            self.assertIn(child.lower(), tables, f"{rel}: missing child table {child}")
            self.assertIn(parent.lower(), tables, f"{rel}: missing parent table {parent}")
            self.assertIn(child_field.lower(), tables[child.lower()], f"{rel}: missing child field {child_field}")
            self.assertIn(parent_field.lower(), tables[parent.lower()], f"{rel}: missing parent field {parent_field}")

    def test_permission_codes_seeded_and_role_granted(self):
        codes_block = re.search(r'codes = Array\((.*?)\)\n\s*names = Array\((.*?)\)', BUILD, re.S)
        self.assertIsNotNone(codes_block)
        codes = re.findall(r'"([A-Z_]+)"', codes_block.group(1))
        names = re.findall(r'"([^"\n]+)"', codes_block.group(2))
        self.assertEqual(len(codes), len(names))
        role_block = re.search(r'For Each permCode In Array\((.*?)\)\n', BUILD, re.S)
        self.assertIsNotNone(role_block)
        role_codes = re.findall(r'"([A-Z_]+)"', role_block.group(1))
        self.assertEqual(set(codes), set(role_codes))

    def test_sensitive_approvals_have_permission_checks(self):
        functions = ["AT_ApproveOwner", "AT_ApproveCharge", "AT_ApproveReceipt", "AT_ApprovePayment", "AT_ApproveTransfer", "AT_ConfirmTransfer", "AT_ApproveChargeAdjustment", "AT_ApplyOwnerCreditToCharge", "AT_ReverseReceiptAllocation"]
        for name in functions:
            match = re.search(rf"(?ims)^Public\s+(?:Sub|Function)\s+{name}\b(.*?)(?:^End Sub|^End Function)", FINANCE)
            self.assertIsNotNone(match, f"Missing procedure {name}")
            self.assertIn("AT_RequirePermission", match.group(1), f"No permission check in {name}")

    def test_number_helper_does_not_open_nested_transaction(self):
        helper = re.search(r"(?ims)^Private Function AT_NextSeriesValueInTransaction\b(.*?)(?:^End Function)", FINANCE)
        self.assertIsNotNone(helper)
        self.assertNotRegex(helper.group(1), r"(?i)BeginTrans|CommitTrans|Rollback")

    def test_number_series_format_never_truncates_overflow_digits(self):
        helper = re.search(r"(?ims)^Private Function AT_FormatSeriesNumber\b(.*?)(?:^End Function)", FINANCE)
        self.assertIsNotNone(helper)
        self.assertIn("If Len(digits) < padLength Then", helper.group(1))
        self.assertIn("AT_FormatSeriesNumber(prefix, n, pad)", FINANCE)
        self.assertNotRegex(FINANCE, r'Right\$\(String\$\(pad,\s*"0"\)\s*&\s*CStr\(n\),\s*pad\)')
        def fmt(prefix: str, value: int, width: int) -> str:
            digits = str(value).rjust(width, "0")
            return prefix + digits
        self.assertEqual(fmt("RC-", 1, 6), "RC-000001")
        self.assertEqual(fmt("RC-", 1_000_000, 6), "RC-1000000")

    def test_custom_at_references_resolve(self):
        modules = list(SRC.glob("*.bas"))
        joined = "\n".join(p.read_text(encoding="utf-8") for p in modules)
        definitions = set(re.findall(r"(?im)^\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(AT_\w+)", joined))
        constants = set(re.findall(r"(?im)^\s*Private Const (AT_\w+)(?:\s+As\s+\w+)?\s*=", joined))
        references = set(re.findall(r"\b(AT_\w+)\b", joined))
        self.assertEqual(references - definitions - constants, set())

    def test_owner_edit_has_reason_and_old_new_audit(self):
        proc = re.search(r"(?ims)^Public Sub AT_UpdateOwnerDetails\b(.*?)(?:^End Sub)", FINANCE)
        self.assertIsNotNone(proc)
        body = proc.group(1)
        for token in ["AT_RequirePermission", "Len(Trim$(reason))", "tblAuditEvents", "AT_AddAuditChange", "FullName", "ApartmentCount"]:
            self.assertIn(token, body)
        self.assertIn("tblAuditChanges", FINANCE)

    def test_account_creation_is_permissioned_and_has_zero_opening_balance(self):
        proc = re.search(r"(?ims)^Public Function AT_CreateAccount\b(.*?)(?:^End Function)", FINANCE)
        self.assertIsNotNone(proc)
        body = proc.group(1)
        self.assertIn('AT_RequirePermission "SETTINGS_ADMIN"', body)
        self.assertIn('AT_Audit "ACCOUNT_CREATE"', body)
        self.assertIn('rs!IsActive = True', body)
        self.assertNotIn('AT_InsertAccountMovement', body)
        self.assertIn('frmAccountCreate', UI)
        self.assertIn('AT_UI_SaveAccount', UI)

    def test_owner_search_and_statement_surfaces_exist(self):
        self.assertIn('"frmOwnerSearch"', UI)
        self.assertIn('"frmOwnerStatementPick"', UI)
        self.assertIn('"frmOwnerStatementList"', UI)
        self.assertIn("AT_UI_SearchOwners", UI)
        self.assertIn("AT_UI_OpenOwnerStatement", UI)
        queries = [name for name, _ in report_query_strings()]
        self.assertIn("qryAT_OwnerStatement", queries)
        self.assertIn("qryAT_CreditLedger", queries)

    def test_entry_forms_are_editable_and_grid_forms_read_only(self):
        base = re.search(r"(?ims)^Private Sub AT_SetFormBase\b(.*?)(?:^End Sub)", UI).group(1)
        self.assertNotIn("AllowEdits = False", base)
        grid = re.search(r"(?ims)^Private Sub AT_CreateReadOnlyGridForm\b(.*?)(?:^End Sub)", UI).group(1)
        self.assertIn("AllowEdits = False", grid)
        self.assertIn("AT_BuildReportQueries", UI)

    def test_initial_ui_registers_owner_search_edit_and_credit_forms(self):
        for name in ["frmOwnerCreate", "frmOwnerEdit", "frmOwnerSearch", "frmOwnerStatementPick", "frmCreditApply", "frmReceiptCorrection", "frmOwnerStatementList"]:
            self.assertIn(f'"{name}"', UI)

    def test_reversal_preserves_owner_credit_when_cash_not_refunded(self):
        proc = re.search(r"(?ims)^Public Sub AT_ReverseReceiptAllocation\b(.*?)(?:^End Sub)", FINANCE).group(1)
        self.assertIn("RECEIPT_ALLOCATION_REVERSED", proc)
        self.assertIn("tblOwnerCreditMovements", proc)
        self.assertIn("AT_CompleteApproval", proc)

    def test_no_direct_financial_delete_statements(self):
        self.assertNotRegex(FINANCE, r"(?i)DELETE\s+FROM\s+(tblFinancialDocuments|tblAccountMovements|tblOwnerCreditMovements|tblChargeSettlementLines)")

    def test_report_sql_parentheses_balance_and_names_unique(self):
        queries = report_query_strings()
        self.assertGreaterEqual(len(queries), 16)
        names = [name for name, _ in queries]
        self.assertEqual(len(names), len(set(names)))
        for name, sql in queries:
            self.assertTrue(sql.strip(), name)
            balance = 0
            for ch in sql:
                if ch == "(": balance += 1
                elif ch == ")": balance -= 1
                self.assertGreaterEqual(balance, 0, f"Unexpected closing parenthesis in {name}")
            self.assertEqual(balance, 0, f"Unclosed SQL parenthesis in {name}: {sql}")


class FinancialRuleSmokeTests(unittest.TestCase):
    """Checks the intended arithmetic only, not the Access implementation/runtime."""
    def test_partial_payment_remaining(self):
        gross, prior_paid, current_paid = 9 * 3000, 20000, 5000
        self.assertEqual(gross - prior_paid - current_paid, 2000)

    def test_overpayment_becomes_separate_credit(self):
        due, received = 2000, 5000
        allocated = min(due, received)
        credit = max(received - allocated, 0)
        self.assertEqual((allocated, credit), (2000, 3000))

    def test_reserved_transfer_reduces_available_balance(self):
        book_balance, active_reservations = 10000, 7000
        self.assertEqual(book_balance - active_reservations, 3000)
        self.assertGreater(5000, book_balance - active_reservations)

    def test_reverse_allocation_reopens_due_and_credits_owner(self):
        due_before, reversed_allocation = 0, 2500
        due_after = due_before + reversed_allocation
        owner_credit = reversed_allocation  # cash stays in account unless separate refund voucher posts
        self.assertEqual((due_after, owner_credit), (2500, 2500))

    def test_credit_use_cannot_exceed_source(self):
        source_credit, used_before, requested = 5000, 3600, 2000
        available = max(source_credit - used_before, 0)
        self.assertEqual(available, 1400)
        self.assertGreater(requested, available)


if __name__ == "__main__":
    unittest.main(verbosity=2)
