#!/usr/bin/env python3
"""Static/source checks for Archon Access 0.5.6. Does not compile or run Microsoft Access."""
from __future__ import annotations
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VBA = ROOT / "src" / "vba"
MODULES = sorted(VBA.glob("*.bas"))
TEXTS = {p.stem: p.read_text(encoding="ascii") for p in MODULES}
ALL = "\n".join(TEXTS.values())
BUILD = TEXTS["modArchonBuild"]
FINANCE = TEXTS["modArchonFinance"]
PROPERTY = TEXTS["modArchonProperty"]
UI = TEXTS["modArchonUI"]
REPORTS = TEXTS["modArchonReports"]
OPERATIONS = TEXTS["modArchonOperations"]
ATTACHMENTS = TEXTS["modArchonAttachments"]


def split_top_level_fields(ddl: str) -> set[str]:
    open_at = ddl.find("(")
    if open_at < 0:
        return set()
    depth = 0
    quoted = False
    chunks: list[str] = []
    current: list[str] = []
    for ch in ddl[open_at + 1:]:
        if ch == '"':
            quoted = not quoted
        if not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                if depth == 0:
                    if current:
                        chunks.append("".join(current))
                    break
                depth -= 1
            elif ch == "," and depth == 0:
                chunks.append("".join(current))
                current = []
                continue
        current.append(ch)
    fields: set[str] = set()
    for part in chunks:
        match = re.match(r"\s*(\w+)\s+", part)
        if match and match.group(1).upper() != "CONSTRAINT":
            fields.add(match.group(1).lower())
    return fields


def schema_fields() -> dict[str, set[str]]:
    result: dict[str, set[str]] = {}
    for literal in re.findall(r'"(CREATE TABLE [^"\n]+)"', BUILD):
        match = re.match(r"CREATE TABLE (\w+)\s+", literal, re.I)
        if match:
            result[match.group(1).lower()] = split_top_level_fields(literal)
    return result


def query_definitions() -> list[tuple[str, str]]:
    lines = REPORTS.splitlines()
    found: list[tuple[str, str]] = []
    i = 0
    while i < len(lines):
        m = re.search(r'AT_ReplaceQuery\s+db,\s*"([^"]+)"', lines[i])
        if not m:
            i += 1
            continue
        name = m.group(1)
        chunks: list[str] = []
        j = i
        while j < len(lines):
            if j > i and re.search(r'AT_ReplaceQuery\s+db,\s*"', lines[j]):
                break
            if j > i and not lines[j].strip():
                break
            # Normal VBA string literals.
            for sm in re.finditer(r'"((?:[^"\n]|"")*)"', lines[j]):
                token = sm.group(1).replace('""', '"')
                # Skip the QueryDef name literal.
                if j == i and token == name:
                    continue
                # AT_U literals are encoded as 4-digit Unicode code points.
                prefix = lines[j][:sm.start()]
                if prefix.rfind("AT_U(") > prefix.rfind(")"):
                    try:
                        token = "".join(chr(int(part, 16)) for part in token.split(","))
                    except ValueError:
                        pass
                chunks.append(token)
            j += 1
        found.append((name, "".join(chunks)))
        i = j
    return found


class PackageIntegrity(unittest.TestCase):
    def test_expected_vba_modules_are_present_and_unique(self):
        expected = {
            "modArchonCore", "modArchonBuild", "modArchonFinance", "modArchonProperty",
            "modArchonClosing", "modArchonOperations", "modArchonConfig", "modArchonAttachments",
            "modArchonReports", "modArchonPrint", "modArchonUI", "modArchonInstaller",
        }
        self.assertEqual(set(TEXTS), expected)
        names = re.findall(r'(?im)^Attribute VB_Name\s*=\s*"([^"]+)"', ALL)
        self.assertEqual(len(names), len(set(names)))
        self.assertEqual(set(names), expected)

    def test_vba_sources_are_ascii_safe_and_have_explicit_option(self):
        for path in MODULES:
            content = path.read_bytes()
            self.assertTrue(all(byte < 128 for byte in content), path.name)
            text = content.decode("ascii")
            self.assertIn("Option Explicit", text, path.name)

    def test_procedure_blocks_are_balanced(self):
        for name, text in TEXTS.items():
            starts = len(re.findall(r"(?im)^\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function)\s+\w+", text))
            ends = len(re.findall(r"(?im)^\s*End\s+(?:Sub|Function)\s*$", text))
            self.assertEqual(starts, ends, name)

    def test_internal_at_references_resolve(self):
        declared = set(re.findall(r"(?im)^\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function)\s+(AT_\w+)", ALL))
        constants = set(re.findall(r"(?im)^\s*(?:Public|Private)?\s*Const\s+(AT_\w+)", ALL))
        references = set(re.findall(r"\b(AT_\w+)\b", ALL))
        self.assertEqual(sorted(references - declared - constants), [])

    def test_unicode_helper_literals_are_well_formed(self):
        calls = re.findall(r'AT_U\("([^"]*)"\)', ALL)
        self.assertGreater(len(calls), 300)
        for literal in calls:
            parts = literal.split(",")
            self.assertTrue(all(re.fullmatch(r"[0-9A-Fa-f]{4}", part) for part in parts), literal[:80])
            self.assertTrue(all(int(part, 16) <= 0xFFFF for part in parts))

    def test_schema_and_relation_targets_are_complete(self):
        tables = schema_fields()
        self.assertGreaterEqual(len(tables), 50)
        relations = re.findall(r'"(rel[^|]+)\|([^|]+)\|([^|]+)\|([^|]+)\|([^"|]+)"', BUILD)
        self.assertGreaterEqual(len(relations), 80)
        issues = []
        for relation, child, child_field, parent, parent_field in relations:
            if child.lower() not in tables or child_field.lower() not in tables.get(child.lower(), set()):
                issues.append((relation, "child", child, child_field))
            if parent.lower() not in tables or parent_field.lower() not in tables.get(parent.lower(), set()):
                issues.append((relation, "parent", parent, parent_field))
        self.assertEqual(issues, [])


    def test_builder_schema_matches_reference_and_has_55_tables(self):
        tables = schema_fields()
        self.assertEqual(len(tables), 55)
        sql = (ROOT / "src" / "sql" / "schema_reference.sql").read_text(encoding="utf-8")
        ref = {}
        for ddl in re.findall(r"(?im)^(CREATE TABLE .*?);\s*$", sql):
            name = re.match(r"CREATE TABLE (\w+)\s+", ddl, re.I).group(1).lower()
            ref[name] = split_top_level_fields(ddl)
        self.assertEqual(set(tables), set(ref))
        self.assertEqual({k: len(v) for k, v in tables.items()}, {k: len(v) for k, v in ref.items()})
        self.assertEqual(sum(map(len, tables.values())), 444)

    def test_relationship_and_index_declarations_are_complete(self):
        relations = re.findall(r'"(rel[^"\n]+\|[^"\n]+\|[^"\n]+\|[^"\n]+\|[^"\n]+)"', BUILD)
        self.assertEqual(len(relations), 86)
        self.assertEqual(len(set(x.split("|")[0] for x in relations)), 86)
        tables = schema_fields()
        index_rows = re.findall(r'AT_EnsureIndex\s+db,\s*"([^"]+)",\s*"([^"]+)",\s*"([^"]+)"', BUILD)
        self.assertEqual(len(index_rows), 43)
        for table, _index, fields in index_rows:
            self.assertIn(table.lower(), tables)
            for field in fields.split(","):
                self.assertIn(field.strip().lower(), tables[table.lower()], (table, _index, field))

    def test_vba_logical_statement_continuations_stay_within_limit(self):
        # VBA limits explicit continuation lines in one logical statement.
        for path in MODULES:
            run = 0
            maximum = 0
            for line in path.read_text(encoding="ascii").splitlines():
                if line.rstrip().endswith(" _"):
                    run += 1
                    maximum = max(maximum, run)
                else:
                    run = 0
            self.assertLessEqual(maximum, 24, f"Too many continuations in {path.name}: {maximum}")


    def test_permission_catalog_has_one_arabic_label_per_code(self):
        seed = re.search(r"codes = Array\((.*?)\)\n    For i", BUILD, re.S)
        self.assertIsNotNone(seed)
        codes = re.findall(r'"([A-Z_]+)"', seed.group(1))
        helper = re.search(r"Private Function AT_PermissionName.*?End Function", BUILD, re.S)
        self.assertIsNotNone(helper)
        labels = re.findall(r'Case "([A-Z_]+)"\s*:\s*AT_PermissionName\s*=\s*AT_U\("([0-9A-Fa-f,]+)"\)', helper.group(0))
        self.assertEqual(len(codes), 31)
        self.assertEqual(len(labels), len(codes))
        self.assertEqual(set(codes), {code for code, _ in labels})


    def test_error_handler_labels_are_not_followed_by_declarations(self):
        # A declaration cannot serve as the first statement beneath an error-handler label.
        problems = []
        for path in MODULES:
            lines = path.read_text(encoding="ascii").splitlines()
            for index, line in enumerate(lines):
                if re.match(r"^\s*[A-Za-z_]\w*\s*:\s*$", line):
                    next_index = index + 1
                    while next_index < len(lines) and (
                        not lines[next_index].strip() or lines[next_index].lstrip().startswith("'")
                    ):
                        next_index += 1
                    if next_index < len(lines) and re.match(r"^\s*(?:Dim|Static|Const)\b", lines[next_index], re.I):
                        problems.append(f"{path.name}:{index + 1}->{next_index + 1}")
        self.assertEqual(problems, [])

    def test_reserved_enum_keyword_is_not_used_as_an_identifier(self):
        # VBA identifiers are case-insensitive; eNum is the reserved keyword Enum.
        self.assertIsNone(re.search(r"\beNum\b", ALL, re.I))
        self.assertGreater(len(re.findall(r"\bsavedErrNumber\b", ALL, re.I)), 100)

    def test_installer_entry_point_declares_handler_state_before_error_label(self):
        installer = TEXTS["modArchonInstaller"]
        match = re.search(r"(?ims)^Public Function AT_InstallSplitProject_Run\(\) As Boolean(.*?)^End Function", installer)
        self.assertIsNotNone(match)
        body = match.group(1)
        declaration = "Dim savedErrNumber As Long, eDesc As String, savedErrSource As String"
        self.assertRegex(body, r"(?im)^\s*" + re.escape(declaration) + r"\s*$")
        self.assertLess(body.index(declaration), body.index("EH:"))
        self.assertNotRegex(body, r"(?im)^EH:\s*\n\s*(?:Dim|Static|Const)\b")

    def test_release_manifest_covers_all_package_files_and_hashes(self):
        import hashlib, json
        manifest_path = ROOT / "MANIFEST_SHA256.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(manifest["version"], "0.5.6")
        actual = {
            p.relative_to(ROOT).as_posix(): p
            for p in ROOT.rglob("*") if p.is_file() and p != manifest_path
        }
        declared = {entry["path"]: entry for entry in manifest["files"]}
        self.assertEqual(set(actual), set(declared))
        self.assertEqual(manifest["files_count"], len(actual))
        for relative, path in actual.items():
            entry = declared[relative]
            data = path.read_bytes()
            self.assertEqual(entry["size_bytes"], len(data), relative)
            self.assertEqual(entry["sha256"], hashlib.sha256(data).hexdigest(), relative)

    def test_powershell_import_list_matches_all_vba_modules(self):
        script = (ROOT / "tools" / "Install_Archon_Tech.ps1").read_text(encoding="utf-8")
        imported = set(re.findall(r"'((?:modArchon)\w+\.bas)'", script))
        expected = {path.name for path in MODULES}
        self.assertEqual(imported, expected)

    def test_powershell_uses_no_argument_function_entry_point(self):
        installer = TEXTS["modArchonInstaller"]
        script = (ROOT / "tools" / "Install_Archon_Tech.ps1").read_text(encoding="utf-8-sig")
        self.assertRegex(installer, r"(?im)^Public Function AT_InstallSplitProject_Run\(\) As Boolean")
        self.assertIn("$access.Run('AT_InstallSplitProject_Run')", script)
        self.assertNotRegex(script, r"(?im)^\s*\$access\.Run\('AT_InstallSplitProject_Run',")
        self.assertIn("[IO.File]::WriteAllText($bootstrapConfigPath, $bePath, [Text.Encoding]::Unicode)", script)
        self.assertIn("fso.OpenTextFile(configPath, 1, False, -1)", installer)
        self.assertIn("fso.DeleteFile configPath, True", installer)
        self.assertIn("if (-not [bool]$installResult)", script)
        self.assertIn("if (Test-Path -LiteralPath $bootstrapConfigPath -PathType Leaf)", script)

    def test_access_run_has_no_byref_backend_argument_and_saves_all_on_success(self):
        script = (ROOT / "tools" / "Install_Archon_Tech.ps1").read_text(encoding="utf-8-sig")
        self.assertNotIn("[ref]$bePath", script)
        self.assertNotIn("$backendPathArgument", script)
        self.assertIn("$access.Quit(1) # acQuitSaveAll", script)

    def test_failed_install_discards_partial_database_before_cleanup(self):
        script = (ROOT / "tools" / "Install_Archon_Tech.ps1").read_text(encoding="utf-8-sig")
        self.assertIn("$access.Quit(2)", script)  # acQuitSaveNone is intentional in failure cleanup.
        self.assertIn("Remove-Item -LiteralPath $partial -Force -ErrorAction Stop", script)

    def test_print_templates_use_existing_saved_queries(self):
        printed = TEXTS["modArchonPrint"]
        known = {name for name, _ in query_definitions()}
        names = set(re.findall(r'AT_CreateTabularReport\s+"[^"]+",\s*"([^"]+)"', printed))
        names.add("qryAT_ReceiptPrint")
        self.assertEqual(names - known, set())
        self.assertGreaterEqual(len(names), 6)

    def test_restore_uses_verified_staging_and_tracked_rollback(self):
        script = (ROOT / "tools" / "Backup_Restore_Archon.ps1").read_text(encoding="utf-8-sig")
        self.assertIn("$safety = Join-Path $ProjectFolder", script)
        self.assertIn("$null = Test-BackupFolder $incoming", script)
        self.assertIn("$null = Test-BackupFolder $safety", script)
        for flag in ["$oldFeMoved", "$oldBeMoved", "$oldAttachmentsMoved", "$newFePlaced", "$newBePlaced", "$newAttachmentsPlaced"]:
            self.assertIn(flag, script)
        self.assertIn("Move-Item -LiteralPath $oldFE -Destination $fe", script)
        self.assertIn("Move-Item -LiteralPath $oldBE -Destination $be", script)
        self.assertIn("$savedAuditLog", script)

    def test_power_shell_files_are_utf8_with_single_bom(self):
        paths = [
            ROOT / "tools" / "Install_Archon_Tech.ps1",
            ROOT / "tools" / "Backup_Restore_Archon.ps1",
            ROOT / "tests" / "Verify-Windows-PowerShell.ps1",
        ]
        marker = b"\xef\xbb\xbf"
        for path in paths:
            data = path.read_bytes()
            self.assertTrue(data.startswith(marker), str(path))
            self.assertFalse(data.startswith(marker + marker), str(path))
            self.assertEqual(len(data) - len(data.lstrip(marker)), 3, str(path))
            data[3:].decode("utf-8")
            self.assertEqual(data.count(b"\n"), data.count(b"\r\n"), str(path))

    def test_one_click_launcher_checks_parser_before_installer(self):
        launcher = (ROOT / "تشغيل_المثبت_وفحصه.cmd").read_text(encoding="ascii")
        self.assertIn("Verify-Windows-PowerShell.ps1", launcher)
        self.assertIn("if errorlevel 1 goto :failed", launcher)
        self.assertIn("Install_Archon_Tech.ps1", launcher)
        self.assertLess(launcher.index("Verify-Windows-PowerShell.ps1"), launcher.index("Install_Archon_Tech.ps1"))

    def test_windows_powershell_parser_check_is_supplied(self):
        parser_test = ROOT / "tests" / "Verify-Windows-PowerShell.ps1"
        self.assertTrue(parser_test.is_file())
        parser_text = parser_test.read_text(encoding="utf-8-sig")
        self.assertIn("System.Management.Automation.Language.Parser]::ParseFile", parser_text)
        self.assertIn("Get-ChildItem -LiteralPath $projectRoot -Filter '*.ps1' -File -Recurse -Force", parser_text)

    def test_report_queries_are_built_once_before_query_consuming_forms(self):
        self.assertEqual(len(re.findall(r"(?im)^\s*Call AT_BuildReportQueries\s*$", UI)), 1)
        self.assertNotIn("AT_BuildReportQueries", TEXTS["modArchonInstaller"])
        self.assertLess(UI.index("Call AT_BuildReportQueries"), UI.index("AT_CreateMainForm"))

    def test_attachment_hash_and_ui_buttons_are_wired(self):
        self.assertIn("Public Function AT_SHA256File", ATTACHMENTS)
        self.assertIn("AT_CALG_SHA_256", ATTACHMENTS)
        self.assertIn("rs!SHA256 = AT_SHA256File(targetPath)", ATTACHMENTS)
        for proc in ["AT_UI_AttachReceiptFile", "AT_UI_AttachPaymentFile", "AT_UI_AttachExpensePaymentFile"]:
            self.assertRegex(UI, rf"(?im)^Public Function {proc}\b")
        self.assertIn('"=AT_UI_AttachReceiptFile()"', UI)
        self.assertIn('"=AT_UI_AttachPaymentFile()"', UI)
        self.assertIn('"=AT_UI_AttachExpensePaymentFile()"', UI)

    def test_audit_fallback_is_written_and_double_failure_is_reported(self):
        self.assertIn("Archon_Audit_Fallback.log", BUILD)
        self.assertIn("vbObjectError + 1491", BUILD)
        self.assertIn("fallbackErrNo", BUILD)

    def test_financial_closing_uses_inclusive_end_of_day_and_date_only_period_check(self):
        closing = TEXTS["modArchonClosing"]
        self.assertIn("DateValue(transactionDate)", closing)
        finance = TEXTS["modArchonFinance"]
        self.assertIn('M.MovementDate<" & AT_SqlDate(DateAdd("d", 1, DateValue(asOfDate)))', finance)
        self.assertIn("AT_AccountBookBalance(CLng(rs!AccountID), DateValue(periodEnd))", closing)

    def test_backend_relink_preflights_and_rolls_back_old_connections(self):
        installer = TEXTS["modArchonInstaller"]
        relink = re.search(r"(?ims)^Public Sub AT_RelinkBackend\b(.*?)(?:^End Sub)", installer)
        self.assertIsNotNone(relink)
        body = relink.group(1)
        self.assertIn("OpenDatabase(backendPath, False, True)", body)
        self.assertIn("AT_TableExistsInDatabase(dbBE, sourceTable)", body)
        self.assertIn("oldConnects(i) = tdf.Connect", body)
        self.assertIn("tdf.Connect = oldConnects(i)", body)
        self.assertIn("tdf.RefreshLink", body)

    def test_backup_detects_hidden_access_lock_files(self):
        script = (ROOT / "tools" / "Backup_Restore_Archon.ps1").read_text(encoding="utf-8-sig")
        self.assertRegex(script, r"Get-ChildItem -LiteralPath \$ProjectFolder -Force -File")
        self.assertIn("'*.laccdb'", script)
        self.assertIn("'*.ldb'", script)

    def test_restore_refuses_mismatched_source_folder(self):
        script = (ROOT / "tools" / "Backup_Restore_Archon.ps1").read_text(encoding="utf-8-sig")
        self.assertIn("$manifest.SourceProjectFolder", script)
        self.assertIn("$recordedRoot.Equals($currentRoot", script)
        self.assertIn("لمنع FE من الاحتفاظ بروابط إلى مسار خلفية خاطئ", script)

    def test_audit_fallback_log_is_preserved_by_backup_restore(self):
        script = (ROOT / "tools" / "Backup_Restore_Archon.ps1").read_text(encoding="utf-8-sig")
        self.assertIn("Archon_Audit_Fallback.log", script)
        self.assertIn("AuditFallbackLogIncluded", script)
        self.assertIn("$oldAuditMoved", script)
        self.assertIn("$newAuditPlaced", script)

    def test_historical_031_test_is_not_in_active_test_suite(self):
        self.assertFalse((ROOT / "tests" / "فحص_ساكن_المرجع_0.3.1.py").exists())
        archived = ROOT / "reference" / "tests_0_3_1" / "فحص_ساكن_المرجع_0.3.1.py"
        self.assertTrue(archived.is_file())
        self.assertIn("access_0_3_1", archived.read_text(encoding="utf-8"))

    def test_owner_property_campaign_model_exists(self):
        tables = schema_fields()
        for table in ["tblBuildings", "tblBuildingOwnership", "tblUnits", "tblUnitOccupancy", "tblUnitExemptionRequests", "tblContributionCampaigns", "tblCampaignBuildingTargets", "tblCampaignUnitTargets", "tblChargeLines"]:
            self.assertIn(table.lower(), tables)
        for proc in ["AT_CreateBuilding", "AT_AddUnit", "AT_TransferBuildingOwnership", "AT_AddUnitOccupancy", "AT_RequestUnitExemption", "AT_DecideUnitExemption", "AT_GenerateCampaignCharges"]:
            self.assertRegex(PROPERTY, rf"(?im)^Public (?:Sub|Function) {proc}\b")
        self.assertIn("OwnerIDSnapshot", BUILD)
        self.assertIn("UnitCodeSnapshot", BUILD)

    def test_receipt_manual_allocation_and_overpayment_logic_are_wired(self):
        for proc in ["AT_SetReceiptAllocation", "AT_ClearReceiptDraftAllocations", "AT_ApproveReceipt"]:
            self.assertIn(proc, FINANCE)
        self.assertIn("tblReceiptAllocationLines", BUILD)
        self.assertIn("qryAT_ReceiptDraftAllocations", UI)
        self.assertIn("AT_UI_SaveReceiptAllocation", UI)
        self.assertIn("AT_CreditMovement", FINANCE)

    def test_single_owner_charge_uses_real_units_not_owner_summary(self):
        self.assertIn("tblBuildingOwnership", FINANCE)
        self.assertIn("tblUnits", FINANCE)
        self.assertIn("tblChargeLines", FINANCE)
        self.assertIn("AT_UnitIsExempt", FINANCE)
        self.assertIn("OwnerIDSnapshot", FINANCE)

    def test_expense_must_precede_normal_payment_and_payment_updates_expense(self):
        self.assertIn("AT_CreateExpenseDraft", OPERATIONS)
        self.assertIn("AT_CreateExpensePaymentDraft", OPERATIONS)
        self.assertIn("StatusCode='PAID'", FINANCE)
        self.assertIn("expenseID = 0 And creditMovementID = 0", FINANCE)
        self.assertIn("AT_EnsureDateOpen expenseDate", OPERATIONS)

    def test_attachments_are_linked_to_receipts_and_expenses(self):
        self.assertIn("tblReceiptAttachments", BUILD)
        self.assertIn("tblExpenseAttachments", BUILD)
        self.assertIn("INSERT INTO tblReceiptAttachments", ATTACHMENTS)
        self.assertIn("INSERT INTO tblExpenseAttachments", ATTACHMENTS)
        self.assertIn("DBEngine(0)(0).CommitTrans", ATTACHMENTS)

    def test_transfers_reserve_until_confirmation_and_not_income(self):
        self.assertIn("AT_ApproveTransfer", FINANCE)
        self.assertIn("AT_ConfirmTransfer", FINANCE)
        self.assertIn("tblAccountReservations", FINANCE)
        self.assertIn("DestinationAccountID", FINANCE)

    def test_financial_controls_and_settings_are_present(self):
        for proc in ["AT_EnsureDateOpen", "AT_CreateFinancialPeriod", "AT_BeginFinancialClosing", "AT_FinalizeFinancialClosing", "AT_ReopenFinancialPeriod"]:
            self.assertIn(proc, ALL)
        self.assertIn("PROJECT_CONTACT", TEXTS["modArchonConfig"])
        self.assertIn("AT_CreateContributionRate", ALL)

    def test_current_ownership_filter_is_outside_left_join_on(self):
        queries = dict(query_definitions())
        self.assertIn("qryAT_CurrentBuildingOwnership", queries)
        helper = queries["qryAT_CurrentBuildingOwnership"]
        self.assertIn("BO.StartDate<=Date()", helper)
        self.assertIn("BO.EndDate Is Null OR BO.EndDate>=Date()", helper)
        for name in ["qryAT_BuildingCurrentOwners", "qryAT_UnitDirectory"]:
            self.assertIn("qryAT_CurrentBuildingOwnership", queries[name])
            self.assertNotIn("tblBuildingOwnership", queries[name])
            self.assertNotRegex(queries[name], r"(?i)\bLEFT\s+JOIN\b[^;]*?\bON\b[^;]*?\b(?:AND|OR)\b")
        names = [name for name, _ in query_definitions()]
        self.assertLess(names.index("qryAT_CurrentBuildingOwnership"), names.index("qryAT_BuildingCurrentOwners"))
        self.assertLess(names.index("qryAT_CurrentBuildingOwnership"), names.index("qryAT_UnitDirectory"))

    def test_each_report_query_is_opened_and_errors_identify_the_query(self):
        self.assertIn("existingQuery.OpenRecordset(dbOpenSnapshot)", REPORTS)
        self.assertIn('Err.Raise savedErrNumber, "AT_ReplaceQuery[" & queryName & "]"', REPORTS)
        self.assertIn("existingSql = existingQuery.SQL", REPORTS)
        self.assertIn("Failed to restore the previous query definition", REPORTS)
        self.assertIn("savedErrSource = Err.Source", UI)
        self.assertIn("Err.Raise savedErrNumber, savedErrSource, eDesc", UI)
        self.assertIn("Call AT_BuildReportQueries", UI)
        self.assertNotIn("vbObjectError + 690", UI)
        self.assertIn("savedErrSource = Err.Source", TEXTS["modArchonInstaller"])

    def test_report_query_names_unique_and_sql_parentheses_balanced(self):
        queries = query_definitions()
        self.assertGreaterEqual(len(queries), 20)
        names = [name for name, _ in queries]
        self.assertEqual(len(names), len(set(names)))
        for name, sql in queries:
            self.assertTrue(sql.strip(), name)
            balance = 0
            in_quote = False
            i = 0
            while i < len(sql):
                ch = sql[i]
                if ch == "'":
                    if in_quote and i + 1 < len(sql) and sql[i + 1] == "'":
                        i += 2
                        continue
                    in_quote = not in_quote
                elif not in_quote and ch == "(":
                    balance += 1
                elif not in_quote and ch == ")":
                    balance -= 1
                    self.assertGreaterEqual(balance, 0, name)
                i += 1
            self.assertEqual(balance, 0, f"SQL parentheses unbalanced in {name}")

    def test_no_direct_deletes_of_approved_ledgers(self):
        self.assertNotRegex(FINANCE, r"(?i)DELETE\s+FROM\s+(tblFinancialDocuments|tblAccountMovements|tblOwnerCreditMovements|tblChargeSettlementLines)")


    def test_main_navigation_forms_and_button_callbacks_resolve(self):
        manifest_block = UI.split("Private Function AT_RequiredGeneratedFormNames()", 1)[1].split("End Function", 1)[0]
        created = set(re.findall(r'"(frm[A-Za-z0-9_]+)"', manifest_block))
        navigation = set(re.findall(r'=AT_UI_OpenForm\(""([^\"]+)""\)', UI))
        self.assertGreaterEqual(len(navigation), 40)
        self.assertEqual(sorted(navigation - created), [])
        callbacks = set(re.findall(r'"=(AT_UI_\w+)\(\)"', UI))
        declared = set(re.findall(r'(?im)^Public Function (AT_UI_\w+)\b', UI))
        self.assertEqual(sorted(callbacks - declared), [])

    def test_builder_and_report_creation_preserve_original_errors(self):
        self.assertIn('Err.Raise savedErrNumber, "AT_BuildDatabase", eDesc', BUILD)
        self.assertIn('Err.Raise savedErrNumber, "AT_BuildPrintableReports", eDesc', TEXTS["modArchonPrint"])
        self.assertIn('Err.Raise savedErrNumber, "AT_CreateTabularReport", eDesc', TEXTS["modArchonPrint"])

    def test_split_installer_and_backup_tools_exist(self):
        self.assertTrue((ROOT / "tools" / "Install_Archon_Tech.ps1").exists())
        self.assertTrue((ROOT / "tools" / "Backup_Restore_Archon.ps1").exists())
        self.assertIn("AT_CreateSchemaOnDatabase dbBE", TEXTS["modArchonInstaller"])
        self.assertIn("AT_LinkBackendTables", TEXTS["modArchonInstaller"])


class FormBuilderIntegrity(unittest.TestCase):
    def required_forms(self) -> list[str]:
        block = UI.split("Private Function AT_RequiredGeneratedFormNames()", 1)[1].split("End Function", 1)[0]
        return re.findall(r'"(frm[A-Za-z0-9_]+)"', block)

    def test_complete_form_inventory_has_48_unique_saved_names(self):
        required = self.required_forms()
        direct_names = re.findall(r'(?im)^\s*AT_SaveAndCloseForm\s+(?:actualName|n),\s*"(frm[A-Za-z0-9_]+)"', UI)
        grid_names = re.findall(r'(?im)^\s*AT_CreateReadOnlyGridForm\s+"(frm[A-Za-z0-9_]+)"', UI)
        self.assertEqual(len(direct_names), 29)
        self.assertEqual(len(grid_names), 19)
        self.assertEqual(len(required), 48)
        self.assertEqual(len(required), len(set(required)))
        self.assertEqual(set(required), set(direct_names) | set(grid_names))
        # 29 directly named forms plus one reusable grid builder = 30 CreateForm sites.
        self.assertEqual(len(re.findall(r'(?im)^\s*Set f\s*=\s*CreateForm\b', UI)), 30)
        self.assertEqual(len(re.findall(r'(?im)^\s*AT_SaveAndCloseForm\s+', UI)), 30)

    def test_all_generated_forms_use_the_validated_temp_save_rename_lifecycle(self):
        self.assertNotRegex(UI, r'(?im)^\s*DoCmd\.Save\s+acForm\s*,\s*"')
        self.assertNotRegex(UI, r'(?im)^\s*DoCmd\.Close\s+acForm\s*,\s*"frm')
        helper = UI.split("Private Sub AT_SaveAndCloseForm(", 1)[1].split("End Sub", 1)[0]
        operations = [
            "DoCmd.Save acForm, temporaryName",
            "DoCmd.Close acForm, temporaryName, acSaveNo",
            "DoCmd.Rename desiredName, acForm, temporaryName",
            "DoCmd.OpenForm desiredName, acDesign, , , , acHidden",
            "DoCmd.Close acForm, desiredName, acSaveNo",
        ]
        positions = [helper.find(x) for x in operations]
        self.assertTrue(all(x >= 0 for x in positions), positions)
        self.assertEqual(positions, sorted(positions))
        self.assertIn("If Not AT_FormExists(desiredName) Then", helper)
        self.assertIn("If AT_FormExists(temporaryName) Then", helper)
        self.assertIn("If AT_FormIsOpen(temporaryName) Then", helper)
        self.assertIn("If AT_FormExists(desiredName) Then DoCmd.DeleteObject acForm, desiredName", helper)
        self.assertIn("ElseIf Not temporaryExistedBefore Then", helper)
        self.assertIn("If AT_FormExists(temporaryName) Then", helper)
        self.assertIn("Err.Clear", helper)
        self.assertIn("temporaryExistedBefore = AT_FormExists(temporaryName)", helper)
        self.assertIn("If AT_FormExists(desiredName) Then", helper)
        # Existing final objects are rejected, never silently overwritten.
        self.assertIn("The destination form already exists; it was not replaced", helper)

    def test_control_creation_uses_unbound_column_and_explicit_control_names(self):
        for control_type in ("acLabel", "acTextBox", "acComboBox", "acCommandButton"):
            self.assertRegex(UI, rf'CreateControl\(formName,\s*{control_type},\s*acDetail,\s*,\s*,')
        self.assertGreaterEqual(len(re.findall(r'(?im)^\s*c\.Name\s*=\s*controlName', UI)), 4)
        self.assertNotRegex(UI, r'CreateControl\([^\n]*,\s*controlName\s*,')
        # The shared helper assigns a generated name to all controls and labels.
        self.assertIn('AT_AddLabel formName, "lbl_" & controlName', UI)
        self.assertIn('c.Name = controlName: c.Caption = captionText', UI)

    def test_each_direct_form_has_unique_generated_controls_and_all_f_references_exist(self):
        create_sites = list(re.finditer(r'(?im)^\s*Set\s+f\s*=\s*CreateForm\b', UI))
        self.assertEqual(len(create_sites), 30)
        checked_direct_forms = 0
        for site in create_sites:
            next_create = re.search(r'(?im)^\s*Set\s+f\s*=\s*CreateForm\b', UI[site.end():])
            next_proc = re.search(r'(?im)^\s*(?:Public|Private)\s+(?:Sub|Function)\s+\w+', UI[site.end():])
            ends = [x.start() for x in (next_create, next_proc) if x is not None]
            segment_end = (site.end() - site.start() + min(ends)) if ends else len(UI) - site.start()
            segment = UI[site.start():site.start() + segment_end]
            finish = re.search(r'(?im)^\s*AT_SaveAndCloseForm\s+(?:actualName|n),\s*"(frm[A-Za-z0-9_]+)"', segment)
            if finish is None:
                # The one dynamic CreateForm call belongs to the shared grid builder;
                # its concrete destination names are enumerated separately above.
                self.assertIn('AT_SaveAndCloseForm n, desiredName', segment)
                continue
            checked_direct_forms += 1
            block = segment[:finish.end()]
            created = []
            calls = re.finditer(r'(?i)\bAT_Add(Label|FieldPair|Combo|Button)\s*(?:\(\s*)?(?:actualName|n)\s*,\s*"([A-Za-z0-9_]+)"', block)
            for call in calls:
                kind, name = call.group(1).lower(), call.group(2)
                if kind in ("fieldpair", "combo"):
                    created.extend(["lbl_" + name, name])
                else:
                    created.append(name)
            referenced = re.findall(r'(?i)\bf\.Controls\("([A-Za-z0-9_]+)"\)', block)
            self.assertEqual(len(created), len(set(created)), finish.group(1) + " duplicate controls")
            self.assertEqual(sorted(set(referenced) - set(created)), [], finish.group(1) + " unresolved f.Controls")
        self.assertEqual(checked_direct_forms, 29)

    def test_runtime_verifier_checks_saved_and_openable_form_objects(self):
        build = UI.split("Public Sub AT_BuildInitialUI()", 1)[1].split("End Sub", 1)[0]
        verify = UI.split("Private Sub AT_VerifyRequiredGeneratedForms()", 1)[1].split("End Sub", 1)[0]
        self.assertIn("AT_VerifyRequiredGeneratedForms", build)
        self.assertIn("AT_FormExists(formName)", verify)
        self.assertIn("DoCmd.OpenForm formName, acNormal, , , acFormReadOnly, acHidden", verify)
        self.assertIn("For Each c In f.Controls", verify)
        self.assertIn("Len(Trim$(c.Name)) = 0", verify)
        self.assertIn("If openedHere Then DoCmd.Close acForm, formName, acSaveNo", verify)
        self.assertIn("For Each ao In CurrentProject.AllForms", UI)

    def test_ui_form_sources_resolve_to_declared_tables_and_report_queries(self):
        tables = schema_fields()
        query_names = {name.lower() for name, _ in query_definitions()}
        sql_literals = re.findall(r'"(SELECT [^"]+)"', UI, re.I)
        references: set[str] = set()
        for sql in sql_literals:
            references.update(re.findall(r'\b(?:FROM|JOIN)\s+([A-Za-z_][A-Za-z0-9_]*)', sql, re.I))
        missing_tables = sorted(x for x in references if x.lower().startswith("tbl") and x.lower() not in tables)
        missing_queries = sorted(x for x in references if x.lower().startswith("qry") and x.lower() not in query_names)
        self.assertEqual(missing_tables, [])
        self.assertEqual(missing_queries, [])

    def test_build_order_and_error_cleanup_do_not_touch_preexisting_forms(self):
        build = UI.split("Public Sub AT_BuildInitialUI()", 1)[1].split("End Sub", 1)[0]
        self.assertLess(build.find("Call AT_BuildReportQueries"), build.find("AT_CreateMainForm"))
        self.assertLess(build.find("AT_VerifyRequiredGeneratedForms"), build.find('DoCmd.OpenForm "frmMain"'))
        self.assertIn("Set formsOpenBeforeBuild = AT_CaptureOpenForms()", build)
        cleanup = UI.split("Private Sub AT_CloseFormsOpenedDuringBuild(", 1)[1].split("End Sub", 1)[0]
        self.assertIn("If Not AT_FormNameWasOpen(formsOpenBeforeBuild, formName) Then", cleanup)
        self.assertIn("DoCmd.Close acForm, formName, acSaveNo", cleanup)
        self.assertNotIn("DoCmd.DeleteObject", cleanup)
        self.assertIn("If Not formsOpenBeforeBuild Is Nothing Then AT_CloseFormsOpenedDuringBuild", build)


class FinancialArithmeticSmoke(unittest.TestCase):
    """Independent arithmetic smoke checks only; not a substitute for Access execution."""
    def test_unit_contribution_rate_is_snapshotted(self):
        old_rate, new_rate, old_units = 3000, 4000, 7
        self.assertEqual(old_units * old_rate, 21000)
        self.assertEqual(old_units * new_rate, 28000)

    def test_partial_and_excess_receipt(self):
        due, received = 7000, 9500
        allocated = min(due, received)
        credit = max(received - allocated, 0)
        self.assertEqual((allocated, credit), (7000, 2500))

    def test_pending_transfer_reserves_source_without_destination_credit(self):
        source_book, pending, destination_book = 12000, 5000, 3400
        self.assertEqual(source_book - pending, 7000)
        self.assertEqual(destination_book, 3400)

    def test_cash_refund_reduces_credit_only_when_paid(self):
        credit_before, approved_refund = 2800, 1200
        self.assertEqual(credit_before - approved_refund, 1600)

    def test_closing_exception_does_not_erase_difference(self):
        book, actual, manager_override = 10000, 9800, True
        difference = actual - book
        self.assertTrue(manager_override)
        self.assertEqual(difference, -200)


if __name__ == "__main__":
    unittest.main(verbosity=2)
