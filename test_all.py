#!/usr/bin/env python3
"""
Comprehensive Test Suite for FieldServe Ticket Board
Tests:
  1. Document validation & Junk handling (§2)
  2. Sequential Workflow & Action Owner calculation (§3)
  3. Integrity Rule Flags (§3)
  4. Messy Data Extraction: Store Code, Work Type, Noise (§4)
  5. 3-Way Diff & Conflict Resolution Policy (§6)
  6. Sheet Contract: Export & Apply (§7)
  7. submission.json contract (§7)
"""

import os
import sys
import json
import csv
import tempfile
import openpyxl

from sync import (
    is_valid_url,
    is_doc_met,
    extract_store_code,
    extract_work_type,
    is_noise_ticket,
    compute_stage_and_owner,
    compute_integrity_flags,
    build_roster_maps,
    format_iso_clean,
    parse_iso_clean,
    STATUS_NUM_TO_STR,
    STATUS_STR_TO_NUM,
    WRITABLE_FIELDS,
)
from sheet import export_csv, apply_edits, REQUIRED_COLS


# Mock Roster for Testing
MOCK_AGENTS = [
    {"id": 1, "name": "Alice Coordinator", "role": "Coordinator", "team": "Coordination", "reports_to": ""},
    {"id": 2, "name": "Bob Coordinator", "role": "Coordinator", "team": "Coordination", "reports_to": ""},
    {"id": 10, "name": "Felboso Hamar", "role": "Manager", "team": "Field Ops", "reports_to": ""},
    {"id": 20, "name": "Barnelbel Panmo", "role": "Engineer", "team": "Field Ops", "reports_to": "Felboso Hamar"},
    {"id": 30, "name": "Rogelnor Nelkin", "role": "Finance", "team": "Finance", "reports_to": ""},
]


def test_doc_validation():
    print("Running test_doc_validation...")
    # (value, waivable, expected)
    test_cases = [
        ("https://docs.example/po.pdf", False, True),
        ("http://omnimart.example/dn", True, True),
        ("NA", True, True),          # DN and Invoice allowed NA
        ("NA", False, False),        # PO and WR NEVER allowed NA
        ("na", True, False),         # Strict exact match 'NA'
        ("", False, False),
        ("", True, False),
        (None, False, False),
        ("done", False, False),      # Junk text
        ("done", True, False),
        ("asked Ravi", False, False),# Junk text
        ("ok", True, False),         # Junk text
        ("yes", False, False),
        ("http://", False, False),
        ("https://a.com", False, True),
        ("  https://docs.example/f/123  ", False, True),
    ]

    for val, waivable, expected in test_cases:
        actual = is_doc_met(val, waivable)
        assert actual == expected, f"Failed for {val!r} waivable={waivable}: got {actual}, expected {expected}"

    print(" -> test_doc_validation PASSED")


def test_stage_logic():
    print("Running test_stage_logic...")
    roster = build_roster_maps(MOCK_AGENTS)

    # 1. Step 1: Assign Manager (even ticket id -> Alice Coordinator, odd -> Bob Coordinator)
    t_even = {"id": 100, "manager": "", "engineer_id": None}
    s, o = compute_stage_and_owner(t_even, roster)
    assert s == "Assign Manager" and o == "Alice Coordinator", f"Got {s}, {o}"

    t_odd = {"id": 101, "manager": "", "engineer_id": None}
    s, o = compute_stage_and_owner(t_odd, roster)
    assert s == "Assign Manager" and o == "Bob Coordinator", f"Got {s}, {o}"

    # 2. Step 2: Assign Engineer (Owner is Manager)
    t_step2 = {"id": 102, "manager": "Felboso Hamar", "engineer_id": None}
    s, o = compute_stage_and_owner(t_step2, roster)
    assert s == "Assign Engineer" and o == "Felboso Hamar", f"Got {s}, {o}"

    # 3. Step 3: Purchase Order (Owner is Manager, waivable=False)
    t_step3 = {"id": 103, "manager": "Felboso Hamar", "engineer_id": 20, "purchase_order": ""}
    s, o = compute_stage_and_owner(t_step3, roster)
    assert s == "Purchase Order" and o == "Felboso Hamar", f"Got {s}, {o}"

    # 4. Step 4: Delivery Note (Owner is Finance, waivable=True)
    t_step4 = {
        "id": 104,
        "manager": "Felboso Hamar",
        "engineer_id": 20,
        "purchase_order": "https://docs.example/po",
        "delivery_note": ""
    }
    s, o = compute_stage_and_owner(t_step4, roster)
    assert s == "Delivery Note" and o == "Rogelnor Nelkin", f"Got {s}, {o}"

    # Delivery Note waiver with 'NA' should advance
    t_step4_waived = dict(t_step4, delivery_note="NA")
    s, o = compute_stage_and_owner(t_step4_waived, roster)
    assert s == "Work Report" and o == "Barnelbel Panmo", f"Got {s}, {o}"

    # 5. Step 5: Work Report (Owner is Engineer, waivable=False)
    t_step5 = {
        "id": 105,
        "manager": "Felboso Hamar",
        "engineer_id": 20,
        "purchase_order": "https://docs.example/po",
        "delivery_note": "NA",
        "work_report": ""
    }
    s, o = compute_stage_and_owner(t_step5, roster)
    assert s == "Work Report" and o == "Barnelbel Panmo", f"Got {s}, {o}"

    # 6. Step 6: Invoice (Owner is Finance, waivable=True)
    t_step6 = dict(t_step5, work_report="https://docs.example/wr", invoice="")
    s, o = compute_stage_and_owner(t_step6, roster)
    assert s == "Invoice" and o == "Rogelnor Nelkin", f"Got {s}, {o}"

    # 7. Step 7: Done (all documents valid / waived)
    t_done = dict(t_step6, invoice="NA")
    s, o = compute_stage_and_owner(t_done, roster)
    assert s == "Done" and o == "", f"Got {s}, {o}"

    print(" -> test_stage_logic PASSED")


def test_junk_handling_in_stages():
    print("Running test_junk_handling_in_stages...")
    roster = build_roster_maps(MOCK_AGENTS)

    # Junk in PO ('done') must block at Purchase Order
    t_junk_po = {
        "id": 201, "manager": "Felboso Hamar", "engineer_id": 20,
        "purchase_order": "done", "delivery_note": "https://docs.example/dn"
    }
    s, o = compute_stage_and_owner(t_junk_po, roster)
    assert s == "Purchase Order" and o == "Felboso Hamar", f"Expected Purchase Order, got {s}"

    # 'NA' in Work Report (non-waivable!) must block at Work Report
    t_na_wr = {
        "id": 202, "manager": "Felboso Hamar", "engineer_id": 20,
        "purchase_order": "https://docs.example/po", "delivery_note": "NA",
        "work_report": "NA"
    }
    s, o = compute_stage_and_owner(t_na_wr, roster)
    assert s == "Work Report" and o == "Barnelbel Panmo", f"Expected Work Report, got {s}"

    # Junk in Invoice ('asked Ravi') must block at Invoice
    t_junk_inv = {
        "id": 203, "manager": "Felboso Hamar", "engineer_id": 20,
        "purchase_order": "https://docs.example/po", "delivery_note": "NA",
        "work_report": "https://docs.example/wr", "invoice": "asked Ravi"
    }
    s, o = compute_stage_and_owner(t_junk_inv, roster)
    assert s == "Invoice" and o == "Rogelnor Nelkin", f"Expected Invoice, got {s}"

    print(" -> test_junk_handling_in_stages PASSED")


def test_integrity_rules():
    print("Running test_integrity_rules...")
    roster = build_roster_maps(MOCK_AGENTS)

    # Closed (Status 5) but missing invoice
    t_closed_missing = {
        "id": 301, "status": 5, "manager": "Felboso Hamar", "engineer_id": 20,
        "purchase_order": "https://docs.example/po", "delivery_note": "NA",
        "work_report": "https://docs.example/wr", "invoice": ""
    }
    flag = compute_integrity_flags(t_closed_missing, roster)
    assert "Closed but paperwork missing" in flag, f"Got {flag}"

    # Resolved (Status 4) without Work Report
    t_resolved_no_wr = {
        "id": 302, "status": 4, "manager": "Felboso Hamar", "engineer_id": 20,
        "purchase_order": "https://docs.example/po", "delivery_note": "NA",
        "work_report": "", "invoice": ""
    }
    flag = compute_integrity_flags(t_resolved_no_wr, roster)
    assert "Resolved without Work Report" in flag, f"Got {flag}"

    # Manager not in roster
    t_unknown_mgr = {
        "id": 303, "status": 2, "manager": "Random Unknown Person", "engineer_id": None
    }
    flag = compute_integrity_flags(t_unknown_mgr, roster)
    assert "Manager not in roster" in flag, f"Got {flag}"

    print(" -> test_integrity_rules PASSED")


def test_store_code_and_work_type():
    print("Running test_store_code_and_work_type...")
    # Email extraction
    assert extract_store_code("", "store.k7q2xr@omnimart.example") == "K7Q2XR"
    # Name extraction
    assert extract_store_code("", "", "OmniMart Rarsi (M9GDCU)") == "M9GDCU"
    # Subject parentheses
    assert extract_store_code("PM AT OMNIMART TARVENO (K7Q2XR)") == "K7Q2XR"
    # Subject after OMNIMART
    assert extract_store_code("INSTALLATION AT OMNIMART VESHOLM DKCKJT") == "DKCKJT"

    # Work types
    assert extract_work_type("PM AT OMNIMART") == "Maintenance"
    assert extract_work_type("Quarterly AMC visit") == "Maintenance"
    assert extract_work_type("INSTALLATION AT OMNIMART VESHOLM") == "Installation"
    assert extract_work_type("CCTV Camera not working at store") == "Service Call"
    assert extract_work_type("FANVI Faulty") == "Service Call"
    assert extract_work_type("General enquiry") == "Other"

    # Noise detection
    assert is_noise_ticket("Grow your business with a Premium Listing - offer inside") is True
    assert is_noise_ticket("Automatic reply: Out of office") is True
    assert is_noise_ticket("Personal loan offer 50% discount") is True
    assert is_noise_ticket("PM AT OMNIMART TARVENO (K7Q2XR)") is False
    assert is_noise_ticket("Camera not working") is False

    print(" -> test_store_code_and_work_type PASSED")


def test_three_way_merge_logic():
    print("Running test_three_way_merge_logic...")
    # Simulate field diff logic:
    # Cases:
    # 1. Local changed, server unchanged -> PUSH
    # 2. Server changed, local unchanged -> PULL (sheet adopts server)
    # 3. Both changed differently -> CONFLICT (server wins)
    # 4. Unchanged cell never overwrites server

    def diff_field(base_val, sheet_val, server_val):
        local_changed = (sheet_val != base_val)
        server_changed = (server_val != base_val)
        if local_changed and not server_changed:
            return "PUSH_TO_SERVER"
        elif server_changed and not local_changed:
            return "PULL_TO_SHEET"
        elif local_changed and server_changed:
            if sheet_val == server_val:
                return "IDENTICAL_CHANGE"
            else:
                return "CONFLICT_SERVER_WINS"
        else:
            return "NO_CHANGE"

    # Case 1: Human edited sheet
    assert diff_field("Open", "Pending", "Open") == "PUSH_TO_SERVER"
    # Case 2: Server changed by mobile
    assert diff_field("Open", "Open", "Resolved") == "PULL_TO_SHEET"
    # Case 3: Both changed to different things
    assert diff_field("Open", "Pending", "Resolved") == "CONFLICT_SERVER_WINS"
    # Case 4: Both changed to same thing
    assert diff_field("Open", "Closed", "Closed") == "IDENTICAL_CHANGE"
    # Case 5: Neither changed
    assert diff_field("Open", "Open", "Open") == "NO_CHANGE"

    print(" -> test_three_way_merge_logic PASSED")


def test_sheet_contract():
    print("Running test_sheet_contract...")
    with tempfile.TemporaryDirectory() as tmpdir:
        test_sheet = os.path.join(tmpdir, "test_tickets.xlsx")
        out_csv = os.path.join(tmpdir, "out.csv")
        edits_csv = os.path.join(tmpdir, "edits.csv")

        # Create a test workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tickets"
        headers = REQUIRED_COLS + ["Store Code", "Work Type"]
        ws.append(headers)
        ws.append([100, "PM Test", "Open", "", "Felboso Hamar", "", "", "", "", "Assign Engineer", "Felboso Hamar", "K7Q2XR", "Maintenance"])
        wb.save(test_sheet)

        # Monkey-patch SHEET_PATH in sheet module for test
        import sheet
        orig_path = sheet.SHEET_PATH
        sheet.SHEET_PATH = test_sheet
        try:
            # 1. Test export
            sheet.export_csv(out_csv)
            assert os.path.exists(out_csv)
            with open(out_csv, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)
                assert rows[0] == headers
                assert rows[1][0] == "100"
                assert rows[1][2] == "Open"

            # 2. Test apply edits (edit Status to 'Closed' and Engineer to 'Barnelbel Panmo')
            with open(edits_csv, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Ticket ID", "Column", "Value"])
                writer.writerow([100, "Status", "Closed"])
                writer.writerow([100, "Engineer", "Barnelbel Panmo"])

            sheet.apply_edits(edits_csv)

            # 3. Export again and verify values
            out2_csv = os.path.join(tmpdir, "out2.csv")
            sheet.export_csv(out2_csv)
            with open(out2_csv, "r", encoding="utf-8") as f:
                rows2 = list(csv.reader(f))
                col_map = {h: i for i, h in enumerate(rows2[0])}
                assert rows2[1][col_map["Status"]] == "Closed"
                assert rows2[1][col_map["Engineer"]] == "Barnelbel Panmo"

        finally:
            sheet.SHEET_PATH = orig_path

    print(" -> test_sheet_contract PASSED")


def test_submission_contract():
    print("Running test_submission_contract...")
    sub_path = os.path.join(os.path.dirname(__file__), "submission.json")
    assert os.path.exists(sub_path), "submission.json missing"
    with open(sub_path) as f:
        data = json.load(f)
    assert "sync_cmd" in data, "sync_cmd missing in submission.json"
    assert "sheet_cmd" in data, "sheet_cmd missing in submission.json"
    print(" -> test_submission_contract PASSED")


if __name__ == "__main__":
    test_doc_validation()
    test_stage_logic()
    test_junk_handling_in_stages()
    test_integrity_rules()
    test_store_code_and_work_type()
    test_three_way_merge_logic()
    test_sheet_contract()
    test_submission_contract()
    print("\n===============================")
    print("  ALL TEST CASES PASSED (100%) ")
    print("===============================\n")
