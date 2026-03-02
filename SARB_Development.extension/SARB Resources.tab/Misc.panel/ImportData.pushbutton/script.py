# -*- coding: utf-8 -*-
from __future__ import print_function

import csv
import os
import datetime as dt
import traceback

from pyrevit import forms, script
from Autodesk.Revit.DB import (
    Transaction, ElementId, StorageType,
    BuiltInParameterGroup,
    SpecTypeId,
    CategoryType,
    ExternalDefinitionCreationOptions,
    SharedParameterElement
)

logger = script.get_logger()
output = script.get_output()  # only for update_progress

REQUIRED_HEADERS = [
    "Match_Type", "Confidence", "GlobalId",
    "ElementID_A", "ElementID_B",
    "IfcClass_A", "IfcClass_B",
    "Parameter", "Value_B"
]

# ---- tuning knobs ----
REPORT_EVERY = 1000
MAX_FAIL_PRINTS = 300
# ----------------------

# ---- DT stamp ----
DT_PARAM_NAME = "DT Data Transferred"
DT_SHARED_GROUP = "DT"
# ------------------


def now_str():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_str(x):
    if x is None:
        return ""
    try:
        return str(x)
    except Exception:
        return ""


def parse_int(s):
    s = safe_str(s).strip()
    if s == "":
        raise ValueError("Blank ElementID_A")
    return int(float(s))  # allow "123.0"


def parse_float(s):
    s = safe_str(s).strip().replace(",", ".")
    if s == "":
        raise ValueError("Blank numeric value")
    return float(s)


def convert_for_param(param, value_raw):
    st = param.StorageType
    s = safe_str(value_raw).strip()

    if st == StorageType.String:
        return s

    if st == StorageType.Integer:
        low = s.lower()
        if low in ("true", "yes", "y", "1"):
            return 1
        if low in ("false", "no", "n", "0"):
            return 0
        return int(float(s))

    if st == StorageType.Double:
        return parse_float(s)

    if st == StorageType.ElementId:
        return ElementId(parse_int(s))

    return s


def set_param(param, converted):
    st = param.StorageType
    if st == StorageType.String:
        return param.Set("" if converted is None else safe_str(converted))
    if st == StorageType.Integer:
        return param.Set(int(converted))
    if st == StorageType.Double:
        return param.Set(float(converted))
    if st == StorageType.ElementId:
        return param.Set(converted)
    return param.Set(safe_str(converted))


def read_csv_dicts(path):
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    text = raw.decode("utf-8", errors="replace").splitlines(True)
    reader = csv.DictReader(text)
    headers = reader.fieldnames or []
    return headers, list(reader)


def heartbeat(i, total, stats, touched_count):
    failed = stats["RowsAttempted"] - stats["RowsSuccess"] - stats["RowsSkippedBlankValue"]
    print("[{}/{}] OK={} FAIL={} SKIP={} Unique={}".format(
        i, total,
        stats["RowsSuccess"],
        failed,
        stats["RowsSkippedBlankValue"],
        touched_count
    ))


def print_fail(fail_prints, match_type, confidence, globalid, eid, param_name, value_b, reason):
    if MAX_FAIL_PRINTS is not None and fail_prints >= MAX_FAIL_PRINTS:
        return fail_prints

    vb = safe_str(value_b)
    if len(vb) > 160:
        vb = vb[:157] + "..."

    print("FAIL | EID={} | Param='{}' | Value='{}' | Match={} Conf={} GlobalId={} | Reason={}".format(
        safe_str(eid),
        safe_str(param_name),
        vb,
        safe_str(match_type),
        safe_str(confidence),
        safe_str(globalid),
        safe_str(reason)
    ))
    return fail_prints + 1


def ensure_dt_shared_text_parameter(doc):
    """
    Revit 2023:
    Ensures DT Data Transferred exists as a SHARED instance TEXT parameter,
    bound to all model categories that allow bound parameters.
    Returns SharedParameterElement.Id (usable for filters, if ever needed).
    """
    app = doc.Application

    sp_path = app.SharedParametersFilename
    if (not sp_path) or (not os.path.exists(sp_path)):
        sp_path = forms.pick_file(file_ext="txt", title="Select Shared Parameters .txt file (required)")
        if not sp_path:
            forms.alert("Shared Parameters file is required to create '{}'.".format(DT_PARAM_NAME), exitscript=True)
        app.SharedParametersFilename = sp_path

    sp_file = app.OpenSharedParameterFile()
    if sp_file is None:
        forms.alert("Could not open Shared Parameters file:\n{}".format(app.SharedParametersFilename), exitscript=True)

    # group
    sp_group = None
    for g in sp_file.Groups:
        if g.Name == DT_SHARED_GROUP:
            sp_group = g
            break
    if sp_group is None:
        sp_group = sp_file.Groups.Create(DT_SHARED_GROUP)

    # definition
    definition = None
    for d in sp_group.Definitions:
        if d.Name == DT_PARAM_NAME:
            definition = d
            break

    if definition is None:
        # Revit 2023: ForgeTypeId-based text spec
        opts = ExternalDefinitionCreationOptions(DT_PARAM_NAME, SpecTypeId.String.Text)
        definition = sp_group.Definitions.Create(opts)

    guid = definition.GUID

    # build category set (all model categories that can take bound params)
    catset = app.Create.NewCategorySet()
    for c in doc.Settings.Categories:
        try:
            if c.CategoryType != CategoryType.Model:
                continue
            if not c.AllowsBoundParameters:
                continue
            catset.Insert(c)
        except Exception:
            pass

    binding = app.Create.NewInstanceBinding(catset)

    # Check if already bound; if not, bind it
    bmap = doc.ParameterBindings
    already_bound = False
    it = bmap.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        try:
            k = it.Key
            if k and k.Name == DT_PARAM_NAME:
                already_bound = True
                break
        except Exception:
            pass

    if not already_bound:
        bmap.Insert(definition, binding, BuiltInParameterGroup.PG_DATA)

    # Resolve SharedParameterElement
    sp_elem = SharedParameterElement.Lookup(doc, guid)
    if sp_elem is None:
        doc.Regenerate()
        sp_elem = SharedParameterElement.Lookup(doc, guid)
    if sp_elem is None:
        forms.alert("Could not resolve SharedParameterElement for '{}'.".format(DT_PARAM_NAME), exitscript=True)

    return sp_elem.Id


def main():
    doc = __revit__.ActiveUIDocument.Document  # noqa

    csv_path = forms.pick_file(file_ext="csv", title="Select __08_Params_To_Write.csv (payload)")
    if not csv_path:
        return

    headers, rows = read_csv_dicts(csv_path)
    missing = [h for h in REQUIRED_HEADERS if h not in headers]
    if missing:
        forms.alert(
            "CSV is missing required columns:\n{}\n\nFile:\n{}".format(", ".join(missing), csv_path),
            exitscript=True
        )

    confirm = forms.alert(
        "Import {} rows into the OPEN Revit model using ElementID_A?\n\n"
        "CSV:\n{}\n\nDocument:\n{}\n\n"
        "Also will:\n"
        "- Ensure shared parameter '{}' exists\n"
        "- Stamp timestamp on successful writes\n\n"
        "Console output:\n- rolling counter every {} rows\n- failures only".format(
            len(rows), csv_path, doc.Title, DT_PARAM_NAME, REPORT_EVERY
        ),
        yes=True, no=True
    )
    if not confirm:
        return

    print("=== IFC Compare CSV Import ===")
    print("Run: {}".format(now_str()))
    print("Document: {}".format(safe_str(doc.Title)))
    print("CSV: {}".format(csv_path))
    print("Rows: {}".format(len(rows)))
    print("Stamp param: '{}'".format(DT_PARAM_NAME))
    print("Reporting: failures only, heartbeat every {} rows".format(REPORT_EVERY))
    print("================================")

    stats = {
        "TotalRows": len(rows),
        "RowsAttempted": 0,
        "RowsSuccess": 0,
        "RowsStampedDT": 0,
        "RowsSkippedBlankValue": 0,
        "RowsFailedBadElementId": 0,
        "RowsFailedElementNotFound": 0,
        "RowsFailedParamNotFound": 0,
        "RowsFailedParamReadOnly": 0,
        "RowsFailedParse": 0,
        "RowsFailedSet": 0,
        "RowsFailedStampDT": 0
    }
    touched = set()
    fail_prints = 0

    t = Transaction(doc, "IFC Compare Import (CSV)")
    t.Start()

    try:
        # Ensure stamp parameter exists
        ensure_dt_shared_text_parameter(doc)

        total = len(rows)
        for i, r in enumerate(rows, start=1):
            stats["RowsAttempted"] += 1

            # progress bar (safe to try; ignore if unsupported)
            try:
                output.update_progress(i, total)
            except Exception:
                pass

            if (i % REPORT_EVERY) == 0:
                heartbeat(i, total, stats, len(touched))

            match_type = safe_str(r.get("Match_Type"))
            confidence = safe_str(r.get("Confidence"))
            globalid = safe_str(r.get("GlobalId"))
            elementid_a_raw = r.get("ElementID_A")
            param_name = safe_str(r.get("Parameter")).strip()
            value_b = r.get("Value_B")

            # Skip blank value
            if value_b is None or safe_str(value_b).strip() == "":
                stats["RowsSkippedBlankValue"] += 1
                continue

            # Parse ElementId
            try:
                eid_int = parse_int(elementid_a_raw)
            except Exception as ex:
                stats["RowsFailedBadElementId"] += 1
                fail_prints = print_fail(
                    fail_prints, match_type, confidence, globalid,
                    elementid_a_raw, param_name, value_b,
                    "Bad ElementID_A: {}".format(ex)
                )
                continue

            el = doc.GetElement(ElementId(eid_int))
            if el is None:
                stats["RowsFailedElementNotFound"] += 1
                fail_prints = print_fail(
                    fail_prints, match_type, confidence, globalid,
                    eid_int, param_name, value_b,
                    "Element not found"
                )
                continue

            touched.add(eid_int)

            p = el.LookupParameter(param_name)
            if p is None:
                stats["RowsFailedParamNotFound"] += 1
                fail_prints = print_fail(
                    fail_prints, match_type, confidence, globalid,
                    eid_int, param_name, value_b,
                    "Parameter not found"
                )
                continue

            if p.IsReadOnly:
                stats["RowsFailedParamReadOnly"] += 1
                fail_prints = print_fail(
                    fail_prints, match_type, confidence, globalid,
                    eid_int, param_name, value_b,
                    "Parameter is read-only"
                )
                continue

            # Convert + set
            try:
                converted = convert_for_param(p, value_b)
            except Exception as ex:
                stats["RowsFailedParse"] += 1
                fail_prints = print_fail(
                    fail_prints, match_type, confidence, globalid,
                    eid_int, param_name, value_b,
                    "Parse error: {}".format(ex)
                )
                continue

            try:
                set_param(p, converted)
                stats["RowsSuccess"] += 1

                # Stamp DT on success
                try:
                    p_dt = el.LookupParameter(DT_PARAM_NAME)
                    if p_dt is not None and (not p_dt.IsReadOnly):
                        p_dt.Set(now_str())
                        stats["RowsStampedDT"] += 1
                    else:
                        stats["RowsFailedStampDT"] += 1
                except Exception:
                    stats["RowsFailedStampDT"] += 1

            except Exception as ex:
                stats["RowsFailedSet"] += 1
                fail_prints = print_fail(
                    fail_prints, match_type, confidence, globalid,
                    eid_int, param_name, value_b,
                    "Set failed: {}".format(ex)
                )

        t.Commit()

    except Exception as e:
        try:
            t.RollBack()
        except Exception:
            pass
        forms.alert(
            "Import crashed:\n\n{}\n\n{}".format(e, traceback.format_exc()),
            exitscript=True
        )

    failed = stats["RowsAttempted"] - stats["RowsSuccess"] - stats["RowsSkippedBlankValue"]

    print("=== DONE ===")
    heartbeat(stats["RowsAttempted"], stats["TotalRows"], stats, len(touched))
    print("Stamped DT rows: {} (stamp failures: {})".format(stats["RowsStampedDT"], stats["RowsFailedStampDT"]))
    print("Breakdown: BadEID={} NotFound={} ParamMissing={} ReadOnly={} ParseErr={} SetFail={}".format(
        stats["RowsFailedBadElementId"],
        stats["RowsFailedElementNotFound"],
        stats["RowsFailedParamNotFound"],
        stats["RowsFailedParamReadOnly"],
        stats["RowsFailedParse"],
        stats["RowsFailedSet"]
    ))
    if MAX_FAIL_PRINTS is not None and fail_prints >= MAX_FAIL_PRINTS:
        print("NOTE: Printed first {} failures only (MAX_FAIL_PRINTS).".format(MAX_FAIL_PRINTS))

    forms.alert(
        "Import complete.\n\nOK: {}\nFailed: {}\nSkipped blank: {}\nUnique elements: {}\n\n"
        "DT stamped: {}\nDT stamp failed: {}\n\n"
        "See pyRevit output panel for failure lines.".format(
            stats["RowsSuccess"], failed, stats["RowsSkippedBlankValue"], len(touched),
            stats["RowsStampedDT"], stats["RowsFailedStampDT"]
        )
    )


if __name__ == "__main__":
    main()