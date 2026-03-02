# -*- coding: utf-8 -*-
# pyRevit | Parameter Governance (Case-Insensitive) — Revit 2024+
# - Compare project parameters to parameters.json (case-insensitive)
# - Report counts, highlight duplicates (by name ignoring case), optional duplicate cleanup (keep one)
# - Add missing names (Text | IFC group) as INSTANCE bound to ALL model categories
# - Reuses existing definitions in shared param file to avoid "already present" errors

import os, json, clr, traceback, tempfile

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import (
    Definition, ElementBinding, InstanceBinding, CategorySet, Category, Categories,
    BindingMap, DefinitionBindingMapIterator, ParameterElement, FilteredElementCollector,
    Transaction, BuiltInCategory, LabelUtils, ForgeTypeId, CategoryType
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from System.Windows.Forms import OpenFileDialog, DialogResult

# ----------------- Revit context -----------------
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None
if not doc:
    print("No active document."); raise SystemExit

app = __revit__.Application

# ----------------- Helpers -----------------
def norm(s):
    return (s or "").strip().lower()

def safe(call, default=None):
    try: return call()
    except: return default

def get_ifc_group_id():
    try:
        from Autodesk.Revit.DB import GroupTypeId
        return GroupTypeId.IFC
    except:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        return BuiltInParameterGroup.PG_IFC

def get_text_spec_id():
    try:
        from Autodesk.Revit.DB import SpecTypeId
        return getattr(SpecTypeId.String, "Text", SpecTypeId.String)
    except:
        return ForgeTypeId("autodesk.spec.aec:string-1.0.0")

def build_categoryset_all_model(doc):
    cs = CategorySet()
    for c in doc.Settings.Categories:
        try:
            if c and c.CategoryType == CategoryType.Model and c.AllowsBoundParameters:
                cs.Insert(c)
        except:
            pass
    return cs

def ask_yes_no(title, main_instruction, content=None):
    td = TaskDialog(title)
    td.MainInstruction = main_instruction
    if content: td.MainContent = content
    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    return td.Show() == TaskDialogResult.Yes

def rescan_bindingmap_casefold(doc):
    """Return dict[lower_name] -> [(Definition, original_name), ...] for project parameters."""
    map_ci = {}
    it = doc.ParameterBindings.ForwardIterator()
    while it.MoveNext():
        defn = it.Key
        bind = it.Current
        if isinstance(defn, Definition) and isinstance(bind, ElementBinding):
            orig = (defn.Name or "").strip()
            if not orig: continue
            map_ci.setdefault(norm(orig), []).append((defn, orig))
    return map_ci

def get_or_create_ext_def(sp_file, group_name, name, spec_id):
    """Reuse an existing ExternalDefinition by name (any group), else create under group_name."""
    # Try reuse
    for g in sp_file.Groups:
        try:
            d = g.Definitions.get_Item(name)
            if d:
                try:
                    if hasattr(d, "GetDataType") and d.GetDataType() != spec_id:
                        print('⚠️  Definition "{}" exists with different spec; reusing anyway.'.format(name))
                except:
                    pass
                return d
        except:
            pass
    # Create under target group
    target_group = None
    for g in sp_file.Groups:
        if g.Name == group_name:
            target_group = g; break
    if target_group is None:
        target_group = sp_file.Groups.Create(group_name)
    from Autodesk.Revit.DB import ExternalDefinitionCreationOptions
    opt = ExternalDefinitionCreationOptions(name, spec_id)
    return target_group.Definitions.Create(opt)

# ----------------- Load JSON -----------------
script_dir = os.path.dirname(__file__)
#json_path  = os.path.join(script_dir, "parameters.json")


# --- choose JSON file via file picker (default: parameters.json next to script) ---
default_json = os.path.join(script_dir, "parameters.json")

ofd = OpenFileDialog()
ofd.Title = "Select parameter JSON file"
ofd.InitialDirectory = script_dir
ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
ofd.FileName = "parameters.json"   # pre-fill

result = ofd.ShowDialog()

if result == DialogResult.OK and ofd.FileName:
    json_path = ofd.FileName
else:
    # If user cancels, fall back to default parameters.json next to script
    json_path = default_json

if not os.path.exists(json_path):
    TaskDialog.Show("Parameter Governance", "parameters.json not found in:\n{}".format(script_dir))
    raise SystemExit

print("Using parameter list from:\n{}".format(json_path))

try:
    with open(json_path, "r") as f:
        raw = json.load(f)
except Exception as ex:
    print("Error reading JSON:", ex); traceback.print_exc()
    TaskDialog.Show("Parameter Governance", "Error reading JSON. See console for details.")
    raise SystemExit

# Canonical (display) names from JSON + lowercased set for matching
target_display_names = [ (p.get("name","").strip()) for p in raw if isinstance(p, dict) and p.get("name")]
target_display_names = [n for n in target_display_names if n]
target_ci = [norm(n) for n in target_display_names]
target_ci_set = set(target_ci)

print("=== Parameter Governance (case-insensitive) ===")
print("Loaded {} target names from parameters.json".format(len(target_display_names)))

# ----------------- Scan existing Project Parameters (case-insensitive) -----------------
ci_map = rescan_bindingmap_casefold(doc)  # lower_name -> [(def, orig_name), ...]

print("\n--- Current counts (Project Parameters; case-insensitive) ---")
missing_ci, duplicate_ci = [], []
for disp_name, nm_ci in zip(target_display_names, target_ci):
    cnt = len(ci_map.get(nm_ci, []))
    flag = ""
    if cnt == 0:
        missing_ci.append(nm_ci); flag = "  <-- MISSING"
    elif cnt > 1:
        duplicate_ci.append(nm_ci); flag = "  <-- DUPLICATE ({} entries)".format(cnt)
    print(u"{:<45} {:>3}{}".format(disp_name, cnt, flag))

print("\nSummary: {} missing; {} duplicate names (case-insensitive).".format(len(missing_ci), len(duplicate_ci)))

# ----------------- Duplicate cleanup (keep one; case-insensitive) -----------------
if duplicate_ci:
    detail = "Duplicate parameter names (ignoring case) were found.\nRemove duplicates (keep one per name)?"
    if ask_yes_no("Duplicate Cleanup", "Remove duplicate project parameter bindings?", detail):
        bmap = doc.ParameterBindings
        t = Transaction(doc, "Remove duplicate project parameter bindings (case-insensitive)")
        t.Start()
        removed = 0
        for nm_ci in duplicate_ci:
            defs = ci_map.get(nm_ci, [])
            # keep the first one, remove the rest
            for defn, _orig in defs[1:]:
                try:
                    if bmap.Remove(defn):
                        removed += 1
                except Exception as ex:
                    print("Error removing duplicate '{}': {}".format(_orig, ex))
        t.Commit()

        # Re-scan after removal
        ci_map = rescan_bindingmap_casefold(doc)
        still_dupes = [nm for nm in duplicate_ci if len(ci_map.get(nm, [])) > 1]
        print("\nDuplicate cleanup removed {} bindings.".format(removed))
        if still_dupes:
            print("Still duplicated (manual review):")
            for nm in still_dupes:
                print(" - {} ({} entries remain)".format(nm, len(ci_map.get(nm, []))))
        else:
            print("All duplicate names reduced to a single binding each.")
    else:
        print("Skipped duplicate cleanup by user choice.")

# Refresh missing after potential cleanup
ci_map = rescan_bindingmap_casefold(doc)
missing_ci = [nm for nm in target_ci if len(ci_map.get(nm, [])) == 0]

# ----------------- Missing -> Add as Text under IFC group, bound to ALL model cats (Instance) -----------------
if missing_ci:
    # for user-facing dialog, map back to display names (first match by ci)
    ci_to_disp = {}
    for disp_name in target_display_names:
        ci_to_disp.setdefault(norm(disp_name), disp_name)

    miss_disp = [ci_to_disp.get(nm, nm) for nm in missing_ci]
    detail = ("Missing parameters (case-insensitive) can be added as Text under the IFC group.\n"
              "They will be bound as INSTANCE to ALL model categories that allow bound parameters.\n\n"
              "Missing:\n- " + "\n- ".join(miss_disp))
    if ask_yes_no("Add Missing Parameters", "Add missing names as Text (IFC) to ALL model categories?", detail):
        ifc_group = get_ifc_group_id()
        text_spec = get_text_spec_id()
        all_model_catset = build_categoryset_all_model(doc)

        # Use a temp Shared Parameter file to create/reuse External Definitions
        sp_temp = os.path.join(tempfile.gettempdir(), "pyrevit_param_temp.shrd.txt")
        try:
            if not os.path.exists(sp_temp):
                with open(sp_temp, "w") as _tmp:
                    _tmp.write("# pyRevit temp shared parameter file\n*GROUP\tpyRevitTemp\n")
            old_spf = app.SharedParametersFilename
            app.SharedParametersFilename = sp_temp
            sp_file = app.OpenSharedParameterFile()
            if sp_file is None:
                raise Exception("OpenSharedParameterFile() returned None.")

            t_add = Transaction(doc, "Add missing project parameters (all model cats, instance)")
            t_add.Start()

            created = []
            for nm_ci in missing_ci:
                # Use display name from JSON (preserve original casing for creation)
                disp = ci_to_disp.get(nm_ci, nm_ci)

                # if appeared meanwhile, skip
                if len(rescan_bindingmap_casefold(doc).get(nm_ci, [])) > 0:
                    continue
                try:
                    ext_def = get_or_create_ext_def(sp_file, "pyRevitTemp", disp, text_spec)

                    # Instance binding to ALL model categories
                    binding = doc.Application.Create.NewInstanceBinding(all_model_catset)

                    # Insert with IFC group (ForgeTypeId signature on 2024+)
                    inserted = False
                    try:
                        inserted = doc.ParameterBindings.Insert(ext_def, binding, ifc_group)
                    except:
                        inserted = doc.ParameterBindings.Insert(ext_def, binding, ifc_group)

                    if inserted:
                        created.append(disp)
                        print('Added "{}" (Text | IFC group | Instance | ALL model categories)'.format(disp))
                    else:
                        print('Failed to insert binding for "{}"'.format(disp))
                except Exception as ex:
                    print('Error creating "{}": {}'.format(disp, ex))

            t_add.Commit()
            app.SharedParametersFilename = old_spf

            TaskDialog.Show("Add Missing Parameters",
                            "Requested: {}\nCreated: {}\nSee console for details."
                            .format(len(miss_disp), len(created)))
        except Exception as ex:
            print("Error while adding missing parameters:", ex)
            traceback.print_exc()
            TaskDialog.Show("Add Missing Parameters", "Error while adding parameters. See console for details.")
    else:
        print("Skipped adding missing parameters by user choice.")
else:
    print("No missing names; model already contains all JSON parameters (case-insensitive match).")

print("\nDone.")
