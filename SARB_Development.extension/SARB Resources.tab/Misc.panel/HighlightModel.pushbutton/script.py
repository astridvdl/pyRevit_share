# -*- coding: utf-8 -*-
from __future__ import print_function

from pyrevit import script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ParameterElement,
    BuiltInCategory,
    CategoryType,
    ElementId,
    ParameterFilterElement,
    ElementParameterFilter,
    ParameterValueProvider,
    FilterStringEquals,
    FilterStringGreater,
    FilterStringRule,
    OverrideGraphicSettings,
    Color,
    FillPatternElement,
    FillPatternTarget,
    Transaction
)

from System.Collections.Generic import List

logger = script.get_logger()
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

PARAM_NAME = "DT Data Transferred"

FILTER_FILLED_NAME = "DT Data Transferred - FILLED (Green)"
FILTER_BLANK_NAME  = "DT Data Transferred - BLANK (Red)"

def get_solid_fill_pattern_id(document):
    """Get Solid fill pattern ElementId."""
    fps = FilteredElementCollector(document).OfClass(FillPatternElement).ToElements()
    for fp in fps:
        try:
            pat = fp.GetFillPattern()
            if pat and pat.Target == FillPatternTarget.Drafting and pat.IsSolidFill:
                return fp.Id
        except:
            pass
    return ElementId.InvalidElementId

def get_parameter_element_by_name(document, pname):
    """Find a project/shared parameter (ParameterElement) by name."""
    pelms = FilteredElementCollector(document).OfClass(ParameterElement).ToElements()
    for pe in pelms:
        try:
            if pe.Name == pname:
                return pe
        except:
            pass
    return None

def get_filterable_category_ids(document):
    """
    Return a broad set of model categories that can take bound parameters.
    This works well for project parameters and avoids missing categories.
    """
    cat_ids = List[ElementId]()
    cats = document.Settings.Categories
    for c in cats:
        try:
            if c.CategoryType == CategoryType.Model and c.AllowsBoundParameters:
                cat_ids.Add(c.Id)
        except:
            pass
    return cat_ids

def get_existing_filter_id_by_name(document, name):
    """Return existing ParameterFilterElement id by name, else InvalidElementId."""
    existing = FilteredElementCollector(document).OfClass(ParameterFilterElement).ToElements()
    for f in existing:
        try:
            if f.Name == name:
                return f.Id
        except:
            pass
    return ElementId.InvalidElementId

def ensure_filter(document, name, categories, element_filter):
    """Create filter if it doesn't exist; otherwise return existing id."""
    fid = get_existing_filter_id_by_name(document, name)
    if fid != ElementId.InvalidElementId:
        return fid

    # Revit 2023-safe: create with an ElementFilter
    return ParameterFilterElement.Create(document, name, categories, element_filter).Id

def build_string_rule_filters(param_id):
    """
    IronPython-safe (3-arg FilterStringRule overload):
    - blank: param == ""
    - filled: param > ""
    """
    pvp = ParameterValueProvider(param_id)

    blank_rule = FilterStringRule(pvp, FilterStringEquals(), "")
    filled_rule = FilterStringRule(pvp, FilterStringGreater(), "")

    blank_filter = ElementParameterFilter(blank_rule)
    filled_filter = ElementParameterFilter(filled_rule)

    return filled_filter, blank_filter

def set_filter_overrides(v, filter_id, rgb_tuple, solid_pattern_id):
    ogs = OverrideGraphicSettings()
    col = Color(rgb_tuple[0], rgb_tuple[1], rgb_tuple[2])

    # Projection fill + line color (helps visibility)
    ogs.SetProjectionLineColor(col)
    if solid_pattern_id != ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternId(solid_pattern_id)
        ogs.SetSurfaceForegroundPatternColor(col)

    v.SetFilterOverrides(filter_id, ogs)

def add_filter_to_view_if_missing(v, filter_id):
    existing_ids = list(v.GetFilters())
    if filter_id not in existing_ids:
        v.AddFilter(filter_id)

def remove_all_other_filters(v, keep_filter_ids):
    """
    Removes all filters from the view except those in keep_filter_ids.
    """
    existing_ids = list(v.GetFilters())

    for fid in existing_ids:
        if fid not in keep_filter_ids:
            try:
                v.RemoveFilter(fid)
            except:
                pass

def main():
    # Basic view safety
    if view is None or view.IsTemplate:
        forms.alert("Active view is a template or invalid. Open a normal view and try again.", exitscript=True)

    # Find parameter
    pe = get_parameter_element_by_name(doc, PARAM_NAME)
    if not pe:
        forms.alert(
            "Could not find a project/shared parameter named:\n\n  '{}'\n\n"
            "Make sure it exists as a Project Parameter / Shared Parameter in this model."
            .format(PARAM_NAME),
            exitscript=True
        )

    param_id = pe.Id
    cats = get_filterable_category_ids(doc)
    if cats.Count == 0:
        forms.alert("No filterable model categories found.", exitscript=True)

    filled_filter, blank_filter = build_string_rule_filters(param_id)
    solid_id = get_solid_fill_pattern_id(doc)

    t = Transaction(doc, "Add DT Data Transferred filled/blank filters")
    t.Start()
    try:
        fid_filled = ensure_filter(doc, FILTER_FILLED_NAME, cats, filled_filter)
        fid_blank  = ensure_filter(doc, FILTER_BLANK_NAME,  cats, blank_filter)

        # Remove all other filters first
        remove_all_other_filters(view, [fid_filled, fid_blank])
        
        # Add to view
        add_filter_to_view_if_missing(view, fid_filled)
        add_filter_to_view_if_missing(view, fid_blank)

        # Ensure visibility
        view.SetFilterVisibility(fid_filled, True)
        view.SetFilterVisibility(fid_blank, True)

        # Apply overrides: GREEN for filled, RED for blank
        set_filter_overrides(view, fid_filled, (0, 200, 0), solid_id)
        set_filter_overrides(view, fid_blank,  (200, 0, 0), solid_id)

        t.Commit()

        forms.toast("DT filters applied to active view:\n- Filled (Green)\n- Blank (Red)")
    except Exception as ex:
        logger.exception("Failed to create/apply filters.")
        t.RollBack()
        forms.alert("Failed to create/apply filters:\n\n{}".format(ex), exitscript=False)

if __name__ == "__main__":
    main()