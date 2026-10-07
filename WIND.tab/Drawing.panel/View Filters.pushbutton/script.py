from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import * 
import Functions as func
from pyrevit import revit, forms, script
import clr
import math
clr.AddReference("System")
from System import Int64
from System.Collections.Generic import List
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
doc     = __revit__.ActiveUIDocument.Document


def apply_filters_to_view(view):
    if view is None:
        raise ValueError("view cannot be None")

    def _norm(s):
        return " ".join(str(s).strip().split()).upper()

    target_filter_names = [
        " REBAR GRIDS",
       # " REBAR BOTTOM CONCENTRIC ",
       # " REBAR TOP CONCENTRIC ",
       # " REBAR BOTTOM RADIAL ",
       # " REBAR TOP RADIAL ",
        " REBAR STOOLS ",
        " REBAR HAIR PINS ",
        " REBAR PLINTH VERTICAL ",
        " REBAR PLINTH FACE CONCENTRICL ",
        " REBAR SLAB FACE CONCENTRIC ",
        " REBAR BASE VERTICAL ",
        " REBAR PLINTH CONCENTRIC ",
        " REBAR RADIAL ANCHOR ",
        " REBAR PLINTH FACE HORIZONTAL ",
        " SECTION BOTTOM RADIAL ",
        " SECTION TOP RADIAL ",
        " SECTION HAIR PINS ",
        " SECTION PLINTH VERTICAL ",
        " SECTION BASE VERTICAL ",
        " REBAR CONCENTRIC BRACING ",
        " REBAR SUPPORT BARS",
        " REBAR RADIATOR BASE",
    ]

    def enable_target_filters_on_view(target_view):
        all_parameter_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
        filters_by_name = {_norm(f.Name): f for f in all_parameter_filters}
        target_name_set = set(_norm(n) for n in target_filter_names)

        for applied_filter_id in list(target_view.GetFilters()):
            applied_filter_elem = doc.GetElement(applied_filter_id)
            if applied_filter_elem is None:
                continue

            if _norm(applied_filter_elem.Name) not in target_name_set:
                target_view.RemoveFilter(applied_filter_id)

        for raw_filter_name in target_filter_names:
            filter_name = " ".join(str(raw_filter_name).strip().split())
            filter_elem = filters_by_name.get(_norm(filter_name))
            if filter_elem is None:
                continue

            filter_id = filter_elem.Id
            if not target_view.IsFilterApplied(filter_id):
                target_view.AddFilter(filter_id)

    def set_filter_surface_overrides(
        target_view,
        filter_id,
        proj_fill_color=None,
        proj_fill_pattern_id=None,
        cut_fill_color=None,
        cut_fill_pattern_id=None,
    ):
        if filter_id is None:
            raise ValueError("filter_id cannot be None")

        if not target_view.IsFilterApplied(filter_id):
            raise ValueError(
                "Filter {} is not applied to view '{}'.".format(
                    filter_id.IntegerValue, target_view.Name
                )
            )

        ogs = target_view.GetFilterOverrides(filter_id)

        def _apply_ogs_setting(ogs_obj, method_names, value):
            if value is None:
                return ogs_obj

            for method_name in method_names:
                method = getattr(ogs_obj, method_name, None)
                if method is None:
                    continue

                result = method(value)
                if result is not None:
                    ogs_obj = result
                return ogs_obj

            raise AttributeError(
                "OverrideGraphicSettings does not support any of: {}".format(
                    ", ".join(method_names)
                )
            )

        ogs = _apply_ogs_setting(
            ogs,
            ["SetProjectionFillColor", "SetSurfaceForegroundPatternColor"],
            proj_fill_color,
        )
        ogs = _apply_ogs_setting(
            ogs,
            ["SetProjectionFillPatternId", "SetSurfaceForegroundPatternId"],
            proj_fill_pattern_id,
        )
        ogs = _apply_ogs_setting(
            ogs,
            ["SetCutFillColor", "SetCutForegroundPatternColor"],
            cut_fill_color,
        )
        ogs = _apply_ogs_setting(
            ogs,
            ["SetCutFillPatternId", "SetCutForegroundPatternId"],
            cut_fill_pattern_id,
        )

        return ogs

    filter_overrides = {
        # name                              proj_fill_color            cut_fill_color
        "REBAR GRIDS":                  {"proj_fill_color": Color(255, 128, 255), "cut_fill_color": Color(255, 128, 255)},
       # "REBAR BOTTOM CONCENTRIC":      {"proj_fill_color": Color(255, 128,   0), "cut_fill_color": Color(255, 128,   0)},
       # "REBAR TOP CONCENTRIC":         {"proj_fill_color": Color(255, 128,   0), "cut_fill_color": Color(255, 128,   0)},
       # "REBAR BOTTOM RADIAL":          {"proj_fill_color": Color(  0, 128,   0), "cut_fill_color": Color(  0, 128,   0)},
       # "REBAR TOP RADIAL":             {"proj_fill_color": Color(  0, 255,   0), "cut_fill_color": Color(  0, 255,   0)},
        "REBAR STOOLS":                 {"proj_fill_color": Color(  0, 255, 255), "cut_fill_color": Color(  0, 255, 255)},
        "REBAR HAIR PINS":              {"proj_fill_color": Color(255, 255,   0), "cut_fill_color": Color(255, 255,   0)},
        "REBAR PLINTH VERTICAL":        {"proj_fill_color": Color(  0, 128, 255), "cut_fill_color": Color(  0, 128, 255)},
        "REBAR PLINTH FACE CONCENTRICL":{"proj_fill_color": Color(255, 128, 255), "cut_fill_color": Color(255, 128, 255)},
        "REBAR SLAB FACE CONCENTRIC":   {"proj_fill_color": Color(255,   0,   0), "cut_fill_color": Color(255,   0,   0)},
        "REBAR BASE VERTICAL":          {"proj_fill_color": Color(  0, 128, 255), "cut_fill_color": Color(  0, 128, 255)},
        "REBAR PLINTH CONCENTRIC":      {"proj_fill_color": Color(255, 128,   0), "cut_fill_color": Color(255, 128,   0)},
        "REBAR RADIAL ANCHOR":          {"proj_fill_color": Color(255,   0,   0), "cut_fill_color": Color(255,   0,   0)},
        "REBAR PLINTH FACE HORIZONTAL": {"proj_fill_color": Color(255, 128,   0), "cut_fill_color": Color(255, 128,   0)},
        "SECTION BOTTOM RADIAL":        {"proj_fill_color": Color(  0, 128,   0), "cut_fill_color": Color(  0, 128,   0)},
        "SECTION TOP RADIAL":           {"proj_fill_color": Color(  0, 255,   0), "cut_fill_color": Color(  0, 255,   0)},
        "SECTION HAIR PINS":            {"proj_fill_color": Color(255, 255,   0), "cut_fill_color": Color(255, 255,   0)},
        "SECTION PLINTH VERTICAL":      {"proj_fill_color": Color(  0, 128, 255), "cut_fill_color": Color(  0, 128, 255)},
        "SECTION BASE VERTICAL":        {"proj_fill_color": Color(  0, 128, 255), "cut_fill_color": Color(  0, 128, 255)},
        "REBAR CONCENTRIC BRACING":     {"proj_fill_color": Color(128,   0, 128), "cut_fill_color": Color(128,   0, 128)},
        "REBAR SUPPORT BARS":           {"proj_fill_color": Color(145, 145,   0), "cut_fill_color": Color(145, 145,   0)},
        "REBAR RADIATOR BASE":          {"proj_fill_color": Color(255,   0, 128), "cut_fill_color": Color(255,   0, 128)},
    }

    enable_target_filters_on_view(view)

    all_filters = {
        doc.GetElement(fid).Name: doc.GetElement(fid).Id
        for fid in list(view.GetFilters())
        if doc.GetElement(fid) is not None
    }

    for raw_name in target_filter_names:
        clean_name = " ".join(str(raw_name).strip().split())
        lookup = _norm(clean_name)

        filter_id = next((fid for fname, fid in all_filters.items() if _norm(fname) == lookup), None)
        if filter_id is None:
            continue

        kwargs = next((v for k, v in filter_overrides.items() if _norm(k) == lookup), {})
        if not kwargs:
            continue

        ogs = set_filter_surface_overrides(view, filter_id, **kwargs)
        view.SetFilterOverrides(filter_id, ogs)


def set_solid_fill_pattern_on_view(view):
    solid_pattern_id = ElementId(4)

    def _try_set(ogs, method_names, value):
        for name in method_names:
            method = getattr(ogs, name, None)
            if method is None:
                continue
            result = method(value)
            return result if result is not None else ogs
        return ogs

    for fid in list(view.GetFilters()):
        ogs = view.GetFilterOverrides(fid)
        ogs = _try_set(ogs, ["SetSurfaceForegroundPatternId", "SetProjectionFillPatternId"], solid_pattern_id)
        ogs = _try_set(ogs, ["SetCutForegroundPatternId", "SetCutFillPatternId"], solid_pattern_id)
        view.SetFilterOverrides(fid, ogs)


def set_view_filter_visibility(view, view_visibility_dict):
    """Set filter visibility for a view based on view_visibility_dict with fallback for schedule-mark filters."""
    view_name = view.Name
    
    if view_name not in view_visibility_dict:
        return
    
    view_data = view_visibility_dict[view_name]
    filter_visibility_map = view_data.get('categories', {})
    
    fallback_map = {
        'TR': 'REBAR TOP RADIAL',
        'BR': 'REBAR BOTTOM RADIAL',
        'TC': 'REBAR TOP CONCENTRIC',
        'BC': 'REBAR BOTTOM CONCENTRIC',
    }
    
    for filter_id in list(view.GetFilters()):
        filter_elem = doc.GetElement(filter_id)
        if filter_elem is None:
            continue
        
        filter_name = str(filter_elem.Name or "").strip()
        visibility = None
        
        if filter_name in filter_visibility_map:
            visibility = filter_visibility_map[filter_name]
        else:
            name_upper = filter_name.upper()
            detected_prefix = None
            if "-" in name_upper:
                detected_prefix = name_upper.split("-", 1)[0].strip()
            for prefix, fallback_name in fallback_map.items():
                if detected_prefix == prefix or name_upper.startswith(prefix + " "):
                    if fallback_name in filter_visibility_map:
                        visibility = filter_visibility_map[fallback_name]
                    break
        
        if visibility is not None:
            view.SetFilterVisibility(filter_id, visibility)


view_ids = [
    446525, 446535, 446567, 453671, 2503522,
    453468, 453478, 2511791, 2511780, 2514065,
    4593885, 5905497, 5915879, 5998757, 9137627,
    603329, 603339, 966797, 3095802, 3096005,
    4355490, 5084746, 6008895, 5051956,
]


view_visibility = {
    '03-251 TOP 1 RADIAL REBAR': {
        'view_id': 2511791,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 TOP 2 RADIAL REBAR': {
        'view_id': 2511780,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 BOTTOM 3 RADIAL REBAR': {
        'view_id': 5915879,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 BOTTOM 1 RADIAL REBAR': {
        'view_id': 453671,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 TOP CONCENTRIC REBAR': {
        'view_id': 446567,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 BOTTOM 1 CONCENTRIC REBAR': {
        'view_id': 5905497,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 3D TOP ISOMETRIC': {
        'view_id': 446535,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 3D VIEW OF SHEAR REINFORCEMENT': {
        'view_id': 2503522,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 3D BOTTOM ISOMETRIC': {
        'view_id': 446525,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 GRID 1': {
        'view_id': 453468,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 GRID 2': {
        'view_id': 453478,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 GRID 3': {
        'view_id': 9137627,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 GRID 4': {
        'view_id': 4593885,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'CONSTRUCTION BARS': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': False,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 3D STOOL SETUP': {
        'view_id': 2514065,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '03-251 3D ST103 SETUP': {
        'view_id': 5998757,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '04-251 3D VIEW OF BOTTOM BURSTING REBAR': {
        'view_id': 603329,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '04-251 3D VIEW OF SURFACE REINFORCEMENT': {
        'view_id': 603339,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '04-251 3D VIEW OF PLINTH REINFORCEMENT': {
        'view_id': 966797,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '04-251 3D VIEW OF TOP BURSTING REBAR 1': {
        'view_id': 3095802,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '04-251 3D VIEW OF TOP BURSTING REBAR 2': {
        'view_id': 3096005,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '04-251 REINFORCEMENT SECTION': {
        'view_id': 4355490,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': True,
            'SECTION HAIR PINS': True,
            'SECTION PLINTH VERTICAL': True,
            'SECTION TOP RADIAL': True,
        },
    },
    '04-251 RADIATOR BASE SECTION A-A': {
        'view_id': 5084746,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '04-251 RADIATOR BASE SECTION B-B': {
        'view_id': 6008895,
        'categories': {
            'REBAR BASE VERTICAL': False,
            'REBAR BOTTOM CONCENTRIC': False,
            'REBAR BOTTOM RADIAL': False,
            'REBAR CONCENTRIC BRACING': False,
            'REBAR GRIDS': False,
            'REBAR HAIR PINS': False,
            'REBAR PLINTH CONCENTRIC': False,
            'REBAR PLINTH FACE CONCENTRICL': False,
            'REBAR PLINTH FACE HORIZONTAL': False,
            'REBAR PLINTH VERTICAL': False,
            'REBAR RADIAL ANCHOR': False,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': False,
            'REBAR STOOLS': False,
            'REBAR SUPPORT BARS': False,
            'REBAR TOP CONCENTRIC': False,
            'REBAR TOP RADIAL': False,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
    '04-251 RADIATOR BASE PLAN': {
        'view_id': 5051956,
        'categories': {
            'REBAR BASE VERTICAL': True,
            'REBAR BOTTOM CONCENTRIC': True,
            'REBAR BOTTOM RADIAL': True,
            'REBAR CONCENTRIC BRACING': True,
            'REBAR GRIDS': True,
            'REBAR HAIR PINS': True,
            'REBAR PLINTH CONCENTRIC': True,
            'REBAR PLINTH FACE CONCENTRICL': True,
            'REBAR PLINTH FACE HORIZONTAL': True,
            'REBAR PLINTH VERTICAL': True,
            'REBAR RADIAL ANCHOR': True,
            'REBAR RADIATOR BASE': True,
            'REBAR SLAB FACE CONCENTRIC': True,
            'REBAR STOOLS': True,
            'REBAR SUPPORT BARS': True,
            'REBAR TOP CONCENTRIC': True,
            'REBAR TOP RADIAL': True,
            'SECTION BASE VERTICAL': False,
            'SECTION BOTTOM RADIAL': False,
            'SECTION HAIR PINS': False,
            'SECTION PLINTH VERTICAL': False,
            'SECTION TOP RADIAL': False,
        },
    },
}


tx = Transaction(doc, "Apply view filters")
tx.Start()
try:
    for vid in view_ids:
        view = doc.GetElement(ElementId(vid))
        apply_filters_to_view(view)
        set_solid_fill_pattern_on_view(view)
        set_view_filter_visibility(view, view_visibility)

    bottom_radial_options = [
        (  0,  50,   0),   # Option 1 - very dark forest green
        ( 80, 160,   0),   # Option 2 - yellow-olive green
        (  0, 100,  80),   # Option 3 - deep teal green
        ( 60,  80,  20),   # Option 4 - army/khaki green
        ( 20,  80,  60),   # Option 5 - dark jade
    ]

    top_radial_options = [
        (  0, 255, 100),   # Option 1 - lime with teal push
        (150, 255,   0),   # Option 2 - yellow-lime
        (  0, 200, 180),   # Option 3 - bright teal
        (100, 255,  80),   # Option 4 - soft yellow-green
        ( 80, 255, 180),   # Option 5 - mint green
    ]

    top_concentric_options = [
        (255, 140,   0),   # Option 1 - Warm amber
        (204,  78,   0),   # Option 2 - Burnt orange
        (255, 179, 102),   # Option 3 - Peach / light orange
        (255,  99,  71),   # Option 4 - Coral
    ]

    bottom_concentric_options = [
        (255, 140,   0),   # Option 1 - Warm amber
        (204,  78,   0),   # Option 2 - Burnt orange
        (255, 179, 102),   # Option 3 - Peach / light orange
        (255,  99,  71),   # Option 4 - Coral
    ]

    #region GET BARMARK DICT FROM EXCEL

    pick_file_fn = getattr(forms, 'pick_file', None)
    if callable(pick_file_fn):
        FPath = pick_file_fn(file_ext='xlsx', multi_file=False, unc_paths=False)
    elif isinstance(pick_file_fn, str):
        FPath = pick_file_fn
    else:
        FPath = None

    if not FPath:
        raise SystemExit

    excel = Excel.ApplicationClass()
    excel.Visible = False
    workbook = excel.Workbooks.Open(FPath)
    xl = workbook.Worksheets["A"]


    BarList_TR = []
    BarList_BR = []
    BarList_TC = []
    BarList_BC = []
    for i in range(1,200):
        i += 1
        if "TR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
            bar_mark = str(xl.Cells(i,1).Value2)
            BarList_TR.append(bar_mark)
        elif "BR" in str(xl.Cells(i, 1).Value2).replace(" ","") :
            bar_mark = str(xl.Cells(i,1).Value2)
            BarList_BR.append(bar_mark)
        # elif "TC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        #     bar_mark = str(xl.Cells(i,1).Value2)
        #     BarList_TC.append(bar_mark)
        # elif "BC" in str(xl.Cells(i, 1).Value2).replace(" ","") :
        #     bar_mark = str(xl.Cells(i,1).Value2)
        #     BarList_BC.append(bar_mark)
        

    FEC = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType()
    for rebar in FEC:
        if "TC" in rebar.LookupParameter("Schedule Mark").AsString():
            bar_mark = rebar.LookupParameter("Schedule Mark").AsString()
            BarList_TC.append(bar_mark)
        elif "BC" in rebar.LookupParameter("Schedule Mark").AsString():
            bar_mark = rebar.LookupParameter("Schedule Mark").AsString()
            BarList_BC.append(bar_mark)

    #close excel
    workbook.Close(False)
    excel.Quit()
    #endregion

    def _natural_sort_key(s):
        import re
        parts = re.split(r'(\d+)', s)
        return [int(p) if p.isdigit() else p.upper() for p in parts]

    def _unique_natural_sorted(items):
        return sorted(set(items), key=_natural_sort_key)

    BarList_TC = _unique_natural_sorted(BarList_TC)
    BarList_BC = _unique_natural_sorted(BarList_BC)

    TR_bar_color_dict = {bar_mark: Color(*top_radial_options[i % len(top_radial_options)]) for i, bar_mark in enumerate(BarList_TR)}
    BR_bar_color_dict = {bar_mark: Color(*bottom_radial_options[i % len(bottom_radial_options)]) for i, bar_mark in enumerate(BarList_BR)}
    TC_bar_color_dict = {bar_mark: Color(*top_concentric_options[i % len(top_concentric_options)]) for i, bar_mark in enumerate(BarList_TC)}
    BC_bar_color_dict = {bar_mark: Color(*bottom_concentric_options[i % len(bottom_concentric_options)]) for i, bar_mark in enumerate(BarList_BC)}



    for view in view_ids:
        view = doc.GetElement(ElementId(view))
        rebar_categories = List[ElementId]([ElementId(BuiltInCategory.OST_Rebar)])
        def _apply_bar_mark_filter(target_view, bar_mark, prefix, color_dict):
            filter_name = "{} - {}".format(prefix, bar_mark)
            existing_filter = next((f for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement) if f.Name == filter_name), None)
            if existing_filter is None:
                new_filter = ParameterFilterElement.Create(doc, filter_name, rebar_categories)
                
                # Rule 1: Schedule Mark equals bar_mark
                schedule_bip = (
                    getattr(BuiltInParameter, "REBAR_ELEM_SCHEDULE_MARK", None)
                    or getattr(BuiltInParameter, "REBAR_SCHEDULE_MARK", None)
                    or getattr(BuiltInParameter, "REBAR_ELEM_BAR_MARK", None)
                    or getattr(BuiltInParameter, "REBAR_BAR_MARK", None)
                )
                if schedule_bip is None:
                    raise AttributeError("No compatible rebar mark BuiltInParameter found")
                param_id = ElementId(schedule_bip)
                provider = ParameterValueProvider(param_id)
                try:
                    rule1 = FilterStringRule(provider, FilterStringEquals(), bar_mark, False)
                except TypeError:
                    rule1 = FilterStringRule(provider, FilterStringEquals(), bar_mark)
                
                # Rule 2: Mark does not contain "SECTION" (optional - only if parameter exists)
                rules = [rule1]
                mark_bip = (
                     getattr(BuiltInParameter, "ALL_MODEL_MARK", None)
                     or getattr(BuiltInParameter, "REBAR_ELEM_MARK", None)
                     or getattr(BuiltInParameter, "REBAR_MARK", None)
                )

                if mark_bip is not None:
                    mark_param_id = ElementId(mark_bip)
                    mark_provider =  ParameterValueProvider(mark_param_id)
                    try:
                        contains_rule = FilterStringRule(mark_provider, FilterStringContains(), "SECTION", False)
                    except TypeError:
                        contains_rule = FilterStringRule(mark_provider, FilterStringContains(), "SECTION")
                    rule2 = FilterInverseRule(contains_rule)
                    rules.append(rule2)
                
                # Combine rules with AND logic
                element_filter = ElementParameterFilter(rules)
                new_filter.SetElementFilter(element_filter)
                fid = new_filter.Id
            else:
                fid = existing_filter.Id
                

            if not target_view.IsFilterApplied(fid):
                target_view.AddFilter(fid)

            ogs = target_view.GetFilterOverrides(fid)
            
            # Set projection/surface fill color
            set_color = getattr(ogs, "SetProjectionFillColor", None) or getattr(ogs, "SetSurfaceForegroundPatternColor", None)
            if set_color is None:
                raise AttributeError("No compatible projection/surface color setter found")
            result = set_color(color_dict[bar_mark])
            if result is not None:
                ogs = result
            
            # Set cut fill color to same value
            set_cut_color = getattr(ogs, "SetCutFillColor", None) or getattr(ogs, "SetCutForegroundPatternColor", None)
            if set_cut_color is not None:
                result = set_cut_color(color_dict[bar_mark])
                if result is not None:
                    ogs = result
            
            # Set cut line color to same value
            set_cut_line = getattr(ogs, "SetCutLineColor", None) or getattr(ogs, "SetCutBackgroundPatternColor", None)
            if set_cut_line is not None:
                result = set_cut_line(color_dict[bar_mark])
                if result is not None:
                    ogs = result
            
            target_view.SetFilterOverrides(fid, ogs)


        for bar_mark in BarList_TR:
            _apply_bar_mark_filter(view, bar_mark, "TR", TR_bar_color_dict)

        for bar_mark in BarList_BR:
            _apply_bar_mark_filter(view, bar_mark, "BR", BR_bar_color_dict)

        for bar_mark in BarList_TC:
            _apply_bar_mark_filter(view, bar_mark, "TC", TC_bar_color_dict)

        for bar_mark in BarList_BC:
            _apply_bar_mark_filter(view, bar_mark, "BC", BC_bar_color_dict)



        set_solid_fill_pattern_on_view(view)
        set_view_filter_visibility(view, view_visibility)







    tx.Commit()
except Exception:
    tx.RollBack()
    raise