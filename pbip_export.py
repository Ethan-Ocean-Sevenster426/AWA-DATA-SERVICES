#!/usr/bin/env python3
"""
pbip_export - generate a Power BI project (.pbip) from a column/measure spec.
=============================================================================
Emits a real Power BI Desktop project (PBIP): a TMDL semantic model wired to the
SharePoint K9 file + measures, and a PBIR report with ONE PAGE PER CATEGORY
(Overview, Over time, By Location, By Booking Party, By Consignee, Aging).

Open the resulting "<name>.pbip" in Power BI Desktop. Requires the PBIR preview:
  Power BI Desktop -> File -> Options -> Preview features ->
  "Store reports using enhanced metadata format (PBIR)" -> ON.

    python pbip_export.py --out "<folder>"

This is the first concrete output of the drop-folder report agent; it will be
generalised to profile any dataset once the format is confirmed loading.
"""
import os, sys, json, argparse, random, shutil

NAME = "K9 Inventory"
ENTITY = "K9 Inventory"

# Power Query that pulls the maintained K9 file from SharePoint (proven query).
M_SOURCE = '''let
    K9 = SharePoint.Contents("https://magnumopusconsultantspty352.sharepoint.com/sites/DataPrime", [ApiVersion = 15]){[Name="Shared Documents"]}[Content]{[Name="Clients"]}[Content]{[Name="ISCM"]}[Content]{[Name="K9"]}[Content],
    Source = Excel.Workbook(K9{[Name="K9 Line Item Inventory.xlsx"]}[Content], null, true){[Item="Inventory",Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Inventory Date", type date}, {"Receive Gate out Time", type datetime}, {"Volume M3", type number}, {"Weight KG", type number}})
in
    Typed'''

# The user-edited K9 Alerts file on the same SharePoint folder (title row skipped,
# real header promoted, blank spacer columns dropped).
ALERTS_M = '''let
    Site = SharePoint.Contents("https://magnumopusconsultantspty352.sharepoint.com/sites/DataPrime", [ApiVersion = 15]),
    File = Site{[Name="Shared Documents"]}[Content]{[Name="Clients"]}[Content]{[Name="ISCM"]}[Content]{[Name="K9"]}[Content]{[Name="K9 Alerts.xlsx"]}[Content],
    Workbook = Excel.Workbook(File, null, true),
    Sheet = Workbook{[Item="K9 alerts (do not delete)",Kind="Sheet"]}[Data],
    Skipped = Table.Skip(Sheet, 1),
    Promoted = Table.PromoteHeaders(Skipped, [PromoteAllScalars=true]),
    Selected = Table.SelectColumns(Promoted, {"Alert Date", "Shipment Number", "Release Date", "Notes"}, MissingField.UseNull),
    Cleaned = Table.SelectRows(Selected, each [Shipment Number] <> null or [Notes] <> null),
    Typed = Table.TransformColumnTypes(Cleaned, {{"Alert Date", type date}, {"Release Date", type date}, {"Shipment Number", type text}, {"Notes", type text}})
in
    Typed'''

def dummy_alerts_m():
    rows = ['{#date(2025,2,24), "S00476127", #date(2025,2,24), "Pallet inspected, K9 screening done, released."}',
            '{#date(2025,4,14), "S00490789", #date(2025,4,21), "Vehicle transferred for inspection, released."}']
    coltypes = '[#"Alert Date"=date, #"Shipment Number"=text, #"Release Date"=date, Notes=text]'
    return ("let\n    Source = #table(type table " + coltypes + ",\n        {\n            "
            + ",\n            ".join(rows) + "\n        })\nin\n    Source")

def dummy_m(rows=40):
    """Inline sample data so the model loads with no SharePoint sign-in."""
    random.seed(7)
    locs = ["W1", "W2", "W3", "E2", "SW-12-A", "SW-29-A"]
    bps = ["AMERICAN WORLDWIDE AGENCIES", "ISLAND CARGO SUPPORT", "CONDOR CARGO CHICAGO", "ROCK-IT GLOBAL"]
    names = ["MOORE FANS LLC", "TRONAIR INC", "STARKIST SAMOA", "ARTEMIDE", "EHPLABS LLC", "RUPERT GIBBON"]
    snaps = [(2026, 6, d) for d in (20, 22, 24, 26, 28)]
    recs = []
    for i in range(rows):
        y, mo, d = random.choice(snaps)
        gy, gmo, gd = 2026, random.choice([1, 2, 3, 4, 5, 6]), random.randint(1, 28)
        rcn = f"S00{random.randint(500000, 580000)}"
        vol = round(random.uniform(0.01, 3.5), 3); wt = round(random.uniform(1, 1200), 1)
        recs.append("{#date(%d,%d,%d), \"ISBINT%s-%05d\", \"%s\", \"%s\", #datetime(%d,%d,%d,8,0,0), "
                    "\"PUT - Putaway\", \"%s\", \"%s\", \"%s\", \"%s M3\", \"%s KG\", %s, %s}" % (
                        y, mo, d, rcn, i, rcn, random.choice(locs), gy, gmo, gd,
                        random.choice(names), random.choice(names), random.choice(bps), vol, wt, vol, wt))
    coltypes = ('[#"Inventory Date"=date, #"Package ID"=text, #"RCN Reference"=text, Location=text, '
                '#"Receive Gate out Time"=datetime, Status=text, Consignee=text, Consignor=text, '
                '#"Booking Party"=text, Volume=text, Weight=text, #"Volume M3"=number, #"Weight KG"=number]')
    body = ",\n            ".join(recs)
    return f"let\n    Source = #table(type table {coltypes},\n        {{\n            {body}\n        }})\nin\n    Source"

STR_COLS = ["Package ID", "RCN Reference", "Location", "Status",
            "Consignee", "Consignor", "Booking Party", "Volume", "Weight"]
NUM_COLS = ["Volume M3", "Weight KG"]
DATE_COLS = ["Inventory Date", "Receive Gate out Time"]

MEASURES = [
    ("Shipment Count", "DISTINCTCOUNT('K9 Inventory'[RCN Reference])", "#,0"),
    ("Total Weight (KG)", "SUM('K9 Inventory'[Weight KG])", "#,0"),
    ("Total CBM (M3)", "SUM('K9 Inventory'[Volume M3])", "#,0.000"),
    ("Package Count", "COUNTROWS('K9 Inventory')", "#,0"),
    ("Avg Days in Warehouse", "AVERAGE('K9 Inventory'[Days in Warehouse])", "#,0"),
]

# Navy chrome + red data theme (filled title bars, red table headers)
NAVY, RED, LBLUE = "#1F3864", "#C00000", "#2E78D2"
THEME = {
    "name": "K9Theme",
    "dataColors": [RED, NAVY, LBLUE, "#E0586B", "#5B9BD5", "#7B1E2B", "#9DC3E6", "#3A4A6B"],
    "background": "#FFFFFF",
    "foreground": NAVY,
    "tableAccent": RED,
    "good": NAVY, "neutral": "#5B9BD5", "bad": RED,
    "maximum": NAVY, "minimum": "#9DC3E6",
    "textClasses": {
        "title": {"color": "#FFFFFF", "fontFace": "Segoe UI Semibold"},
        "header": {"color": "#FFFFFF", "fontFace": "Segoe UI Semibold"},
        "callout": {"color": NAVY, "fontFace": "Segoe UI Semibold"},
        "label": {"color": "#2C3E50"},
    },
    "visualStyles": {
        "*": {"*": {
            "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
            "border": [{"show": True, "color": {"solid": {"color": NAVY}}, "radius": 8}],
            "title": [{"show": True, "fontColor": {"solid": {"color": "#FFFFFF"}},
                       "background": {"solid": {"color": NAVY}}, "alignment": "center",
                       "fontSize": 13, "bold": True, "titleWrap": True}],
            "visualHeader": [{"show": False}],
            # readable dark fonts on the transparent (light) background
            "categoryAxis": [{"labelColor": {"solid": {"color": NAVY}}, "titleColor": {"solid": {"color": NAVY}},
                              "gridlineColor": {"solid": {"color": "#D6E0F0"}}}],
            "valueAxis": [{"labelColor": {"solid": {"color": NAVY}}, "titleColor": {"solid": {"color": NAVY}},
                           "gridlineColor": {"solid": {"color": "#D6E0F0"}}}],
            "legend": [{"labelColor": {"solid": {"color": NAVY}}, "titleColor": {"solid": {"color": NAVY}}}],
            "labels": [{"color": {"solid": {"color": NAVY}}}],
        }},
        "card": {"*": {
            "labels": [{"color": {"solid": {"color": RED}}, "fontSize": 34, "bold": True}],
            "categoryLabels": [{"color": {"solid": {"color": NAVY}}, "fontSize": 12, "bold": True}],
            "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}}],
            "border": [{"show": True, "color": {"solid": {"color": NAVY}}, "radius": 8}],
        }},
        "tableEx": {"*": {
            "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}}],
            "border": [{"show": True, "color": {"solid": {"color": NAVY}}, "radius": 8}],
            "columnHeaders": [{"fontColor": {"solid": {"color": "#FFFFFF"}},
                               "backColor": {"solid": {"color": RED}}, "bold": True}],
            "values": [{"fontColorPrimary": {"solid": {"color": "#2C3E50"}},
                        "backColorPrimary": {"solid": {"color": "#FFFFFF"}},
                        "backColorSecondary": {"solid": {"color": "#EEF3FB"}}}],
            "grid": [{"gridVerticalColor": {"solid": {"color": "#D6E0F0"}},
                      "gridHorizontalColor": {"solid": {"color": "#D6E0F0"}}}],
            "total": [{"fontColor": {"solid": {"color": "#FFFFFF"}},
                       "backColor": {"solid": {"color": RED}}, "bold": True}],
        }},
        "pivotTable": {"*": {
            "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}}],
            "border": [{"show": True, "color": {"solid": {"color": NAVY}}, "radius": 8}],
            "columnHeaders": [{"fontColor": {"solid": {"color": "#FFFFFF"}},
                               "backColor": {"solid": {"color": RED}}, "bold": True}],
            "values": [{"backColorSecondary": {"solid": {"color": "#EEF3FB"}}}],
        }},
        "slicer": {"*": {
            "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}}],
            "border": [{"show": True, "color": {"solid": {"color": NAVY}}, "radius": 8}],
            "header": [{"fontColor": {"solid": {"color": NAVY}}, "bold": True}],
            "items": [{"fontColor": {"solid": {"color": NAVY}}}],
        }},
        "lineChart": {"*": {
            "lineStyles": [{"strokeWidth": 3, "interpolation": "smooth", "smoothType": "monotone"}],
        }},
        "lineClusteredColumnComboChart": {"*": {
            "lineStyles": [{"strokeWidth": 3, "interpolation": "smooth", "smoothType": "monotone"}],
        }},
        "page": {"*": {
            "background": [{"color": {"solid": {"color": "#EEF3FB"}}, "transparency": 0}],
            "outspace": [{"color": {"solid": {"color": "#EEF3FB"}}, "transparency": 0}],
        }},
    },
}

def lt(n):  # deterministic lineage tags
    return f"a0000000-0000-0000-0000-{n:012d}"

# --------------------------------------------------------------------------- #
def tmdl_table(m_source):
    L = []
    L.append(f"table '{ENTITY}'")
    L.append(f"\tlineageTag: {lt(1)}")
    L.append("")
    i = 100
    for name, dax, fmt in MEASURES:
        L.append(f"\tmeasure '{name}' = {dax}")
        if fmt:
            L.append(f"\t\tformatString: {fmt}")
        L.append(f"\t\tlineageTag: {lt(i)}"); i += 1
        L.append("")
    # calculated columns for aging
    L.append("\tcolumn 'Days in Warehouse' = DATEDIFF('K9 Inventory'[Receive Gate out Time], "
             "MAXX(ALL('K9 Inventory'), 'K9 Inventory'[Inventory Date]), DAY)")
    L.append("\t\tdataType: int64")
    L.append(f"\t\tlineageTag: {lt(i)}"); i += 1
    L.append("\t\tsummarizeBy: none")
    L.append("")
    L.append("\tcolumn 'Age Bucket' = SWITCH(TRUE(), 'K9 Inventory'[Days in Warehouse] <= 7, \"0-7 d\", "
             "'K9 Inventory'[Days in Warehouse] <= 30, \"8-30 d\", 'K9 Inventory'[Days in Warehouse] <= 90, \"31-90 d\", "
             "'K9 Inventory'[Days in Warehouse] <= 180, \"91-180 d\", 'K9 Inventory'[Days in Warehouse] <= 365, \"181-365 d\", \"365+ d\")")
    L.append("\t\tdataType: string")
    L.append(f"\t\tlineageTag: {lt(i)}"); i += 1
    L.append("\t\tsummarizeBy: none")
    L.append("")
    # source columns
    for c in DATE_COLS:
        L.append(f"\tcolumn '{c}'")
        L.append("\t\tdataType: dateTime")
        L.append("\t\tformatString: dd mmm yyyy" + (" hh:nn" if c == "Receive Gate out Time" else ""))
        L.append(f"\t\tlineageTag: {lt(i)}"); i += 1
        L.append("\t\tsummarizeBy: none")
        L.append(f"\t\tsourceColumn: {c}")
        L.append("")
        L.append("\t\tannotation SummarizationSetBy = Automatic")
        L.append("")
    for c in STR_COLS:
        L.append(f"\tcolumn '{c}'")
        L.append("\t\tdataType: string")
        L.append(f"\t\tlineageTag: {lt(i)}"); i += 1
        L.append("\t\tsummarizeBy: none")
        L.append(f"\t\tsourceColumn: {c}")
        L.append("")
        L.append("\t\tannotation SummarizationSetBy = Automatic")
        L.append("")
    for c in NUM_COLS:
        L.append(f"\tcolumn '{c}'")
        L.append("\t\tdataType: double")
        L.append("\t\tformatString: #,0.000")
        L.append(f"\t\tlineageTag: {lt(i)}"); i += 1
        L.append("\t\tsummarizeBy: sum")
        L.append(f"\t\tsourceColumn: {c}")
        L.append("")
        L.append("\t\tannotation SummarizationSetBy = Automatic")
        L.append("")
    # partition (M source)
    L.append(f"\tpartition '{ENTITY}' = m")
    L.append("\t\tmode: import")
    L.append("\t\tsource =")
    for line in m_source.splitlines():
        L.append("\t\t\t\t" + line)
    L.append("")
    L.append("\tannotation PBI_ResultType = Table")
    L.append("")
    return "\n".join(L)

def tmdl_alerts_table(m_source):
    L = ["table 'K9 Alerts'", f"\tlineageTag: {lt(900)}", ""]
    cols = [("Alert Date", "dateTime", "dd mmm yyyy"), ("Shipment Number", "string", None),
            ("Release Date", "dateTime", "dd mmm yyyy"), ("Notes", "string", None)]
    i = 910
    for name, dt, fmt in cols:
        L.append(f"\tcolumn '{name}'")
        L.append(f"\t\tdataType: {dt}")
        if fmt:
            L.append(f"\t\tformatString: {fmt}")
        L.append(f"\t\tlineageTag: {lt(i)}"); i += 1
        L.append("\t\tsummarizeBy: none")
        L.append(f"\t\tsourceColumn: {name}")
        L.append("")
        L.append("\t\tannotation SummarizationSetBy = Automatic")
        L.append("")
    L.append("\tpartition 'K9 Alerts' = m")
    L.append("\t\tmode: import")
    L.append("\t\tsource =")
    for line in m_source.splitlines():
        L.append("\t\t\t\t" + line)
    L.append("")
    L.append("\tannotation PBI_ResultType = Table")
    L.append("")
    return "\n".join(L)

def tmdl_calendar_table():
    L = ["table Calendar", f"\tlineageTag: {lt(700)}", "",
         "\tcolumn Date",
         "\t\tdataType: dateTime",
         "\t\tisNameInferred",
         "\t\tformatString: dd mmm yyyy",
         f"\t\tlineageTag: {lt(701)}",
         "\t\tsummarizeBy: none",
         "\t\tsourceColumn: [Date]",
         "",
         "\t\tannotation SummarizationSetBy = Automatic",
         ""]
    calc = [("Year", "YEAR(Calendar[Date])", "int64", "0", None, "Automatic"),
            ("MonthNo", "MONTH(Calendar[Date])", "int64", "0", None, "Automatic"),
            ("Month", 'FORMAT(Calendar[Date], "mmm")', "string", None, "MonthNo", "None"),
            ("Month Year", 'FORMAT(Calendar[Date], "mmm yyyy")', "string", None, None, "None"),
            ("Quarter", '"Q" & FORMAT(ROUNDUP(MONTH(Calendar[Date]) / 3, 0), "0")', "string", None, None, "None")]
    i = 702
    for name, dax, dt, fmt, sortby, sumby in calc:
        nm = f"'{name}'" if " " in name else name
        L.append(f"\tcolumn {nm} = {dax}")
        L.append(f"\t\tdataType: {dt}")
        if fmt:
            L.append(f"\t\tformatString: {fmt}")
        L.append(f"\t\tlineageTag: {lt(i)}"); i += 1
        L.append("\t\tsummarizeBy: none")
        if sortby:
            L.append(f"\t\tsortByColumn: {sortby}")
        L.append("")
        L.append(f"\t\tannotation SummarizationSetBy = {sumby}")
        L.append("")
    L.append("\tpartition Calendar = calculated")
    L.append("\t\tmode: import")
    L.append("\t\tsource = CALENDAR(MIN('K9 Inventory'[Inventory Date]), MAX('K9 Inventory'[Inventory Date]))")
    L.append("")
    return "\n".join(L)

def tmdl_relationships(has_alerts):
    L = [f"relationship {lt(801)}",
         "\tfromColumn: 'K9 Inventory'.'Inventory Date'",
         "\ttoColumn: Calendar.Date", ""]
    if has_alerts:
        L += [f"relationship {lt(802)}",
              "\tfromColumn: 'K9 Alerts'.'Alert Date'",
              "\ttoColumn: Calendar.Date", ""]
    return "\n".join(L)

def tmdl_model():
    return ("model Model\n"
            "\tculture: en-US\n"
            "\tdefaultPowerBIDataSourceVersion: powerBI_V3\n"
            "\tsourceQueryCulture: en-US\n"
            "\n"
            "annotation PBI_QueryOrder = [\"K9 Inventory\", \"K9 Alerts\", \"Calendar\"]\n"
            "\n"
            "annotation __PBI_TimeIntelligenceEnabled = 0\n")

# --------------------------------------------------------------------------- #
# PBIR report helpers
# --------------------------------------------------------------------------- #
def col_field(prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": ENTITY}}, "Property": prop}}

def measure_field(prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": ENTITY}}, "Property": prop}}

def projection(field, prop):
    return {"field": field, "queryRef": f"{ENTITY}.{prop}", "nativeQueryRef": prop}

def visual_json(name, vtype, x, y, w, h, wells):
    qs = {}
    for role, items in wells.items():
        qs[role] = {"projections": [projection(f, p) for (f, p) in items]}
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": 0, "width": w, "height": h, "tabOrder": 0},
        "visual": {
            "visualType": vtype,
            "query": {"queryState": qs},
            "drillFilterOtherVisuals": True,
            "visualContainerObjects": {
                "visualHeader": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
            },
        },
    }

def card(name, x, y, w, h, measure):
    return visual_json(name, "card", x, y, w, h, {"Values": [(measure_field(measure), measure)]})

def titled(v, text):
    """Give a visual a custom title-bar caption (header)."""
    v["visual"].setdefault("objects", {})["title"] = [
        {"properties": {
            "show": {"expr": {"Literal": {"Value": "true"}}},
            "text": {"expr": {"Literal": {"Value": "'" + text.replace("'", " ") + "'"}}}}}]
    return v

def slicer(name, x, y, w, h, prop):
    """A clean dropdown slicer (no navy title bar), rendered in front of other visuals."""
    v = visual_json(name, "slicer", x, y, w, h, {"Values": [(col_field(prop), prop)]})
    v["position"]["z"] = 900                       # bring to front so the dropdown list overlays charts
    v["visual"]["objects"] = {
        "data": [{"properties": {"mode": {"expr": {"Literal": {"Value": "'Dropdown'"}}}}}],
        "title": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
    }
    return v

def entity_table(name, x, y, w, h, entity, cols):
    """A tableEx bound to columns of a specific entity (used for the K9 Alerts table)."""
    projs = [{"field": {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": c}},
              "queryRef": f"{entity}.{c}", "nativeQueryRef": c} for c in cols]
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": 0, "width": w, "height": h, "tabOrder": 0},
        "visual": {"visualType": "tableEx",
                   "query": {"queryState": {"Values": {"projections": projs}}},
                   "drillFilterOtherVisuals": True,
                   "visualContainerObjects": {
                       "visualHeader": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]}},
    }

def cal_slicer(name, x, y, w, h):
    """Date hierarchy slicer (Year > Quarter > Month > Day) on the Calendar table -
    filters every related table through the relationships."""
    fields = ["Year", "Quarter", "Month", "Date"]
    projs = [{"field": {"Column": {"Expression": {"SourceRef": {"Entity": "Calendar"}}, "Property": f}},
              "queryRef": f"Calendar.{f}", "nativeQueryRef": f} for f in fields]
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": 900, "width": w, "height": h, "tabOrder": 0},
        "visual": {
            "visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": projs}}},
            "objects": {
                "data": [{"properties": {"mode": {"expr": {"Literal": {"Value": "'Dropdown'"}}}}}],
                "title": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
            },
            "drillFilterOtherVisuals": True,
            "visualContainerObjects": {
                "visualHeader": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]},
        },
    }

def image_visual(name, x, y, w, h, item_name="logo.png"):
    """An image visual bound to a registered resource image."""
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": 0, "width": w, "height": h, "tabOrder": 0},
        "visual": {
            "visualType": "image",
            "objects": {"general": [{"properties": {
                "imageScalingType": {"expr": {"Literal": {"Value": "'Fit'"}}},
                "imageUrl": {"expr": {"ResourcePackageItem": {
                    "PackageName": "RegisteredResources", "PackageType": 1, "ItemName": item_name}}},
            }}]},
            "visualContainerObjects": {
                "visualHeader": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]},
        },
    }

def write(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content if isinstance(content, str) else json.dumps(content, indent=2))

def build(out_dir, m_source, with_visuals=True, logo_path=None, alerts_source=None):
    sm = os.path.join(out_dir, f"{NAME}.SemanticModel")
    rp = os.path.join(out_dir, f"{NAME}.Report")
    has_logo = bool(logo_path and os.path.exists(logo_path))

    # ---- .pbip ----
    write(os.path.join(out_dir, f"{NAME}.pbip"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    })

    # ---- Semantic model ----
    write(os.path.join(sm, ".platform"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": "00000000-0000-0000-0000-0000000000a1"},
    })
    write(os.path.join(sm, "definition.pbism"), {"version": "4.2", "settings": {}})
    write(os.path.join(sm, "definition", "model.tmdl"), tmdl_model())
    write(os.path.join(sm, "definition", "tables", f"{ENTITY}.tmdl"), tmdl_table(m_source))
    if alerts_source:
        write(os.path.join(sm, "definition", "tables", "K9 Alerts.tmdl"), tmdl_alerts_table(alerts_source))
    # shared Date dimension + relationships to the fact tables
    write(os.path.join(sm, "definition", "tables", "Calendar.tmdl"), tmdl_calendar_table())
    write(os.path.join(sm, "definition", "relationships.tmdl"), tmdl_relationships(bool(alerts_source)))

    # ---- Report ----
    write(os.path.join(rp, ".platform"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": "00000000-0000-0000-0000-0000000000a2"},
    })
    write(os.path.join(rp, "definition.pbir"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}},
    })
    write(os.path.join(rp, "definition", "report.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.0.0/schema.json",
        "layoutOptimization": "None",
        "themeCollection": {
            "baseTheme": {"name": "CY24SU06", "reportVersionAtImport": "5.55", "type": "SharedResources"},
            "customTheme": {"name": "K9Theme.json", "reportVersionAtImport": "5.55", "type": "RegisteredResources"},
        },
        "resourcePackages": [{
            "name": "RegisteredResources", "type": "RegisteredResources",
            "items": ([{"name": "K9Theme.json", "path": "K9Theme.json", "type": "CustomTheme"}]
                      + ([{"name": "logo.png", "path": "logo.png", "type": "Image"}] if has_logo else [])),
        }],
    })
    # Registered resources live under StaticResources/RegisteredResources (Power BI convention)
    write(os.path.join(rp, "StaticResources", "RegisteredResources", "K9Theme.json"), THEME)
    if has_logo:
        os.makedirs(os.path.join(rp, "StaticResources", "RegisteredResources"), exist_ok=True)
        shutil.copy(logo_path, os.path.join(rp, "StaticResources", "RegisteredResources", "logo.png"))
    # PBIR requires a version marker in the report definition folder
    write(os.path.join(rp, "definition", "version.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
        "version": "2.0.0",
    })

    inv_date = (col_field("Inventory Date"), "Inventory Date")
    pages = []

    def add_page(pid, title, visuals):
        pages.append((pid, title))
        write(os.path.join(rp, "definition", "pages", pid, "page.json"), {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json",
            "name": pid, "displayName": title,
            "displayOption": "FitToPage", "height": 720, "width": 1280,
        })
        if with_visuals:
            for v in visuals:
                write(os.path.join(rp, "definition", "pages", pid, "visuals", v["name"], "visual.json"), v)

    M = lambda p: (measure_field(p), p)      # measure projection
    C = lambda p: (col_field(p), p)          # column projection

    # Shared header (logo + 4 dropdown filters), repeated on every page. y 8..92
    def header(pid):
        hv, sx0 = [], 8
        if has_logo:
            hv.append(image_visual(f"logo_{pid}", 8, 8, 150, 84))
            sx0 = 168
        sw = (1272 - sx0 - 3 * 8) // 4
        hv.append(cal_slicer(f"sl_date_{pid}", sx0, 8, sw, 84))     # Year>Quarter>Month>Day
        for i, (k, prop) in enumerate([("loc", "Location"), ("bp", "Booking Party"),
                                       ("st", "Status")], start=1):
            hv.append(slicer(f"sl_{k}_{pid}", sx0 + i * (sw + 8), 8, sw, 84, prop))
        return hv

    # page bodies sit below the header band (y 104..712)
    add_page("overview", "Overview", header("ov") + [
        titled(visual_json("combo_main", "lineClusteredColumnComboChart", 8, 104, 628, 284, {
            "Category": [C("Inventory Date")], "Y": [M("Total Weight (KG)")], "Y2": [M("Shipment Count")]}),
            "Inventory Weight and Shipments Over Time"),
        titled(visual_json("tbl_main", "pivotTable", 644, 104, 628, 608, {
            "Rows": [C("Consignee")],
            "Values": [M("Package Count"), M("Total Weight (KG)"), M("Total CBM (M3)")]}),
            "Detailed Consignee Breakdown"),
        titled(visual_json("donut_main", "donutChart", 8, 396, 310, 316, {
            "Category": [C("Booking Party")], "Y": [M("Total CBM (M3)")]}),
            "CBM Share by Booking Party"),
        titled(visual_json("bar_main", "clusteredBarChart", 326, 396, 310, 316, {
            "Category": [C("Consignee")], "Y": [M("Total Weight (KG)")]}),
            "Top Consignees by Weight"),
    ])
    add_page("overtime", "Over time", header("ot") + [
        titled(visual_json("line_ot", "lineChart", 8, 104, 1264, 372, {
            "Category": [C("Inventory Date")],
            "Y": [M("Shipment Count"), M("Total Weight (KG)"), M("Total CBM (M3)")]}),
            "Weight, CBM and Shipments Over Time"),
        titled(visual_json("col_ot", "clusteredColumnChart", 8, 484, 1264, 228, {
            "Category": [C("Inventory Date")], "Y": [M("Package Count")]}),
            "Packages Over Time"),
    ])
    add_page("bylocation", "By Location", header("loc") + [
        titled(visual_json("col_loc", "clusteredColumnChart", 8, 104, 628, 608, {
            "Category": [C("Location")], "Y": [M("Package Count"), M("Total CBM (M3)")]}),
            "Packages and CBM by Location"),
        titled(visual_json("tbl_loc", "tableEx", 644, 104, 628, 608, {
            "Values": [C("Location"), M("Package Count"), M("Total Weight (KG)"), M("Total CBM (M3)")]}),
            "Location Detail"),
    ])
    add_page("bybooking", "By Booking Party", header("bp") + [
        titled(visual_json("bar_bp", "clusteredBarChart", 8, 104, 628, 608, {
            "Category": [C("Booking Party")], "Y": [M("Package Count")]}),
            "Packages by Booking Party"),
        titled(visual_json("donut_bp", "donutChart", 644, 104, 628, 608, {
            "Category": [C("Booking Party")], "Y": [M("Total CBM (M3)")]}),
            "CBM Share by Booking Party"),
    ])
    add_page("byconsignee", "By Consignee", header("cn") + [
        titled(visual_json("bar_cn", "clusteredBarChart", 8, 104, 628, 608, {
            "Category": [C("Consignee")], "Y": [M("Package Count")]}),
            "Packages by Consignee"),
        titled(visual_json("tbl_cn", "tableEx", 644, 104, 628, 608, {
            "Values": [C("Consignee"), M("Package Count"), M("Total Weight (KG)")]}),
            "Consignee Detail"),
    ])
    add_page("aging", "Aging", header("ag") + [
        card("card_avgage", 8, 104, 628, 100, "Avg Days in Warehouse"),
        titled(visual_json("col_age", "clusteredColumnChart", 8, 212, 628, 500, {
            "Category": [C("Age Bucket")], "Y": [M("Package Count")]}),
            "Packages by Age Bucket"),
        titled(visual_json("tbl_age", "tableEx", 644, 104, 628, 608, {
            "Values": [C("Age Bucket"), M("Package Count"), M("Total CBM (M3)"), M("Total Weight (KG)")]}),
            "Aging Detail"),
    ])
    if alerts_source:
        add_page("alerts", "K9 Alerts",
                 ([image_visual("logo_al", 8, 8, 150, 84)] if has_logo else []) + [
                     titled(entity_table("tbl_alerts", 8, 104, 1264, 608, "K9 Alerts",
                            ["Alert Date", "Shipment Number", "Release Date", "Notes"]),
                            "K9 Inspection Alerts  -  edit on SharePoint"),
                 ])

    write(os.path.join(rp, "definition", "pages", "pages.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [p for p, _ in pages],
        "activePageName": pages[0][0],
    })
    return [t for _, t in pages]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="folder to write the .pbip project into")
    ap.add_argument("--dummy", action="store_true", help="embed inline sample data (no SharePoint sign-in)")
    ap.add_argument("--no-visuals", action="store_true", help="blank pages only (reliable load; add visuals in the UI)")
    ap.add_argument("--logo", help="path to a logo image (png) for the header")
    args = ap.parse_args()
    m_source = dummy_m() if args.dummy else M_SOURCE
    alerts_source = dummy_alerts_m() if args.dummy else ALERTS_M
    titles = build(args.out, m_source, with_visuals=not args.no_visuals, logo_path=args.logo,
                   alerts_source=alerts_source)
    print(f"Wrote Power BI project '{NAME}.pbip' into: {args.out}")
    print(f"  data: {'inline dummy' if args.dummy else 'SharePoint K9 file'}")
    print(f"  pages: {', '.join(titles)}  (visuals: {'no' if args.no_visuals else 'yes'})")

if __name__ == "__main__":
    main()
