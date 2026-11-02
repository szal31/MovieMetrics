"""Generates the MovieMetrics Power BI project (.pbip) next to this script.
Run:  python build_pbip.py   then open MovieMetrics.pbip in Power BI Desktop and press Refresh."""
import json, os, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(HERE, "movies_data.csv")
T = "movies_data"
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"


def write(rel, obj):
    p = os.path.join(HERE, rel)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2) if not isinstance(obj, str) else f.write(obj)


# ---------------- semantic model ----------------
TEXT = ["Movie", "Director", "Actor 1", "Actor 2", "Actor 3", "Genre"]
TYPES = {  # column -> (M type, TOM type)
    "Movie": ("type text", "string"), "Director": ("type text", "string"),
    "Running time": ("Int64.Type", "int64"), "Actor 1": ("type text", "string"),
    "Actor 2": ("type text", "string"), "Actor 3": ("type text", "string"),
    "Genre": ("type text", "string"), "Budget": ("Currency.Type", "decimal"),
    "Box Office": ("Currency.Type", "decimal"), "Actors Box Office %": ("type number", "double"),
    "Director Box Office %": ("type number", "double"), "Earnings": ("Currency.Type", "decimal"),
    "Oscar and Golden Globes nominations": ("Int64.Type", "int64"),
    "Oscar and Golden Globes awards": ("Int64.Type", "int64"),
    "Release year": ("Int64.Type", "int64"), "IMDb score": ("type number", "double"),
}
m = f'''let
    Source = Csv.Document(File.Contents("{CSV}"), [Delimiter=",", Columns=16, Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Cleaned Text" = Table.TransformColumns(#"Promoted Headers", {{{", ".join('{"%s", each Text.Trim(Text.Clean(_)), type text}' % c for c in TEXT)}}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Cleaned Text", {{{", ".join('{"%s", %s}' % (c, t[0]) for c, t in TYPES.items())}}}, "en-US"),
    #"Actor 3 Unknown" = Table.TransformColumns(#"Changed Type", {{"Actor 3", each if _ = null or _ = "" then "Unknown" else _, type text}}),
    #"Awards Zero" = Table.ReplaceValue(#"Actor 3 Unknown", null, 0, Replacer.ReplaceValue, {{"Oscar and Golden Globes awards"}}),
    #"Removed Exact Duplicates" = Table.Distinct(#"Awards Zero"),
    #"Sorted Box Office" = Table.Buffer(Table.Sort(#"Removed Exact Duplicates", {{{{"Box Office", Order.Descending}}, {{"Oscar and Golden Globes awards", Order.Descending}}, {{"Oscar and Golden Globes nominations", Order.Descending}}}})),
    #"Removed Conflicting Duplicates" = Table.Distinct(#"Sorted Box Office", {{"Movie", "Release year", "Director"}}),
    #"Fixed Invalid %" = Table.TransformColumns(#"Removed Conflicting Duplicates", {{"Actors Box Office %", each if _ > 100 then null else _, type number}})
in
    #"Fixed Invalid %"'''

cols = [{"name": c, "dataType": t[1], "sourceColumn": c,
         **({"formatString": "\\$#,0"} if t[1] == "decimal" else {}),
         **({"summarizeBy": "none"} if c == "Release year" else {})} for c, t in TYPES.items()]
calc = [
    ("ROI", "double", "DIVIDE(movies_data[Earnings], movies_data[Budget]) * 100", "#,0.0\"%\""),
    ("Profit Status", "string", 'IF(movies_data[Earnings] > 0, "Profit", IF(movies_data[Earnings] < 0, "Loss", "Break-even"))', None),
    ("Budget Category", "string", 'SWITCH(TRUE(), movies_data[Budget] < 20000000, "Low", movies_data[Budget] <= 80000000, "Medium", "High")', None),
    ("Budget Order", "int64", "SWITCH(TRUE(), movies_data[Budget] < 20000000, 1, movies_data[Budget] <= 80000000, 2, 3)", None),
    ("Decade", "string", 'FLOOR(movies_data[Release year], 10) & "s"', None),
    ("IMDb Category", "string", 'SWITCH(TRUE(), movies_data[IMDb score] >= 7, "High", movies_data[IMDb score] >= 5, "Medium", "Low")', None),
    ("IMDb Order", "int64", "SWITCH(TRUE(), movies_data[IMDb score] >= 7, 1, movies_data[IMDb score] >= 5, 2, 3)", None),
    ("Award Status", "string", 'IF(movies_data[Oscar and Golden Globes awards] > 0, "Award winner", "No award")', None),
    # hidden bin starts for the Chapter 3 EDA histograms
    ("Budget Band ($M)", "double", "FLOOR(movies_data[Budget], 20000000) / 1000000", "#,0"),
    ("Box Office Band ($M)", "double", "MIN(FLOOR(movies_data[Box Office], 100000000) / 1000000, 1000)", "#,0"),
    ("IMDb Band", "double", "FLOOR(movies_data[IMDb score], 0.5)", "0.0"),
    ("Running Time Band", "double", "FLOOR(movies_data[Running time], 20)", "0"),
]
for name, dt, expr, fmt in calc:
    c = {"type": "calculated", "name": name, "dataType": dt, "expression": expr}
    if fmt: c["formatString"] = fmt
    if name.endswith("Order") or "Band" in name: c.update(isHidden=True, summarizeBy="none")
    cols.append(c)
for c in cols:
    if c["name"] == "Budget Category": c["sortByColumn"] = "Budget Order"
    if c["name"] == "IMDb Category": c["sortByColumn"] = "IMDb Order"

PCT = '#,0.0"%"'
measures = [
    ("Total Movies", "COUNTROWS(movies_data)", "#,0"),
    ("Total Budget", "SUM(movies_data[Budget])", "\\$#,0"),
    ("Total Box Office", "SUM(movies_data[Box Office])", "\\$#,0"),
    ("Total Earnings", "SUM(movies_data[Earnings])", "\\$#,0"),
    ("Median ROI", "MEDIAN(movies_data[ROI])", PCT),
    ("Loss %", 'DIVIDE(CALCULATE(COUNTROWS(movies_data), movies_data[Profit Status] = "Loss"), COUNTROWS(movies_data))', "0.0%"),
    ("Median Budget", "MEDIAN(movies_data[Budget])", "\\$#,0"),
    ("Median Earnings", "MEDIAN(movies_data[Earnings])", "\\$#,0"),
    # genres under 25 films are hidden from the ROI ranking (see 3.6 balance check)
    ("Median ROI (25+ films)", "IF(COUNTROWS(movies_data) >= 25, [Median ROI])", PCT),
    ("Top 10 Director Box Office", "IF(ISINSCOPE(movies_data[Director]) && RANKX(ALLSELECTED(movies_data[Director]), [Total Box Office]) <= 10, [Total Box Office])", "\\$#,0"),
    ("Top 10 Actor Box Office", "IF(ISINSCOPE(movies_data[Actor 1]) && RANKX(ALLSELECTED(movies_data[Actor 1]), [Total Box Office]) <= 10, [Total Box Office])", "\\$#,0"),
]
model = {"compatibilityLevel": 1567, "model": {
    "culture": "en-US", "defaultPowerBIDataSourceVersion": "powerBI_V3",
    "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
    "tables": [{"name": T, "columns": cols,
                "measures": [{"name": n, "expression": e, "formatString": f} for n, e, f in measures],
                "partitions": [{"name": T, "mode": "import", "source": {"type": "m", "expression": m.split("\n")}}]}],
    "annotations": [{"name": "PBI_QueryOrder", "value": json.dumps([T])}]}}

# ---------------- report ----------------
def lit(v): return {"expr": {"Literal": {"Value": v}}}
def col(name): return {"Column": {"Expression": {"SourceRef": {"Entity": T}}, "Property": name}}
def mea(name): return {"Measure": {"Expression": {"SourceRef": {"Entity": T}}, "Property": name}}
def proj(f):
    prop = next(iter(f.values()))["Property"]
    return {"field": f, "queryRef": f"{T}.{prop}", "nativeQueryRef": prop}

BLUE, ORANGE, GREY = "#1F6FB2", "#E07B22", "#5B6770"


def visual(vtype, pos, title=None, roles=None, sort=None, objects=None):
    v = {"visualType": vtype}
    if roles:
        v["query"] = {"queryState": {r: {"projections": [proj(f) for f in fs]} for r, fs in roles.items()}}
        if sort:
            v["query"]["sortDefinition"] = {"sort": [{"field": f, "direction": d} for f, d in sort], "isDefaultSort": True}
    if objects: v["objects"] = objects
    vc = {"title": [{"properties": {"show": lit("true"), "text": lit(f"'{title}'"), "fontSize": lit("12D"), "bold": lit("true"),
                                    "fontColor": {"solid": {"color": lit("'#1B2A41'")}}}}]} if title else \
         {"title": [{"properties": {"show": lit("false")}}]}
    v["visualContainerObjects"] = vc
    v["drillFilterOtherVisuals"] = True
    x, y, w, h = pos
    return {"position": {"x": x, "y": y, "z": 0, "width": w, "height": h}, "visual": v}


def textbox(pos, runs):
    paras = [{"textRuns": [{"value": t, "textStyle": s}]} for t, s in runs]
    v = visual("textbox", pos, objects={"general": [{"properties": {"paragraphs": paras}}]})
    del v["visual"]["visualContainerObjects"]
    return v


LABELS = {"labels": [{"properties": {"show": lit("true")}}]}
def colour(c): return {"dataPoint": [{"properties": {"fill": {"solid": {"color": lit(f"'{c}'")}}}}]}
def header(text):
    return textbox((20, 10, 1240, 50), [(text, {"fontSize": "22pt", "fontWeight": "bold", "color": "#1B2A41"})])
def slicer(pos, field, title, dropdown=False):
    obj = {"data": [{"properties": {"mode": lit("'Dropdown'")}}]} if dropdown else None
    return visual("slicer", pos, title, {"Values": [col(field)]}, objects=obj)
def card(pos, measure, title):
    return visual("card", pos, title, {"Values": [mea(measure)]})
def combo(pos, title, cat, y, y2, sort):
    return visual("lineClusteredColumnComboChart", pos, title,
                  {"Category": [col(cat)], "Y": [mea(y)], "Y2": [mea(y2)]}, sort,
                  {**LABELS, **colour(BLUE)})

pages = {
    "about": ("About Us", [
        header("MovieMetrics: Box Office Performance and Profitability"),
        textbox((20, 90, 1240, 230), [
            ("About this dashboard", {"fontSize": "16pt", "fontWeight": "bold"}),
            ("MovieMetrics helps studio executives, producers, investors, directors and marketing managers see which films made money in the past, so they can plan new films with less risk.", {"fontSize": "13pt"}),
            ("Dataset: 3,907 movies after cleaning (3,974 raw records), released between 1929 and 2016, with 16 original attributes and 6 derived variables.", {"fontSize": "13pt"}),
            ("Pages: Main Overview - Financial and Award Analysis - Talent Analysis - Film Details", {"fontSize": "13pt"}),
            ("How to use: pick a genre, decade, budget level or IMDb level in the slicers on the left of each page. Click any bar or slice to filter the other charts. Hover over a chart to see exact values.", {"fontSize": "13pt"}),
        ]),
        card((20, 330, 300, 120), "Total Movies", "Movies"),
        card((340, 330, 300, 120), "Total Box Office", "Total Box Office"),
        card((660, 330, 300, 120), "Total Earnings", "Total Earnings"),
        card((980, 330, 280, 120), "Median ROI", "Median ROI"),
    ]),
    "overview": ("Main Overview", [
        header("Main Overview"),
        slicer((10, 70, 190, 300), "Genre", "Genre"),
        slicer((10, 380, 190, 510), "Decade", "Decade"),
        card((210, 70, 205, 100), "Total Movies", "Total Movies"),
        card((423, 70, 205, 100), "Total Box Office", "Total Box Office"),
        card((636, 70, 205, 100), "Total Earnings", "Total Earnings"),
        card((849, 70, 205, 100), "Median ROI", "Median ROI"),
        card((1062, 70, 208, 100), "Loss %", "Loss %"),
        visual("clusteredBarChart", (210, 180, 520, 430), "Total Box Office by Genre",
               {"Category": [col("Genre")], "Y": [mea("Total Box Office")]},
               [(mea("Total Box Office"), "Descending")], {**LABELS, **colour(BLUE)}),
        visual("clusteredBarChart", (740, 180, 530, 430), "Median ROI by Genre (genres with 25+ films)",
               {"Category": [col("Genre")], "Y": [mea("Median ROI (25+ films)")]},
               [(mea("Median ROI (25+ films)"), "Descending")], {**LABELS, **colour(BLUE)}),
        combo((210, 620, 1060, 270), "Median Budget and Median ROI by Decade (ROI shown for decades with 25+ films)",
              "Decade", "Median Budget", "Median ROI (25+ films)", [(col("Decade"), "Ascending")]),
    ]),
    "financial": ("Financial and Award Analysis", [
        header("Financial and Award Analysis"),
        slicer((10, 70, 190, 300), "Genre", "Genre"),
        slicer((10, 380, 190, 160), "Budget Category", "Budget Category"),
        slicer((10, 550, 190, 160), "IMDb Category", "IMDb Category"),
        combo((210, 70, 520, 310), "Loss % and Median ROI by Budget Category",
              "Budget Category", "Loss %", "Median ROI", [(col("Budget Category"), "Ascending")]),
        visual("scatterChart", (740, 70, 530, 310), "Budget against Box Office",
               {"Category": [col("Movie")], "X": [mea("Total Budget")], "Y": [mea("Total Box Office")]},
               objects=colour(BLUE)),
        visual("donutChart", (210, 390, 340, 320), "Share of Films by IMDb Category",
               {"Category": [col("IMDb Category")], "Y": [mea("Total Movies")]},
               [(col("IMDb Category"), "Ascending")],
               {"labels": [{"properties": {"show": lit("true"), "labelStyle": lit("'Category, percent of total'")}}]}),
        visual("clusteredColumnChart", (560, 390, 340, 320), "Median ROI by IMDb Category",
               {"Category": [col("IMDb Category")], "Y": [mea("Median ROI")]},
               [(col("IMDb Category"), "Ascending")], {**LABELS, **colour(BLUE)}),
        combo((910, 390, 360, 320), "Earnings and Loss % by Award Status",
              "Award Status", "Median Earnings", "Loss %", [(col("Award Status"), "Ascending")]),
    ]),
    "talent": ("Talent Analysis", [
        header("Talent Analysis"),
        slicer((10, 70, 190, 300), "Genre", "Genre"),
        slicer((10, 380, 190, 90), "Director", "Director", dropdown=True),
        slicer((10, 480, 190, 90), "Actor 1", "Lead Actor", dropdown=True),
        visual("clusteredBarChart", (210, 70, 520, 640), "Top 10 Directors by Total Box Office",
               {"Category": [col("Director")], "Y": [mea("Top 10 Director Box Office")]},
               [(mea("Top 10 Director Box Office"), "Descending")], {**LABELS, **colour(BLUE)}),
        visual("clusteredBarChart", (740, 70, 530, 640), "Top 10 Lead Actors by Total Box Office",
               {"Category": [col("Actor 1")], "Y": [mea("Top 10 Actor Box Office")]},
               [(mea("Top 10 Actor Box Office"), "Descending")], {**LABELS, **colour(ORANGE)}),
    ]),
    "details": ("Film Details", [
        header("Film Details"),
        slicer((10, 70, 190, 300), "Genre", "Genre"),
        slicer((10, 380, 190, 90), "Director", "Director", dropdown=True),
        visual("tableEx", (210, 70, 1060, 640), "Films",
               {"Values": [col(c) for c in ["Movie", "Genre", "Release year", "Director", "Budget",
                                            "Box Office", "Earnings", "ROI", "IMDb score", "Profit Status"]]},
               [(col("Box Office"), "Descending")]),
    ]),
}

def hist(pos, title, band, vtype="clusteredColumnChart"):
    return visual(vtype, pos, title, {"Category": [col(band)], "Y": [mea("Total Movies")]}, [(col(band), "Ascending")],
                  {**LABELS, **colour(BLUE), "categoryAxis": [{"properties": {"axisType": lit("'Categorical'")}}]})

# Chapter 3 EDA page: counts and distributions only, kept out of the dashboard navigation
pages["eda"] = ("EDA (Chapter 3)", [
    header("Data Overview (EDA)"),
    visual("clusteredBarChart", (10, 70, 620, 380), "Number of Films by Genre",
           {"Category": [col("Genre")], "Y": [mea("Total Movies")]}, [(mea("Total Movies"), "Descending")], {**LABELS, **colour(BLUE)}),
    hist((650, 70, 620, 380), "Number of Films by Decade", "Decade"),
    hist((10, 460, 620, 230), "Budget Distribution (films per $20M band, $M)", "Budget Band ($M)"),
    hist((650, 460, 620, 230), "Box Office Distribution (films per $100M band, $M; 1000 = $1bn+)", "Box Office Band ($M)"),
    hist((10, 700, 620, 230), "IMDb Score Distribution (films per 0.5 band)", "IMDb Band"),
    hist((650, 700, 620, 230), "Running Time Distribution (films per 20-minute band)", "Running Time Band"),
    visual("donutChart", (10, 940, 620, 230), "Films by Profit Status",
           {"Category": [col("Profit Status")], "Y": [mea("Total Movies")]}, [(mea("Total Movies"), "Descending")],
           {"labels": [{"properties": {"show": lit("true"), "labelStyle": lit("'Category, percent of total'")}}]}),
    hist((650, 940, 620, 230), "Films by Number of Oscar and Golden Globe Nominations", "Oscar and Golden Globes nominations"),
])
PAGE_HEIGHT = {"overview": 900, "eda": 1180}

def nav_button(pos, target, label, current):
    # shortcut: buttons only navigate; in Power BI Desktop edit mode they need Ctrl+click
    v = visual("actionButton", pos, objects={
        "icon": [{"properties": {"shapeType": lit("'blank'")}, "selector": {"id": "default"}}],
        # the on/off toggle lives in a selector-less entry; styling lives under the "default" state
        "text": [{"properties": {"show": lit("true")}},
                 {"properties": {"text": lit(f"'{label}'"), "fontColor": {"solid": {"color": lit("'#FFFFFF'")}},
                                 "fontSize": lit("10D")}, "selector": {"id": "default"}}],
        "fill": [{"properties": {"show": lit("true")}},
                 {"properties": {"fillColor": {"solid": {"color": lit("'#1B2A41'" if current else f"'{BLUE}'")}},
                                 "transparency": lit("0D")}, "selector": {"id": "default"}}]})
    v["visual"]["visualContainerObjects"] = {"visualLink": [{"properties": {
        "show": lit("true"), "type": lit("'PageNavigation'"), "navigationSection": lit(f"'{target}'")}}]}
    return v

DASH = [p for p in pages if p != "eda"]
for pid in DASH:
    visuals = pages[pid][1]
    big = pid == "about"
    w, h, gap = (240, 50, 10) if big else (118, 32, 6)
    x0, y = (20, 480) if big else (1270 - 5 * w - 4 * gap, 14)
    for i, target in enumerate(DASH):
        name = pages[target][0]
        short = {"about": "About Us", "overview": "Overview", "financial": "Financial", "talent": "Talent", "details": "Film Details"}
        visuals.append(nav_button((x0 + i * (w + gap), y, w, h), target, name if big else short[target], target == pid))

theme = {"name": "MovieMetrics", "dataColors": [BLUE, ORANGE, GREY, "#7FA7CF", "#F2B880", "#A3AAB0", "#2E8B8B", "#8C6BB1"],
         "background": "#FFFFFF", "foreground": "#1B2A41", "tableAccent": BLUE}

if __name__ == "__main__":
    # semantic model folder is kept so Power BI's data cache (.pbi) survives a rebuild
    shutil.rmtree(os.path.join(HERE, "MovieMetrics.Report"), ignore_errors=True)
    write("MovieMetrics.pbip", {"version": "1.0", "artifacts": [{"report": {"path": "MovieMetrics.Report"}}],
                                "settings": {"enableAutoRecovery": True}})
    write("MovieMetrics.SemanticModel/definition.pbism", {"version": "1.0", "settings": {}})
    write("MovieMetrics.SemanticModel/model.bim", model)
    write("MovieMetrics.Report/definition.pbir", {"version": "4.0", "datasetReference": {"byPath": {"path": "../MovieMetrics.SemanticModel"}}})
    write("MovieMetrics.Report/definition/version.json", {"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    write("MovieMetrics.Report/StaticResources/SharedResources/BaseThemes/MovieMetrics.json", theme)
    write("MovieMetrics.Report/definition/report.json", {
        "$schema": f"{SCHEMA}/report/1.0.0/schema.json",
        "themeCollection": {"baseTheme": {"name": "MovieMetrics", "reportVersionAtImport": "5.59", "type": "SharedResources"}},
        "layoutOptimization": "None",
        "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                              "items": [{"name": "MovieMetrics", "path": "BaseThemes/MovieMetrics.json", "type": "BaseTheme"}]}]})
    write("MovieMetrics.Report/definition/pages/pages.json", {"$schema": f"{SCHEMA}/pagesMetadata/1.0.0/schema.json",
                                                             "pageOrder": list(pages), "activePageName": "overview"})
    for pid, (name, visuals) in pages.items():
        write(f"MovieMetrics.Report/definition/pages/{pid}/page.json",
              {"$schema": f"{SCHEMA}/page/1.0.0/schema.json", "name": pid, "displayName": name,
               "displayOption": "FitToPage", "height": PAGE_HEIGHT.get(pid, 720), "width": 1280})
        for i, v in enumerate(visuals):
            vid = f"{pid}_v{i:02d}"
            write(f"MovieMetrics.Report/definition/pages/{pid}/visuals/{vid}/visual.json",
                  {"$schema": f"{SCHEMA}/visualContainer/1.0.0/schema.json", "name": vid, **v})
    print("written to", HERE)
