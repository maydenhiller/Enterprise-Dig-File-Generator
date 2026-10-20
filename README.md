# Enterprise Dig File Generator

Streamlit app: upload Enterprise dig packages (.xlsm), the staking report template, a cheat
sheet template, the pipeline KMZ and, optionally, Enterprise's Excavation Survey Report (.pdf)
and Profile (.xlsx) templates. Get a filled staking report, excavation survey report and
profile for every `Dig (NN)` sheet, plus one cheat sheet per package, in one zip with a
folder per AID.

## How it works
- The template's StakingReport tab is formulas into its own `Dig Sheet` tab. The app copies
  the package's Dig sheet (cached values, cell for cell) into `Dig Sheet`, writes surveyor,
  directions and the aerial (the Staking Notes box keeps the template's text), and leaves all formulas alone. It refreshes each
  formula's cached value so previews are right; Excel also recalculates on open.
- County, legal, tract, alignment sheet, HCA, references and weld distances all come from the
  dig sheet itself - no alignment sheet upload needed.
- KMZ: centerline for the aerial; each dig placemark is matched on **assessment ID + Feature ID**
  (dig numbers repeat across packages) and cross-checked against the dig sheet coordinates.
- Aerial: rendered to the template slot's aspect; image width on the ground is exact.
  U/S and D/S reference flags come from the dig sheet's reference coordinates.
  The dig's own pipeline (matched on Line ID = KMZ PipelineNumber, then name, then nearest) is drawn thicker than the others,
  named on the picture and marked "(this dig)" in the legend; other lines are thinner.
- Cheat sheet Latitude/Longitude are filled from the dig sheet (anomaly row). On the staking report, Lat/Long, Elevation, EDOC, Survey Date and photos are never written. Phone (C12) stays an XLOOKUP.
- All digs in the uploaded packages are generated in one run.

## Output

    AID 115/
      AID 115 - Dig Stake Cheat Sheet.xlsx
      Staking Reports/02-03 Survey Staking Report_Dig#02B.xlsm
      Excavation Survey Reports/ExcavationSurvey Report_Dig 02B.pdf
      Profile Reports/Profile Dig #02B.xlsx
      As-Builts/As-Built Dig #02B.vsdx
    AID 116/ ...

**Excavation Survey Report** is an Adobe XFA form. It is filled by appending an incremental
update to the template (the original bytes, and the signature that lets Adobe Reader fill and
save the form, are untouched), so the dropdowns and the check box stay live. Filled: Pipeline
Code (always 763), Pipeline Name, Pipeline Number (the Line ID), Assessment ID, Assessment
Segment Name, Dig #. The girth weld references, elevations and everything below stay blank and
editable; accuracy stays Sub-Meter. All of it comes from the dig sheet's title line, e.g.
`Line ID 600-601-602 / Asmt ID 116,  East Leg Mainline, 8.625" Kearney to Moberly`.

**Profile** is the template's Blank sheet with Line Name, Segment Name, `Line ID x /ASMT ID y`
and Dig # filled in from the same title line, plus the U/S and D/S weld IDs (U6, U8) and
stationing (U7, U9) read from the package's `Site Ref` tabs; the rest is as the template has it.
A weld not listed on the Site Ref tabs is worked out from the other weld: its stationing is the known weld's plus or minus the joint length, and a warning says so.

**As-Built** (optional .vsdx template) is the template's page with: pipeline / section / LID / AID / dig # in the
title block (narrowed to sit between the two logos, font shrunk per line so it never covers one); the U/S and D/S weld
numbers and stationing on the two weld callouts; the Gibson compass rose turned so north is where it is on the ground
(flow runs up the page, so it is turned by the U/S-to-D/S bearing of the two welds' Site Ref coordinates); and the
lower-right box (county/state, legal, tract, aerial marker, milepost, alignment sheet). D.O.C., As-Built Prepared by and
Date are left for the person to fill in, and the lower-left callouts are untouched. Not checked in Visio itself, only in a
LibreOffice render.

Each excavation PDF is the size of the template (about 4 MB), so a run over 100+ digs makes a
zip of several hundred MB. Use a machine with the memory for it, or upload fewer packages per run.

## Run
    pip install -r requirements.txt
    streamlit run app.py
Mapbox token (directions + Mapbox imagery): `.streamlit/secrets.toml` -> `[mapbox] token = "pk..."`.
Without one, imagery uses Esri and directions are left blank.

## Tests
    EXAMPLES=/path/to/example/files pytest tests
