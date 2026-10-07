# Enterprise Dig File Generator

Streamlit app: upload Enterprise dig packages (.xlsm), the staking report template,
a cheat sheet template and the pipeline KMZ. Get one filled staking report per
`Dig (NN)` sheet plus one cheat sheet per package.

## How it works
- The template's StakingReport tab is formulas into its own `Dig Sheet` tab. The app copies
  the package's Dig sheet (cached values, cell for cell) into `Dig Sheet`, writes surveyor,
  directions, staking notes and the aerial, and leaves all formulas alone. It refreshes each
  formula's cached value so previews are right; Excel also recalculates on open.
- County, legal, tract, alignment sheet, HCA, references and weld distances all come from the
  dig sheet itself - no alignment sheet upload needed.
- KMZ: centerline for the aerial; each dig placemark is matched on **assessment ID + Feature ID**
  (dig numbers repeat across packages) and cross-checked against the dig sheet coordinates.
- Aerial: rendered to the template slot's aspect; image width on the ground is exact.
  U/S and D/S reference flags come from the dig sheet's reference coordinates.
- Cheat sheet Latitude/Longitude are filled from the dig sheet (anomaly row). On the staking report, Lat/Long, Elevation, EDOC, Survey Date and photos are never written. Phone (C12) stays an XLOOKUP.
- All digs in the uploaded packages are generated in one run.
- Output names: `AID 115 - Dig 02B_Staking Report.xlsm`, `AID 115 - Dig Stake Cheat Sheet.xlsx`.

## Run
    pip install -r requirements.txt
    streamlit run app.py
Mapbox token (directions + Mapbox imagery): `.streamlit/secrets.toml` -> `[mapbox] token = "pk..."`.
Without one, imagery uses Esri and directions are left blank.

## Tests
    EXAMPLES=/path/to/example/files pytest tests
