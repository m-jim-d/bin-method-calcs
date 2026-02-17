# Bin-Method Calculator JS Port — Architecture & Calculation Flow

This document describes the high-level architecture of the JavaScript port of the
Bin-Method Calculator. It is intended to make the calculation paths easier to
follow than reading the code alone.

For the original ASP architecture, see `readme.txt`.

---

## File Map

```
Controls.html          Main page: form UI, layout, CSS
controls.js            Page-level script (extracted from Controls.html inline <script>)
help_viewer.js         Help modal (opens Help_Controls.html in an iframe overlay)
google_charts_rtucc.js Chart prototypes — ASP-only; NOT used by the JS version

include/
  engine_module.js     Core bin-method engine (the "Engine.asp" equivalent)
  performance_module.js  Correction factors, S/T ratio, staging, fan/condenser power
  psychro.js           Psychrometric functions (humidity ratio, wet bulb, BPF/ADP)
  database_module.js   Weather data access (Stations.json, Tbins_new.json)
  classes.js           Data classes: StageState, StagePair, SystemProperties
```

### ASP-only files (development/testing environment)

```
Controls.asp           ASP version of the controls page
Engine.asp             ASP version of the engine (server-side calculation + rendering)
include/*.asp          ASP include files (psychro, performance, database, classes, misc)
google_charts_rtucc.js Used only by Engine.asp for chart rendering
```

---

## How ASP Worked vs. How JS Works

### ASP (server-side)

```
User fills form → Submit → POST to Engine.asp
                           Engine.asp runs calculations server-side (VBScript)
                           Engine.asp renders results HTML + chart data
                           Browser receives complete results page
```

### JS (client-side)

```
User fills form → Submit → controls.js calls submitToEngine()
                           submitToEngine() dynamically imports engine_module.js
                           engine_module.js runs calculations in the browser
                           engine_module.js returns JSON results
                           controls.js builds results HTML from JSON
                           Google Charts renders charts from the JSON data
```

Key difference: In ASP, the engine produces HTML directly. In JS, the engine
produces structured JSON, and `controls.js` builds the HTML.

---

## Main Calculation Flow

### 1. Page Load

```
Controls.html loads
  └─ controls.js executes (via <script defer>)
       ├─ IIFEs populate dropdowns (Total Capacity 36-360, Degradation 0-50)
       └─ loadCityData()
            └─ .then → recalcVentilation()    ← "Establish_SIV" equivalent
```

`recalcVentilation()` is the JS equivalent of ASP's `Establish_SIV`. It runs
the engine iteratively (up to 4 times) to converge on a ventilation rate that
is consistent with the unit's actual sensible capacity at design conditions.

### 2. User Interaction

Form changes trigger `onChangeHandler(controlName)` which:
- Updates dependent fields (e.g., changing capacity recalculates fan power)
- Highlights non-default values via `checkForDefaults()`
- Calls `recalcVentilation()` when inputs affect ventilation (capacity, building
  type, S/T ratio, humidity settings)

### 3. Submit (the main calculation)

```
submitToEngine()                                    [controls.js]
  │
  ├─ import engine_module.js
  │
  ├─ engine.exportBinCalcsJson(form, opts)          [engine_module.js]
  │    │
  │    ├─ runBinCalcs(form, opts)                   ← the core engine
  │    │    │
  │    │    ├─ Phase 1: Input Collection & Validation
  │    │    │    ├─ Read all form values with _pickNumber() fallback chains
  │    │    │    ├─ Compute BPF/ADP for both units (psychro.js)
  │    │    │    ├─ Build SystemProperties objects (_systemFromForm)
  │    │    │    ├─ Parse spreadsheet data if provided (ParseAndModel)
  │    │    │    └─ Load weather data (Stations.json, Tbins_new.json)
  │    │    │
  │    │    ├─ Phase 2: Load Line
  │    │    │    ├─ computeLoadLine()
  │    │    │    │    ├─ getDesignConditions() → ODB, OWB, elevation, pressure
  │    │    │    │    ├─ Psychrometric calcs at design (OHR, IHR, IRH)
  │    │    │    │    ├─ Mixed-air entering conditions (Mixer2)
  │    │    │    │    ├─ S/T ratio at entering conditions
  │    │    │    │    ├─ Sensible capacity at design
  │    │    │    │    └─ slope & intercept of non-ventilation load line
  │    │    │    │
  │    │    │    └─ Or use locked load line from form if chkLockLoadLine
  │    │    │
  │    │    ├─ Phase 3: Bin-by-Bin Energy Calculation
  │    │    │    ├─ runOneSystem(candidateSD)
  │    │    │    │    ├─ For each occupied-hour temperature bin:
  │    │    │    │    │    ├─ Compute loads (non-vent + vent + economizer)
  │    │    │    │    │    ├─ Mixed-air conditions at this bin's ODB
  │    │    │    │    │    ├─ Staging / runtime (StageLevel or integrated econ)
  │    │    │    │    │    ├─ Condenser energy (CondenserPower_PL_kW × hours × RT)
  │    │    │    │    │    ├─ Fan energy (FanPower_PL_kW × hours, mode-dependent)
  │    │    │    │    │    └─ Aux energy (flat kW × hours)
  │    │    │    │    │
  │    │    │    │    ├─ For each unoccupied-hour bin (setback):
  │    │    │    │    │    └─ Same as above but with IDB+setback, fan cycles
  │    │    │    │    │
  │    │    │    │    └─ Sum annual energy: condenser + fan + aux
  │    │    │    │
  │    │    │    └─ runOneSystem(standardSD)
  │    │    │         └─ (same process for the standard unit)
  │    │    │
  │    │    └─ Phase 4: Results Assembly
  │    │         ├─ Scale energy by nUnits
  │    │         ├─ Compute demand costs
  │    │         └─ Store in _lastRunPhase1 (module-level variable)
  │    │
  │    └─ exportBinCalcsJson reads _lastRunPhase1 + _lastRunInputs
  │       └─ Returns structured JSON with: annual energy, economics,
  │          design conditions, per-bin details, input snapshot, overrides
  │
  ├─ buildResultsHTML(jsJson)                       [controls.js]
  │    ├─ Economics table (energy, costs, LCC, payback, ROR, SIR)
  │    ├─ Bin tables (loads/conditions + performance, if "Show Details")
  │    ├─ Parameter summary
  │    └─ Spreadsheet model summaries (if spreadsheet data provided)
  │
  └─ Draw charts (Google Charts)
       ├─ drawPaybackChart()
       ├─ drawAllBinCharts() → drawBinLoadsChart() + drawBinPerfChart()
       └─ (grid view available via toggleBinChartsGrid)
```

---

## Key Concepts

### Load Line

The non-ventilation sensible load line is a linear equation:

```
NonVentLoad = slope × (ODB − IDB) + intercept
```

It represents how the building's cooling load (excluding ventilation) varies
with outdoor temperature. The slope and intercept are derived from the building
type model and the unit's sensible capacity at design conditions.

The load line can be **locked** (user provides slope/intercept directly), which
is used for custom building models or sensitivity analysis.

### Ventilation Refinement (Establish_SIV)

The ventilation rate affects the load line, which affects the sensible capacity
at design, which affects the ventilation rate. This circular dependency is
resolved by iteration:

```
1. Run engine with initial ventilation estimate
2. Extract sensibleCapacityDesign from load line
3. Refine ventilation using refineVentilationFromCapacity()
4. Repeat until ventilation converges (≤0.1 change) or 4 iterations
```

### Staging

The engine supports single-stage, multi-stage, and variable-speed compressors.
`StageLevel()` determines how much of the hour the unit runs at each stage,
based on the ratio of load to capacity. The `StagePair` object tracks runtime,
flow fraction, and capacity fraction for up to two adjacent stages.

### Economizer

When enabled, the economizer provides free cooling by increasing outdoor air
flow when ODB < IDB. The engine computes:
1. Economizer-only cooling capacity
2. If insufficient, integrated mode (economizer + DX at first stage)
3. If integrated fails, falls back to DX-only

### Correction Factors

At each temperature bin, the engine applies correction factors to rated capacity
and power:
- **TCF** — Total Capacity Factor (adjusts rated capacity for conditions)
- **ECF** — Efficiency Correction Factor (adjusts rated power input)
- **PCF** — Power Correction Factor (= TCF/ECF)
- **OCF** — Overall Correction Factor (= PCF / (TCF × S/T_CF))
- **S/T** — Sensible-to-Total ratio at entering conditions

These come from either DOE-2 curves or manufacturer spreadsheet regression
models, depending on configuration.

### BPF/ADP

The Bypass Factor (BPF) and Apparatus Dew Point (ADP) are psychrometric
properties computed from the unit's rated S/T ratio, capacity, and airflow.
They are used to determine the coil's dehumidification performance. A BPF/ADP
calculation failure (common at high S/T ratios like 0.80) produces the
"Try lowering the S/T ratio" warning.

---

## Module Communication

```
controls.js                          engine_module.js
───────────                          ────────────────
                                     Module-level state:
                                       _lastRunPhase1
                                       _lastRunInputs
                                       _lastLoadLine
                                       _dataCache

submitToEngine()
  └─ engine.exportBinCalcsJson(form)
       └─ runBinCalcs(form)
            ├─ sets _lastLoadLine     ← available via getLastLoadLine()
            ├─ sets _lastRunInputs
            └─ sets _lastRunPhase1
       └─ reads _lastRunPhase1
       └─ reads _lastRunInputs
       └─ returns JSON ──────────────→ used by buildResultsHTML()

recalcVentilation()
  └─ engine.exportBinCalcsJson(form)
  └─ engine.getLastLoadLine() ───────→ reads _lastLoadLine
       └─ extracts sensibleCapacityDesign for ventilation refinement

globalThis.setFanPowerValuesAndDefaults
  └─ defined in controls.js
  └─ read by engine_module.js (_fanPowerDefaultFromPage) to extract
     fan power coefficients from the function's source code
```

---

## Economics (in controls.js, not the engine)

The engine returns raw annual energy (kWh) for each unit. All economic
calculations happen in `buildResultsHTML()`:

- **Annual Operating Cost** = (energy × elec rate) + maintenance + demand cost
- **LCC** = (unit cost × $1000) + (annual cost × UPV)
- **Annualized Cost** = LCC / UPV
- **NPV** = LCC_standard − LCC_candidate
- **Simple Payback** = capital cost difference / annual savings
- **Discounted Payback** = Newton iteration to find life where NPV = 0
- **ROR** = Newton iteration to find discount rate where NPV = 0
- **SIR** = (discounted annual savings) / capital cost difference

---

## ASP ↔ JS Parity Testing

The ASP version remains in the development environment for parity testing.
Both versions should produce identical results for the same inputs. Key
comparison points:
- Annual energy (condenser, fan, aux) for both units
- Load line slope and intercept
- Per-bin energy values in the detailed tables
- Economic results (LCC, payback, ROR, SIR)

The `runInputs` object captured by the engine includes the source of every
input value, making it possible to trace discrepancies to specific form
field resolution differences between ASP and JS.
