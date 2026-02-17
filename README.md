# Bin-Method Calculator — RTU Cooling Energy

A client-side engineering calculator that estimates annual cooling energy usage for rooftop air conditioning units (RTUs) using the bin method of analysis. It compares a high-efficiency candidate unit against a standard-efficiency baseline and provides economic metrics for the investment.

**[Launch the Calculator](https://m-jim-d.github.io/bin-method-calcs/Controls.html)**

## Overview

The calculator simulates hour-by-hour cooling energy for two rooftop units across all temperature bins for a selected U.S. city. It accounts for:

- Sensible and latent load modeling with building-type-specific load profiles
- Multi-stage, single-stage, and variable-speed compressor operation
- Economizer integration (when enabled)
- Psychrometric corrections to capacity and power at each temperature bin
- Occupied and unoccupied schedule periods with setback

Economic results include life-cycle cost (LCC), simple and discounted payback, rate of return (ROR), and savings-to-investment ratio (SIR).

## Project Structure

```
Controls.html            Main page — form UI and results display
controls.js              Page-level logic — form handling, results rendering, charts
help_viewer.js           Help modal overlay
load_header_footer.js    Dynamic header/footer injection

engine/
  engine_module.js       Core bin-method engine (ES module)
  performance_module.js  Correction factors, S/T ratio, staging, fan/condenser power
  psychro.js             Psychrometric functions (humidity ratio, wet bulb, BPF/ADP)
  database_module.js     Weather data access
  classes.js             Data classes (StageState, StagePair, SystemProperties)

data/
  stations.json          U.S. weather station design conditions
  Tbins_new.json         Temperature bin hours by station

shared/                  CSS stylesheets
docs/                    PDF documentation (user manual, enhancement plan)
methods/                 Engineering methods pages describing the calculations
```

## Documentation

- **[User Manual](docs/PNNL-24130.pdf)** — Detailed guide to the calculator's features and inputs
- **[Quick Start](quickstart.html)** — Step-by-step instructions for running a comparison
- **[Engineering Methods](methods/methods_outline.html)** — Description of the calculation methods
- **[Revision History](RevisionHistory.html)** — Version changelog

## Technical Notes

- All calculations run client-side in the browser using ES modules — no server required.
- Charts are rendered with [Google Charts](https://developers.google.com/chart).
- Weather data covers U.S. cities with ASHRAE design conditions and TMY3 temperature bin distributions.
- The calculator was originally developed as a Classic ASP application at Pacific Northwest National Laboratory (PNNL) and has been ported to client-side JavaScript.

## References

- PNNL-24130: *RTU Comparison Calculator User Manual*
- PNNL-23239 Rev 1: *Advanced Rooftop Unit Control Retrofit — Enhancement Plan*

## License

This project is provided as a demonstration tool for preliminary analysis. Its output should not be used for final purchasing decisions.
