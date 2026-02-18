// ============================================================
// engine_module.js — Main calculation engine for the bin-method tool
// ============================================================
//
// Orchestrates the full bin-method calculation:
//   1. Reads form inputs and builds SystemProperties for Candidate & Standard
//   2. Loads weather/station JSON data (cached in memory)
//   3. Computes the non-ventilation load line (computeLoadLine)
//   4. Loops over temperature bins (occupied + unoccupied) to compute
//      staging, economizer, energy, and demand for each unit (runOneSystem)
//   5. Assembles economics (LCC, payback, ROR, SIR) and returns a JSON
//      result object via exportBinCalcsJson()
//
// Converted from Engine.asp + Controls.asp (VBScript/ASP).
// See ARCHITECTURE.md for the full module dependency graph and data flow.
//
// Key exports:
//   exportBinCalcsJson(form, opts) — run engine, return structured JSON
//   computeLoadLine({...})         — compute non-ventilation load line
//   runBinCalcs(form, opts)        — internal: full engine pass
//   getLastLoadLine()              — retrieve load line from last run
// ============================================================

import { Phr_wb, Phr_rh, Prh_hr, Pwb_hr, Ph_hr, dblStandardPressure, dblKWtoKBTUH, BPF_ADP_SC, A0_FromBPF } from './psychro.js';
import {
  IRH_Track_OR_Set,
  FF,
  FanPower_PL_kW,
  CondenserPower_PL_kW,
  Mixer2,
  NetSenCap_Stage_Adjusted_KBtuh,
  SensVentLoad,
  StageLevel,
  ST_Ratio_engine,
  Tot_Capacity_Correction,
} from './performance_module.js?v=14';
import { getDesignConditions, getWeatherRecords, getHours } from './database_module.js?v=1';
import { StageState, StagePair, SystemProperties } from './classes.js?v=6';

const ENGINE_VERSION = 14;

// Module-level state (replaces globalThis globals for internal communication)
let _lastRunPhase1 = null;   // set by runBinCalcs, read by exportBinCalcsJson
let _lastRunInputs = null;   // set by runBinCalcs, read by exportBinCalcsJson
let _lastLoadLine = null;    // set by runBinCalcs, read externally via getLastLoadLine()
const _dataCache = {};       // fetch cache for JSON data files

/** @returns the load line result from the most recent engine run */
export function getLastLoadLine() { return _lastLoadLine; }

const DEFAULTS = {
  ST_Ratio_AtTest_C: 0.72,
  N_Affinity: 2.5,
  BFn_slope_kW_per_kBtuh: 0.0132,
  BFn_int_kW: -0.2283,
  TrackOHR: 'on',
  VentilationUnits: '% of Fan Cap.',
  VentilationValue: 10,
  IRH_pct: 60,
  SandI_fraction: 0.5,
  DOE2_Curves: 'on',
  Specific_RTU_C: 'None',
  dblCFMperTon: 400,
};

function _escapeHtml(s) {
  const str = String(s ?? '');
  return str
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function _el(form, name) {
  try {
    if (!form) return null;
    const els = form.elements;
    if (!els) return null;
    const byName = els.namedItem ? els.namedItem(name) : null;
    if (byName) return byName;
    return els[name] || null;
  } catch {
    return null;
  }
}

function _val(form, name) {
  const el = _el(form, name);
  if (!el) return '';
  const v = el.value;
  return v === null || v === undefined ? '' : String(v);
}

function _valOrDefault(form, name, fallback) {
  const v = _val(form, name);
  return v !== '' ? v : fallback;
}

function _checked(form, name) {
  const el = _el(form, name);
  return !!(el && el.checked);
}

function _checkedOrDefault(form, name, fallbackChecked) {
  const el = _el(form, name);
  if (!el) return !!fallbackChecked;
  return !!el.checked;
}

function _asNumber(form, name) {
  const raw = _val(form, name);
  if (raw == null || String(raw).trim() === '') return NaN;
  const n = Number(raw);
  return Number.isFinite(n) ? n : NaN;
}

function _pickNumber(candidates) {
  const list = Array.isArray(candidates) ? candidates : [];
  for (const c of list) {
    const v = c?.value;
    if (v === null || v === undefined) continue;
    if (typeof v === 'string' && v.trim() === '') continue;
    const n = _num(v);
    if (n !== null) return { source: String(c?.source ?? ''), value: n };
  }
  return { source: '', value: NaN };
}

function _textById(form, id) {
  try {
    const doc = form?.ownerDocument || globalThis.document;
    const el = doc ? doc.getElementById(id) : null;
    if (!el) return '';
    const t = el.textContent;
    return t === null || t === undefined ? '' : String(t).trim();
  } catch {
    return '';
  }
}

function _valById(form, id) {
  try {
    const doc = form?.ownerDocument || globalThis.document;
    const el = doc ? doc.getElementById(id) : null;
    if (!el) return '';
    const v = el.value;
    return v === null || v === undefined ? '' : String(v);
  } catch {
    return '';
  }
}

function _numFromText(s) {
  const str = String(s ?? '').trim();
  if (!str) return NaN;
  const m = str.match(/-?\d+(?:\.\d+)?/);
  if (!m) return NaN;
  const n = Number(m[0]);
  return Number.isFinite(n) ? n : NaN;
}

function _selectDefaultText(el) {
  try {
    if (!el || !el.options) return '';
    for (const opt of Array.from(el.options)) {
      if (opt && opt.defaultSelected) return String(opt.text ?? opt.value ?? '').trim();
    }
    // Fallback: if nothing is marked defaultSelected, assume first option is default.
    const first = el.options[0];
    return first ? String(first.text ?? first.value ?? '').trim() : '';
  } catch {
    return '';
  }
}

function _fanPowerDefaultFromPage(totalCapKBtuh) {
  try {
    const totalCap = Number(totalCapKBtuh);
    if (!Number.isFinite(totalCap)) return null;

    const fn = globalThis.setFanPowerValuesAndDefaults;
    if (typeof fn !== 'function') return null;
    const src = String(fn);

    const slopeMatch = src.match(/dblBFn_slope_kW_per_kBtuh\s*=\s*([0-9.]+)/i);
    const intMatch = src.match(/dblBFn_int_kW\s*=\s*([-0-9.]+)/i);
    if (!slopeMatch || !intMatch) return null;

    const slope = Number(slopeMatch[1]);
    const intercept = Number(intMatch[1]);
    if (!Number.isFinite(slope) || !Number.isFinite(intercept)) return null;

    const kW = slope * totalCap + intercept;
    return Number.isFinite(kW) ? _roundTo(kW, 3) : null;
  } catch {
    return null;
  }
}

function _fanPowerDefaultFromAspDefaults(totalCapKBtuh) {
  const totalCap = Number(totalCapKBtuh);
  if (!Number.isFinite(totalCap)) return null;
  const kW = DEFAULTS.BFn_slope_kW_per_kBtuh * totalCap + DEFAULTS.BFn_int_kW;
  return Number.isFinite(kW) ? _roundTo(kW, 3) : null;
}

function _num(x) {
  if (x === null || x === undefined) return null;
  if (typeof x === 'string' && x.trim() === '') return null;
  const n = Number(x);
  return Number.isFinite(n) ? n : null;
}

function _roundToEvenInt(n) {
  const x = Number(n);
  if (!Number.isFinite(x)) return n;
  const i = Math.trunc(x);
  const frac = Math.abs(x - i);
  if (frac < 0.5) return i;
  if (frac > 0.5) return i + (x >= 0 ? 1 : -1);
  return (i % 2 === 0) ? i : (i + (x >= 0 ? 1 : -1));
}

function _roundTo(n, decimals) {
  const x = Number(n);
  if (!Number.isFinite(x)) return n;
  const p = Math.pow(10, decimals);
  return _roundToEvenInt(x * p) / p;
}

function _stageLevelArray(nStages) {
  const n = Number(nStages);
  if (n === 1) return [1.0];
  if (n === 2) return [0.5, 1.0];
  if (n === 3) return [0.4, 0.6, 1.0];
  if (n === 5) return [0.2, 0.4, 0.6, 0.8, 1.0];
  if (n === 10) return [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0];
  return [1.0];
}

function _asNumberOr(form, name, fallback) {
  const v = _asNumber(form, name);
  return Number.isFinite(v) ? v : fallback;
}

// Build a SystemProperties object for one unit (unitKey='C' or 'S') by
// reading form inputs, applying defaults, and inferring condenser power
// from EER when not explicitly provided.  Matches ASP FormValues.Establish.
function _systemFromForm(form, unitKey, opts) {
  const totalCap = opts.totalCap;
  const cfm = opts.cfm;
  const advancedControlsChecked = !!opts.advancedControlsChecked;

  // Controls.asp renders the power input controls even when Advanced Controls are hidden,
  // but in that mode they are typically blank and ASP falls back to default relationships.
  // Match ASP: treat power inputs as "provided" only when there is a non-empty posted value.
  const txtBFnName = unitKey === 'C' ? 'txtBFn_kw_C' : 'txtBFn_kw_S';
  const txtAuxName = unitKey === 'C' ? 'txtAux_kw_C' : 'txtAux_kw_S';
  const txtCondName = unitKey === 'C' ? 'txtCond_kw_C' : 'txtCond_kw_S';
  const hasPowerInputs = advancedControlsChecked && (
    String(_val(form, txtBFnName) ?? '').trim() !== '' ||
    String(_val(form, txtAuxName) ?? '').trim() !== '' ||
    String(_val(form, txtCondName) ?? '').trim() !== ''
  );

  const eerVal = _asNumber(form, unitKey === 'C' ? 'txtEER' : 'txtEER_Standard');
  const eer = Number.isFinite(eerVal) ? eerVal : null;

  // Defaults (legacy: FanPower_Default() and cstAux_kW)
  // Note: ASP's non-advanced Standard condenser power uses the *unrounded* default fan power,
  // then rounds the final condenser kW to 3 decimals.
  const defaultBFn_raw = (DEFAULTS.BFn_slope_kW_per_kBtuh * totalCap) + DEFAULTS.BFn_int_kW;
  const defaultBFn = _roundTo(defaultBFn_raw, 3);
  const defaultAux = 0;

  const sd = new SystemProperties();
  sd.SystemName = unitKey === 'C' ? 'Candidate' : 'Standard';
  sd.EER = eer;
  sd.UnitCost = _asNumberOr(form, unitKey === 'C' ? 'txtUnitCost' : 'txtUnitCost_Standard', unitKey === 'C' ? 4.5 : 4.0);
  sd.Maintenance = _asNumberOr(form, unitKey === 'C' ? 'txtMaintenance_Candidate' : 'txtMaintenance_Standard', 0);
  sd.Economizer = _checked(form, unitKey === 'C' ? 'chkEconomizer_C' : 'chkEconomizer_S') ? 'on' : 'off';
  sd.Aux_kw = hasPowerInputs ? _asNumberOr(form, txtAuxName, defaultAux) : defaultAux;
  sd.BFn_kw = hasPowerInputs ? _asNumberOr(form, txtBFnName, defaultBFn) : defaultBFn;
  sd.Cond_kw = hasPowerInputs ? _asNumberOr(form, txtCondName, 0) : 0;
  sd.CondFanPercent = _asNumberOr(form, unitKey === 'C' ? 'txtCondFanPercent_C_hidden' : 'txtCondFanPercent_S_hidden', 0);
  sd.FanControls = _valOrDefault(form, unitKey === 'C' ? 'cmbFanControls_C' : 'cmbFanControls_S', '1-Spd: Always ON');
  // Match ASP: when Advanced Controls are hidden, the cycling degradation factor defaults to 25%.
  // Controls.asp may still post a hidden/blank/0 value for cmbPLDegrFactor_*; treat that as default.
  sd.PLDegrFactor = advancedControlsChecked
    ? _asNumberOr(form, unitKey === 'C' ? 'cmbPLDegrFactor_C' : 'cmbPLDegrFactor_S', 25)
    : 25;
  sd.N_Affinity = opts.nAffinity;
  sd.ST_Ratio_AtTest = opts.stAtTest;
  sd.A0_BPF = opts.a0Bpf;
  sd.DOE2_Curves = opts.doe2Curves;
  sd.Specific_RTU = opts.specificRtu;
  sd.CapacityFraction_Min = 0.15;

  // Spreadsheet data.
  const ssText = _val(form, unitKey === 'C' ? 'txtSpreadsheetData_C' : 'txtSpreadsheetData_S');
  const ssChecked = _checked(form, unitKey === 'C' ? 'chkSpreadsheetControl_C' : 'chkSpreadsheetControl_S');
  sd.Spreadsheet = (ssChecked && ssText) ? String(ssText) : '';

  const nStages = Number(_valOrDefault(form, unitKey === 'C' ? 'cmbNstages_C' : 'cmbNstages_S', 1));
  sd.N_Stages = nStages;
  sd.StageLevels = _stageLevelArray(nStages);

  // If condenser power appears to be missing but we can infer it from nominal capacity/EER.
  if (!Number.isFinite(sd.Cond_kw) || sd.Cond_kw === 0) {
    if (Number.isFinite(sd.EER) && sd.EER > 0) {
      // Match ASP (FormValues.Establish) / Controls.asp defaults:
      // When Advanced Controls are OFF and condenser kW must be inferred,
      // use the *unrounded* default blower-fan kW in the condenser kW derivation,
      // then round the resulting condenser kW to 3 decimals.
      // Using the rounded BFn_kW here can shift Cond_kW by ~0.001 and flip annual whole-kWh rounding.
      const fanKwForCond = !hasPowerInputs ? defaultBFn_raw : sd.BFn_kw;
      sd.Cond_kw = _roundTo((totalCap / sd.EER) - (fanKwForCond + sd.Aux_kw), 3);
    }
  }

  // When Advanced Controls are OFF, condenser fan % is not present; legacy defaults to 9%.
  if (!Number.isFinite(sd.CondFanPercent) || sd.CondFanPercent === 0) {
    sd.CondFanPercent = 9.0;
  }

  // Basic sanity.
  if (!Number.isFinite(sd.PLDegrFactor)) sd.PLDegrFactor = 25;
  if (!Number.isFinite(sd.N_Affinity)) sd.N_Affinity = DEFAULTS.N_Affinity;
  if (!Number.isFinite(sd.BFn_kw)) sd.BFn_kw = 0;
  if (!Number.isFinite(sd.Aux_kw)) sd.Aux_kw = 0;
  if (!Number.isFinite(sd.Cond_kw)) sd.Cond_kw = 0;
  if (!Number.isFinite(sd.EER)) sd.EER = totalCap / Math.max(0.001, (sd.BFn_kw + sd.Aux_kw + sd.Cond_kw));

  // Store these to mimic object-shape expectations in performance module.
  sd.CFM = cfm;
  sd.TotalCap = totalCap;

  return sd;
}

function _modelDebug(objSD) {
  try {
    if (!objSD) return null;

    function packModel(m) {
      if (!m) return null;
      return {
        isSolved: !!m.IsSolved,
        lastSolveError: m.LastSolveError ? String(m.LastSolveError) : '',
        terms: Array.isArray(m._terms) ? m._terms.slice() : null,
        coefficients: Array.isArray(m._coefficients) ? m._coefficients.slice() : null,
        tValues: Array.isArray(m._tValues) ? m._tValues.slice() : null,
        r2: m._r2,
        residualSE: m._residualSE,
      };
    }

    return {
      versionOK: !!(objSD.SSD && objSD.SSD.VersionOK),
      rawSpreadsheet: objSD.Spreadsheet || '',
      effective: {
        eer: Number(objSD.EER),
        bfn_kw: Number(objSD.BFn_kw),
        aux_kw: Number(objSD.Aux_kw),
        cond_kw: Number(objSD.Cond_kw),
      },
      nameplate: {
        evapFanPower: objSD?.SSD?.EvaporatorFanPower,
        condenserPower: objSD?.SSD?.CondenserPower,
        auxPower: objSD?.SSD?.AuxilaryPower,
        totalSystemPower: objSD?.SSD?.TotalSystemPower,
      },
      models: {
        grossCapacity: packModel(objSD.GrossCapacity_KBtu_Model),
        condenserKW: packModel(objSD.Condenser_kW_Model),
        stRatio: packModel(objSD.ST_Ratio_Model),
        neer: packModel(objSD.NEER_PL_Model),
      },
    };
  } catch {
    return null;
  }
}

// Integrated economizer + DX staging for bins where the economizer alone
// cannot meet the load.  Determines the DX runtime fraction needed to
// supplement economizer cooling.  Returns { ok, runtime, debug }.
function _capacityLevelIntegrated_Staged(objSD, SP, BOC, nonVentLoad, idbSetpoint, pressure, totalCap, cfm, ventCFM, ventilationFraction) {
  // Port of legacy CapacityLevel_Integrated() for staged systems.
  // Returns { ok, runtime } where runtime is DX runtime fraction for stage A, or ok:false if integrated fails.
  try {
    SP.IntegratedState = 'Attempting Integrated';

    // Flow fraction when economizer is running (compressor off).
    const flowFractionEcon = FF(objSD, SP.A.CapacityFraction, SP, 'C-Off', true, BOC.DB, ventilationFraction);

    // Positive when colder outside than inside.
    const econCoolingCapMax = -1.0 * SensVentLoad(cfm, BOC.HR, BOC.DB, idbSetpoint, pressure);
    const econCoolingCap = econCoolingCapMax * flowFractionEcon;

    // First-stage capacity fraction.
    const stageArray = Array.isArray(objSD.StageLevels) ? objSD.StageLevels : [1.0];
    const capFraction = Number(stageArray[0] ?? 1.0);
    SP.A.CapacityFraction = capFraction;

    // When DX is running (integrated economizer), fan flow fraction follows the C-On path.
    const flowFractionDX = FF(objSD, capFraction, SP, 'C-On', false, BOC.DB, ventilationFraction);
    SP.A.FlowFraction = flowFractionDX;

    // Approximate integrated sensible cooling capacity at stage A.
    // DX coil: use stage-adjusted net sensible capacity at the current outdoor conditions.
    // Economizer: treat as additional sensible cooling (same basis as econCoolingCapMax) at DX flow.
    const dxCoilCoolingCap = NetSenCap_Stage_Adjusted_KBtuh(
      objSD,
      totalCap,
      cfm,
      SP.A,
      1.0,
      BOC.DB,
      BOC.WB,
      BOC.DB,
      pressure
    );
    const dxEconCoolingCap = econCoolingCapMax * flowFractionDX;

    const dxTotalCoolingCap = dxCoilCoolingCap + dxEconCoolingCap;

    let dxRuntime = 0;
    if (econCoolingCap > nonVentLoad) {
      dxRuntime = 0.0;
    } else {
      const runtime = (nonVentLoad - econCoolingCap) / (dxTotalCoolingCap - econCoolingCap);
      if (runtime >= 0.0 && runtime <= 1.0) {
        // Legacy parity: integrated economizer bins intentionally use LoadFraction=0.0 so
        // CyclingEfficiency() applies maximum cycling degradation for the DX contribution.
        // (Legacy ASP historically arrived at this via bin-to-bin state carryover; ASP has
        // now been made explicit and JS matches that explicit behavior here.)
        SP.A.LoadFraction = 0.0;
        SP.A.RunTime = runtime;
        return {
          ok: true,
          runtime,
          debug: {
            flowFractionEcon,
            econCoolingCapMax,
            econCoolingCap,
            capFraction,
            dxEconCoolingCap,
            dxCoilCoolingCap,
            dxTotalCoolingCap,
          },
        };
      }
    }

    SP.IntegratedState = 'Failed to Satisfy Load';
    return { ok: false, runtime: 0, debug: null };
  } catch {
    SP.IntegratedState = 'Failed to Satisfy Load';
    return { ok: false, runtime: 0, debug: null };
  }
}

// Primary entry point: run the full engine and return a structured JSON
// object containing annual energy, economics, design conditions, bin
// details, and spreadsheet model summaries.  Called by submitToEngine()
// in controls.js.
export async function exportBinCalcsJson(form, opts = {}) {
  // Always recompute so the export reflects the current form state and latest engine logic.
  // (Avoids stale globals causing parity deltas to persist after code changes.)
  const htmlOrOk = await runBinCalcs(form, { ...opts });
  const last = _lastRunPhase1;
  if (!last || !last.resC || !last.resS) {
    // runBinCalcs() returns an HTML string on error; try to surface that root cause.
    let detail = '';
    try {
      if (typeof htmlOrOk === 'string') {
        const m = htmlOrOk.match(/<pre>([\s\S]*?)<\/pre>/i);
        if (m && m[1]) {
          detail = String(m[1])
            .replaceAll('&amp;', '&')
            .replaceAll('&lt;', '<')
            .replaceAll('&gt;', '>')
            .replaceAll('&quot;', '"')
            .replaceAll('&#39;', "'")
            .trim();
        }
      }
    } catch {
      // Ignore parsing failure; fall back to generic message.
    }

    throw new Error(
      'exportBinCalcsJson: engine did not produce lastRunBinCalcsPhase1' + (detail ? `\nEngine error: ${detail}` : '')
    );
  }

  const inp = _lastRunInputs || {};

  function _pushOverride(dst, key, value, defaultValue, source) {
    try {
      if (!dst || !key) return;
      const v = value;
      const d = defaultValue;
      const same = Number.isFinite(Number(v)) && Number.isFinite(Number(d))
        ? (Number(v) === Number(d))
        : (String(v ?? '') === String(d ?? ''));
      if (same) return;
      dst[key] = {
        value: v,
        default: d,
        source: source || undefined,
      };
    } catch {
    }
  }

  function _computeOverrides() {
    const o = {};

    const fd = (inp && inp.formDefaults && typeof inp.formDefaults === 'object') ? inp.formDefaults : {};

    function _finiteOrUndef(n) {
      return Number.isFinite(Number(n)) ? Number(n) : undefined;
    }

    function _normVentUnitsLabel(s) {
      const t = String(s ?? '').trim();
      if (!t) return '';
      // Controls.asp default cell sometimes renders just "of Fan Cap"; normalize to the select label.
      return (/fan\s*cap/i.test(t) && !/%/i.test(t)) ? '% of Fan Cap.' : t;
    }

    _pushOverride(o, 'totalCap_kBtuh', inp.totalCap, fd.totalCap_kBtuh_default, undefined);
    _pushOverride(o, 'idb_F', inp.idb, fd.idb_F_default, undefined);
    _pushOverride(o, 'nUnits', inp.nUnits, fd.nUnits_default, undefined);

    // Compare against the Controls.asp form defaults first; fall back to DEFAULTS when the element isn't present.
    const doe2DefaultChecked = (typeof fd.chkDOE2_Curves_defaultChecked === 'boolean')
      ? fd.chkDOE2_Curves_defaultChecked
      : (DEFAULTS.DOE2_Curves === 'on');
    const doe2DefaultLabel = doe2DefaultChecked ? 'DOE2' : 'Carrier';
    _pushOverride(o, 'DOE2_Curves', inp.doe2Curves, doe2DefaultLabel, inp.doe2CurvesSource);

    const specificRtuDefault = (fd.cmbSpecific_RTU_C_defaultValue !== undefined)
      ? fd.cmbSpecific_RTU_C_defaultValue
      : DEFAULTS.Specific_RTU_C;
    _pushOverride(o, 'Specific_RTU_C', inp.specificRtu_C, specificRtuDefault, inp.specificRtuSource_C);

    // Candidate fan controls (UI label).
    {
      const fanControlsDefault = String(fd.cmbFanControls_C_defaultValue ?? '').trim();
      if (fanControlsDefault) {
        _pushOverride(o, 'FanControls_C', inp.fanControls_C, fanControlsDefault, inp.fanControlsSource_C);
      }
    }

    // Candidate power inputs at rating conditions.
    // When Advanced Controls are hidden, these <td> defaults are not rendered and parse as NaN.
    // Only compare/push when we have a real finite default.
    {
      const bfnDef = _finiteOrUndef(fd.tdBFn_kw_C_defaultValue);
      if (bfnDef !== undefined) _pushOverride(o, 'BFn_kw_C', inp.bfnKw_C, bfnDef, inp.bfnSource_C);
      const condDef = _finiteOrUndef(fd.tdCond_kw_C_defaultValue);
      if (condDef !== undefined) _pushOverride(o, 'Cond_kw_C', inp.condKw_C, condDef, inp.condSource_C);
    }

    // Ventilation + S&I inputs (advanced controls). Defaults come from the corresponding td cells.
    // Controls.asp always shows ventilation, but its defaults may come from either td cells or the non-advanced field defaultValue.
    const ventValueDefault = (
      _finiteOrUndef(fd.ventilationValue_default) ??
      _finiteOrUndef(fd.txtVentilationValue_defaultValue) ??
      _finiteOrUndef(fd.txtVentilationValue_NotAdvanced_defaultValue)
    );
    _pushOverride(o, 'VentilationValue', inp.ventilationValue, ventValueDefault, inp.ventilationValueSource);

    const ventUnitsDefault = _normVentUnitsLabel(
      fd.cmbVentilationUnits_defaultValue ??
      fd.ventilationUnits_default
    );
    if (String(ventUnitsDefault ?? '').trim()) {
      _pushOverride(o, 'VentilationUnits', inp.ventilationUnitsLabel, ventUnitsDefault, undefined);
    }

    const sandIDefault = (
      _finiteOrUndef(fd.sandI_fraction_default) ??
      _finiteOrUndef(fd.txtSI_Fraction_defaultValue) ??
      _finiteOrUndef(fd.txtSI_Fraction_NotAdvanced_defaultValue) ??
      _finiteOrUndef(DEFAULTS.SandI_fraction)
    );
    _pushOverride(o, 'SandI_fraction', inp.sandI_fraction, sandIDefault, inp.sandISource);

    const nAffinityDefault = (fd.cmbN_Affinity_defaultValue !== undefined)
      ? fd.cmbN_Affinity_defaultValue
      : DEFAULTS.N_Affinity;
    _pushOverride(o, 'N_Affinity', inp.nAffinity, nAffinityDefault, inp.nAffinitySource);

    const stDefault = (fd.cmbST_Ratio_C_defaultValue !== undefined)
      ? fd.cmbST_Ratio_C_defaultValue
      : DEFAULTS.ST_Ratio_AtTest_C;
    _pushOverride(o, 'ST_Ratio_AtTest_C', inp.stAtTest_C, stDefault, inp.stSource_C);

    // Checkboxes / toggles: compare to the page defaultChecked.
    // Advanced controls default is indicated by the UI default cell (Hidden vs Shown).
    // The checkbox may be server-rendered with checked based on the current state, so defaultChecked is not reliable.
    const advDefaultFromUi = String(fd.tdAdvancedControls_defaultLabel || '').toLowerCase().includes('hidden') ? false :
      (String(fd.tdAdvancedControls_defaultLabel || '').toLowerCase().includes('shown') ? true : undefined);
    const advDefault = (typeof advDefaultFromUi === 'boolean')
      ? advDefaultFromUi
      : ((typeof fd.chkAdvancedControls_defaultChecked === 'boolean') ? fd.chkAdvancedControls_defaultChecked : false);
    _pushOverride(o, 'AdvancedControls', !!inp.advancedControlsChecked, advDefault, undefined);

    // Economizers are default ON in the UI; compare to defaultChecked.
    const econCDefault = (typeof fd.chkEconomizer_C_defaultChecked === 'boolean') ? fd.chkEconomizer_C_defaultChecked : true;
    _pushOverride(o, 'Economizer_Candidate', !!inp.econCandidate, econCDefault, undefined);

    const econSDefault = (typeof fd.chkEconomizer_S_defaultChecked === 'boolean') ? fd.chkEconomizer_S_defaultChecked : true;
    _pushOverride(o, 'Economizer_Standard', !!inp.econStandard, econSDefault, undefined);

    // Load line lock + values.
    const lockDefault = (typeof fd.chkLockLoadLine_defaultChecked === 'boolean') ? fd.chkLockLoadLine_defaultChecked : false;
    _pushOverride(o, 'LockLoadLine', !!inp.lockLoadLine, lockDefault, undefined);
    if (!!inp.lockLoadLine) {
      const slopeDefault = fd.txtSlope_defaultValue;
      const interceptDefault = fd.txtIntercept_defaultValue;
      _pushOverride(o, 'LockedSlope', inp.lockedSlope, slopeDefault, undefined);
      _pushOverride(o, 'LockedIntercept', inp.lockedIntercept, interceptDefault, undefined);
    }

    return o;
  }

  // ---- Economics / payback (match ASP Engine.asp) ----
  const ec = last.econ || {};
  const elecRate = Number(ec.electricityRate) || 0;
  const drRate = Number(ec.discountRate) || 0;
  const eqLife = Number(ec.equipmentLife) || 15;
  const dMonths = Number(ec.demandMonths) || 0;
  const dCostKW = Number(ec.demandCostPerKW) || 0;
  const isChartPW = !!ec.chartPW;

  // Economics use per-unit energy (matches ASP Engine.asp lines 3221-3222:
  // objSD_C.AnnualCost = (objBD_C_local.Energy_Annual_Total * ElectricityRate) + Maintenance + DemandCost
  // where Energy_Annual_Total is per-unit, not nUnits-scaled).
  const rawC = {
    total: Number(last.resC?.annualTotal ?? 0),
    peak: Number(last.resC?.peakDemand ?? 0),
  };
  const rawS = {
    total: Number(last.resS?.annualTotal ?? 0),
    peak: Number(last.resS?.peakDemand ?? 0),
  };

  const demandCost_C = rawC.peak * dCostKW * dMonths;
  const demandCost_S = rawS.peak * dCostKW * dMonths;
  const unitCost_C = Number(last.objSD_C?.UnitCost ?? 0);
  const unitCost_S = Number(last.objSD_S?.UnitCost ?? 0);
  const maint_C = Number(last.objSD_C?.Maintenance ?? 0);
  const maint_S = Number(last.objSD_S?.Maintenance ?? 0);
  const annualCost_C = (rawC.total * elecRate) + maint_C + demandCost_C;
  const annualCost_S = (rawS.total * elecRate) + maint_S + demandCost_S;
  const costAnnualSavings = annualCost_S - annualCost_C;
  const capCostSavings = 1000 * (unitCost_C - unitCost_S);

  const simplePayback = (costAnnualSavings !== 0) ? (capCostSavings / costAnnualSavings) : -1;

  // UPV: Uniform Present Value factor (match ASP)
  function _upv(life, dr) {
    if (!dr || dr === 0) return life;
    let a;
    try { a = Math.pow(1 + dr, life); } catch { a = 1; }
    if (!Number.isFinite(a) || a === 0) a = 1;
    return (a - 1) / (dr * a);
  }

  // NPV: Net Present Value (match ASP — depends on annualCost_C/S, unitCost_C/S)
  function _npv(dr, life) {
    const lccC = (annualCost_C * _upv(life, dr)) + (unitCost_C * 1000);
    const lccS = (annualCost_S * _upv(life, dr)) + (unitCost_S * 1000);
    return lccS - lccC;
  }

  // Payback Newton iteration (match ASP PayBack function)
  function _paybackNewton(initialGuess) {
    let prev = initialGuess;
    let est;
    let j;
    for (j = 0; j < 10; j++) {
      const errPrev = 0 - _npv(drRate, prev);
      const errDelta = 0 - _npv(drRate, prev + 0.0005);
      const slope = (errPrev - errDelta) / 0.0005;
      if (slope === 0 || !Number.isFinite(slope)) return null;
      est = prev + errPrev / slope;
      const errEst = 0 - _npv(drRate, est);
      if (Math.abs(errEst) < 0.01) break;
      prev = est;
    }
    if (est != null && Number.isFinite(est) && j <= 10 && est > 0) return est;
    return null;
  }

  let discountedPayback;
  if (simplePayback < 0) {
    discountedPayback = 0; // "Immediate"
  } else if (simplePayback > 100) {
    discountedPayback = -1;
  } else {
    const pb = _paybackNewton(simplePayback);
    discountedPayback = (pb !== null && pb > 0) ? pb : -1;
  }

  // NPV = LCC_Standard - LCC_Candidate (at the form's discount rate & equipment life)
  const upvVal = _upv(eqLife, drRate);
  const lccC = (unitCost_C * 1000) + (annualCost_C * upvVal);
  const lccS = (unitCost_S * 1000) + (annualCost_S * upvVal);
  const npvVal = lccS - lccC;

  // SIR = LCC_Annual_Savings / capitalCostSavings
  const lccAnnualSavings = costAnnualSavings * upvVal;
  const sirVal = (capCostSavings !== 0) ? (lccAnnualSavings / capCostSavings) : 0;

  // ROR: Newton iteration finding discount rate where NPV = 0
  // Initial guess = 1/simplePayback (= annualSavings/capitalCost)
  function _rorNewton(initialGuess) {
    let prev = initialGuess;
    let est;
    let j;
    for (j = 0; j < 10; j++) {
      const errPrev = 0 - _npv(prev, eqLife);
      const errDelta = 0 - _npv(prev + 0.0005, eqLife);
      const slope = (errPrev - errDelta) / 0.0005;
      if (slope === 0 || !Number.isFinite(slope)) return null;
      est = prev + errPrev / slope;
      const errEst = 0 - _npv(est, eqLife);
      if (Math.abs(errEst) < 0.01) break;
      prev = est;
    }
    if (est != null && Number.isFinite(est) && j < 10 && est > 0) return 100 * est;
    return null;
  }

  let rorVal = null;
  if (capCostSavings !== 0) {
    rorVal = _rorNewton(costAnnualSavings / capCostSavings);
  }

  const unitsFactor = (Number.isFinite(last.nUnits) && last.nUnits > 0) ? last.nUnits : 1;

  // Fields that should be scaled by nUnits (energy quantities).
  // Condition/performance fields (owb, ohr, tcf, stRatio, etc.) must NOT be scaled.
  const _scaleFields = new Set(['eFan', 'eCond', 'eAux']);

  function mapToArray(m, scale) {
    const arr = Array.from((m || new Map()).values()).sort((a, b) => Number(a.odb) - Number(b.odb));
    if (scale && scale !== 1) {
      return arr.map(bin => {
        const out = {};
        for (const k of Object.keys(bin)) {
          if (_scaleFields.has(k) && typeof bin[k] === 'number') {
            out[k] = bin[k] * scale;
          } else {
            out[k] = bin[k];
          }
        }
        return out;
      });
    }
    return arr;
  }
  return {
    meta: {
      export: 'js',
      engineVersion: ENGINE_VERSION,
      nUnits: Number(inp.nUnits || 1),
      overrides: _computeOverrides(),
      loadLine: last.resC?.loadLine || null,
      inputs: {
        totalCap_kBtuh: Number(inp.totalCap || 0),
        cfm: Number(inp.cfm || 0),
        advancedControlsChecked: !!inp.advancedControlsChecked,
        ventilationValue: inp.ventilationValue,
        ventilationUnits: inp.ventilationUnitsLabel,
        nAffinity: inp.nAffinity,
        trackOhr: inp.trackOhr,
        irhPct: inp.irhPct,
        lockLoadLine: !!inp.lockLoadLine,
        demandMonths: dMonths,
        demandCostPerKW: dCostKW,
        bmSlope: inp.raw?.txtBM_Slope_hidden,
        bmIntercept: inp.raw?.txtBM_Intercept_hidden,
        bmVentFrac: inp.raw?.txtBM_VentSlopeFraction_hidden,
        candidate: {
          fanControls: String(last.objSD_C?.FanControls || ''),
          bfn_kw: Number(last.objSD_C?.BFn_kw || 0),
          cond_kw: Number(last.objSD_C?.Cond_kw || 0),
          aux_kw: Number(last.objSD_C?.Aux_kw || 0),
          eer: Number(last.objSD_C?.EER || 0),
          condFanPercent: Number(last.objSD_C?.CondFanPercent || 0),
          plDegrFactor: Number(last.objSD_C?.PLDegrFactor || 0),
          nStages: Number(last.objSD_C?.N_Stages || 1),
          stRatioAtTest: Number(last.objSD_C?.ST_Ratio_AtTest || 0),
          specificRtu: String(last.objSD_C?.Specific_RTU || 'None'),
          spreadsheet: !!(last.objSD_C?.Spreadsheet),
        },
        standard: {
          fanControls: String(last.objSD_S?.FanControls || ''),
          bfn_kw: Number(last.objSD_S?.BFn_kw || 0),
          cond_kw: Number(last.objSD_S?.Cond_kw || 0),
          aux_kw: Number(last.objSD_S?.Aux_kw || 0),
          eer: Number(last.objSD_S?.EER || 0),
          condFanPercent: Number(last.objSD_S?.CondFanPercent || 0),
          plDegrFactor: Number(last.objSD_S?.PLDegrFactor || 0),
          nStages: Number(last.objSD_S?.N_Stages || 1),
          stRatioAtTest: Number(last.objSD_S?.ST_Ratio_AtTest || 0),
          specificRtu: String(last.objSD_S?.Specific_RTU || 'None'),
          spreadsheet: !!(last.objSD_S?.Spreadsheet),
        },
      },
      debug: {
        candidate_aux_hours_occ: Number(last.resC_scaled?.auxHoursOcc ?? last.resC.auxHoursOcc ?? 0),
        candidate_aux_hours_unocc: Number(last.resC_scaled?.auxHoursUnocc ?? last.resC.auxHoursUnocc ?? 0),
        standard_aux_hours_occ: Number(last.resS_scaled?.auxHoursOcc ?? last.resS.auxHoursOcc ?? 0),
        standard_aux_hours_unocc: Number(last.resS_scaled?.auxHoursUnocc ?? last.resS.auxHoursUnocc ?? 0),
      },
      models: inp.spreadsheet || null,
    },
    designConditions: (() => {
      const _odb = last.ll?.design?.odb ?? 0;
      const _owb = last.ll?.design?.owb ?? 0;
      const _ohr = last.ll?.design?.ohr ?? 0;
      const _p   = last.ll?.design?.pressure ?? 0;
      const _idb = last.ll?.design?.idb ?? 0;
      const _ihr = last.ll?.design?.insideHR ?? 0;
      const _edb = last.ll?.design?.entering?.edb ?? 0;
      const _ewb = last.ll?.design?.entering?.ewb ?? 0;
      const _ehr = last.ll?.design?.entering?.ehr ?? 0;
      // ORH = Prh_hr(ODB, OHR, P) * 100
      let _orh = 0; try { _orh = Prh_hr(_odb, _ohr, _p) * 100; } catch { }
      // ERH = Prh_hr(EDB, EHR, P) * 100
      let _erh = 0; try { _erh = Prh_hr(_edb, _ehr, _p) * 100; } catch { }
      // IWB = Pwb_hr(IDB, IHR, P)
      let _iwb = 0; try { _iwb = Pwb_hr(_idb, _ihr, _p); } catch { }
      // TCF at design = Tot_Capacity_Correction(objSD_C, 'generic', totalCap, stage, ODB, EWB)
      let _tcf = 0;
      try {
        const _stage = new StageState(); _stage.CapacityFraction = 1.0;
        _tcf = Tot_Capacity_Correction(last.objSD_C, 'generic', Number(inp.totalCap || 0), _stage, _odb, _ewb);
      } catch { }
      // BPF from runBinCalcs (stored in last)
      const _bpfC = last.bpfC ?? 0;
      const _bpfS = last.bpfS ?? 0;
      return {
        outdoor: {
          odb: _odb, owb: _owb, ohr: _ohr, orh: _orh,
          elevation: last.ll?.design?.elevation ?? 0, pressure: _p,
        },
        indoor: {
          idb: _idb, iwb: _iwb, insideHR: _ihr,
          insideRH: last.ll?.design?.insideRH ?? 0,
        },
        entering: {
          edb: _edb, ewb: _ewb, ehr: _ehr, erh: _erh,
          tcf: _tcf,
          stRatio: last.ll?.design?.entering?.stRatio ?? 0,
        },
        capacity: {
          sensibleAtTest: last.ll?.debug?.capacity?.capacitySensibleAtTest ?? 0,
          sensibleAtDesign: last.ll?.debug?.capacity?.sensibleCapacityDesign ?? 0,
          cfm: Number(inp.cfm || 0),
          ventCFM: last.ll?.debug?.ventilation?.ventCFM ?? 0,
        },
        loads: {
          sensVentLoadDesign: last.ll?.debug?.ventilation?.sensVentLoadDesign ?? 0,
          sensNonVentLoadDesign: last.ll?.debug?.loads?.sensibleNonVentLoadAtDesign ?? 0,
          sandIfrac: last.ll?.debug?.inputs?.sandIfrac ?? 0,
        },
        loadLine: {
          slope: last.ll?.slope ?? 0,
          intercept: last.ll?.intercept ?? 0,
          locked: !!(last.ll?.debug?.lock?.applied),
        },
        bpf: {
          candidate: _bpfC,
          standard: _bpfS,
        },
      };
    })(),
    annual: {
      candidate: {
        condenser: Math.round(Number(last.resC_scaled?.annualCondenser ?? last.resC.annualCondenser ?? 0)),
        efan: Math.round(Number(last.resC_scaled?.annualEFan ?? last.resC.annualEFan ?? 0)),
        aux: Math.round(Number(last.resC_scaled?.annualAux ?? last.resC.annualAux ?? 0)),
        total: Math.round(Number(last.resC_scaled?.annualTotal ?? last.resC.annualTotal ?? 0)),
      },
      standard: {
        condenser: Math.round(Number(last.resS_scaled?.annualCondenser ?? last.resS.annualCondenser ?? 0)),
        efan: Math.round(Number(last.resS_scaled?.annualEFan ?? last.resS.annualEFan ?? 0)),
        aux: Math.round(Number(last.resS_scaled?.annualAux ?? last.resS.annualAux ?? 0)),
        total: Math.round(Number(last.resS_scaled?.annualTotal ?? last.resS.annualTotal ?? 0)),
      },
    },
    annual_raw: {
      candidate: {
        condenser: Number(last.resC_scaled?.annualCondenser ?? last.resC.annualCondenser ?? 0),
        efan: Number(last.resC_scaled?.annualEFan ?? last.resC.annualEFan ?? 0),
        aux: Number(last.resC_scaled?.annualAux ?? last.resC.annualAux ?? 0),
        total: Number(last.resC_scaled?.annualTotal ?? last.resC.annualTotal ?? 0),
      },
      standard: {
        condenser: Number(last.resS_scaled?.annualCondenser ?? last.resS.annualCondenser ?? 0),
        efan: Number(last.resS_scaled?.annualEFan ?? last.resS.annualEFan ?? 0),
        aux: Number(last.resS_scaled?.annualAux ?? last.resS.annualAux ?? 0),
        total: Number(last.resS_scaled?.annualTotal ?? last.resS.annualTotal ?? 0),
      },
    },
    economics: {
      electricityRate: elecRate,
      discountRate: drRate,
      equipmentLife: eqLife,
      chartPW: isChartPW,
      demandMonths: dMonths,
      demandCostPerKW: dCostKW,
      candidate: {
        unitCost: unitCost_C,
        maintenance: maint_C,
        demandCost: demandCost_C,
        annualCost: annualCost_C,
      },
      standard: {
        unitCost: unitCost_S,
        maintenance: maint_S,
        demandCost: demandCost_S,
        annualCost: annualCost_S,
      },
      savings: {
        energy_kWh: rawS.total - rawC.total,
        annualCost: costAnnualSavings,
        capitalCost: capCostSavings,
      },
      simplePayback: simplePayback,
      discountedPayback: discountedPayback,
      npv: npvVal,
      ror: rorVal,
      sir: sirVal,
    },
    bins: {
      candidate: {
        occ: mapToArray(last.resC.binsOcc, unitsFactor),
        unocc: mapToArray(last.resC.binsUnocc, unitsFactor),
        econOcc: mapToArray(last.resC.binsEconOcc, unitsFactor),
        econUnocc: mapToArray(last.resC.binsEconUnocc, unitsFactor),
      },
      standard: {
        occ: mapToArray(last.resS.binsOcc, unitsFactor),
        unocc: mapToArray(last.resS.binsUnocc, unitsFactor),
        econOcc: mapToArray(last.resS.binsEconOcc, unitsFactor),
        econUnocc: mapToArray(last.resS.binsEconUnocc, unitsFactor),
      },
    },
  };
}

// Compute the non-ventilation sensible load line (slope and intercept)
// from design conditions, S&I fraction, ventilation, and capacity.
// Optionally applies a locked load line from the form.  Returns
// { ok, slope, intercept, design, debug, warnings } or { ok:false, errorMessage }.
export function computeLoadLine({ formValues, candidateSD, stageState, datasets }) {
  try {
    if (!formValues) return { ok: false, errorMessage: 'Missing formValues' };
    if (!candidateSD) return { ok: false, errorMessage: 'Missing candidateSD' };
    if (!stageState) return { ok: false, errorMessage: 'Missing stageState' };
    if (!datasets?.stations) return { ok: false, errorMessage: 'Missing datasets.stations' };
    if (!datasets?.weather) return { ok: false, errorMessage: 'Missing datasets.weather' };

    const idb = _num(formValues.IDB);
    if (idb === null) return { ok: false, errorMessage: 'Invalid IDB (setpoint)' };

    const state = formValues.State;
    const city = formValues.CityName2;
    const schedule = formValues.Schedule;
    if (!state || !city || !schedule) {
      return { ok: false, errorMessage: 'Missing State/City/Schedule' };
    }

    const weatherCity = getWeatherRecords(datasets.weather, state, city, schedule);
    const dc = getDesignConditions(datasets.stations, weatherCity, city, idb, false);
    if (!Number.isFinite(dc?.ODB_Design) || !Number.isFinite(dc?.OWB_Design) || !Number.isFinite(dc?.Pressure_Design)) {
      return { ok: false, errorMessage: dc?.warning || 'Unable to determine design conditions' };
    }

    const odbDesign = dc.ODB_Design;
    const owbDesign = dc.OWB_Design;
    const pressureDesign = dc.Pressure_Design;

    const ohrDesign = Phr_wb(odbDesign, owbDesign, pressureDesign);

    const insideRH = IRH_Track_OR_Set(
      formValues.TrackOHR,
      ohrDesign,
      idb,
      _num(formValues.IRH) ?? 0,
      pressureDesign
    );

    const insideHR = Phr_rh(idb, insideRH, pressureDesign);
    const insideH = Ph_hr(idb, insideHR);

    const totalCap = _num(formValues.TotalCap);
    const cfm = _num(formValues.CFM);
    if (totalCap === null || cfm === null) {
      return { ok: false, errorMessage: 'Invalid TotalCap/CFM' };
    }

    const capacitySensibleAtTest = NetSenCap_Stage_Adjusted_KBtuh(
      candidateSD,
      totalCap,
      cfm,
      stageState,
      1,
      95,
      67,
      80,
      dblStandardPressure
    );

    let ventCFM;
    if (formValues.VentilationUnits === 'CFM') {
      ventCFM = _num(formValues.VentilationValue);
    } else {
      const frac = _num(formValues.VentilationValue);
      ventCFM = frac === null ? null : cfm * (frac / 100);
    }

    if (ventCFM === null) return { ok: false, errorMessage: 'Invalid ventilation value' };

    const sensVentLoadDesign = SensVentLoad(ventCFM, ohrDesign, odbDesign, idb, pressureDesign);

    const mix = Mixer2(
      candidateSD,
      idb,
      insideHR,
      odbDesign,
      ohrDesign,
      pressureDesign,
      ventCFM,
      cfm
    );

    const ehrEntering = mix?.EHR;
    const edbEntering = mix?.EDB;
    const ewbEntering = mix?.EWB;

    if (![ehrEntering, edbEntering, ewbEntering].every(Number.isFinite)) {
      return { ok: false, errorMessage: 'Unable to compute mixed-air entering conditions' };
    }

    let stRatioAtEntering;
    if (candidateSD?.ST_Ratio_Model?.IsSolved) {
      stRatioAtEntering = candidateSD.Predict_ST_Ratio(odbDesign, ewbEntering, edbEntering);
    } else {
      const maker = formValues.Manufacturer || 'generic';
      const stRes = ST_Ratio_engine(maker, totalCap, cfm, stageState, candidateSD, 1, odbDesign, ewbEntering, edbEntering, pressureDesign);
      stRatioAtEntering = stRes?.stRatio;
    }

    const sensibleCapacityDesign = NetSenCap_Stage_Adjusted_KBtuh(
      candidateSD,
      totalCap,
      cfm,
      stageState,
      1,
      odbDesign,
      ewbEntering,
      edbEntering,
      pressureDesign
    );

    const oversizeLRF = _num(formValues.Oversizing_LoadReductionFactor) ?? 1;
    const sensibleLoadAtDesign = sensibleCapacityDesign / oversizeLRF;

    let sensibleNonVentLoadAtDesign = sensibleLoadAtDesign - sensVentLoadDesign;
    if (sensibleNonVentLoadAtDesign < 0) sensibleNonVentLoadAtDesign = 0;

    const sandIfrac = _num(formValues.SandI_fraction);
    if (sandIfrac === null) return { ok: false, errorMessage: 'Invalid SandI_fraction' };
    const sandI = sensibleLoadAtDesign * sandIfrac;

    if (odbDesign === idb) {
      return {
        ok: false,
        errorMessage: `Outside design temperature (${odbDesign}) is equal to the set point (${idb}).`,
      };
    }

    let slope = (sensibleNonVentLoadAtDesign - sandI) / (odbDesign - idb);
    let intercept = sandI;

    const debug = {
      inputs: {
        manufacturer: formValues.Manufacturer,
        totalCap: totalCap,
        cfm: cfm,
        idb: idb,
        state,
        city,
        schedule,
        ventilationUnits: formValues.VentilationUnits,
        ventilationValue: _num(formValues.VentilationValue),
        trackOhr: formValues.TrackOHR,
        irh: _num(formValues.IRH),
        oversizeLRF,
        sandIfrac,
      },
      design: {
        odbDesign,
        owbDesign,
        pressureDesign,
        elevationDesign: dc.Elevation_Design,
        ohrDesign,
        insideRH,
        insideHR,
        insideH,
      },
      weather: {
        weatherRowCount: Array.isArray(weatherCity) ? weatherCity.length : null,
      },
      ventilation: {
        ventCFM,
        sensVentLoadDesign,
      },
      entering: {
        edbEntering,
        ewbEntering,
        ehrEntering,
        stRatioAtEntering,
      },
      capacity: {
        capacitySensibleAtTest,
        sensibleCapacityDesign,
      },
      loads: {
        sensibleLoadAtDesign,
        sensibleNonVentLoadAtDesign,
        sandI,
      },
      result: {
        slopeComputed: slope,
        interceptComputed: intercept,
      },
      lock: {
        lockLoadLine: formValues.LockLoadLine === 'on',
        lockedSlope: _num(formValues.Slope),
        lockedIntercept: _num(formValues.Intercept),
      },
    };

    // Parity: if the server already provided baseline load-line parameters, prefer those.
    // (Avoids tiny client-side differences in design-condition picks and psychro calculations.)
    const bmSlope = _num(formValues.BM_Slope_hidden);
    const bmIntercept = _num(formValues.BM_Intercept_hidden);
    if (formValues.LockLoadLine !== 'on' && bmSlope !== null && bmIntercept !== null) {
      slope = bmSlope;
      intercept = bmIntercept;
      debug.result.slopeUsed = slope;
      debug.result.interceptUsed = intercept;
      debug.result.source = 'BM_hidden';
    }

    if (formValues.LockLoadLine === 'on') {
      const lockedSlope = _num(formValues.Slope);
      const lockedIntercept = _num(formValues.Intercept);
      if (lockedSlope !== null && lockedIntercept !== null) {
        slope = lockedSlope;
        intercept = lockedIntercept;
        debug.result.slopeUsed = slope;
        debug.result.interceptUsed = intercept;
        debug.result.source = 'locked';
        debug.lock.applied = true;
      } else {
        debug.lock.applied = false;
      }
    } else {
      debug.lock.applied = false;
    }

    if (slope < 0) {
      if ((odbDesign - idb) < 0) {
        return {
          ok: false,
          errorMessage: `Outside design temperature (${odbDesign}) is lower than the set point (${idb}).`,
        };
      }

      return {
        ok: false,
        errorMessage:
          'For the specified ventilation rate, the non-ventilation load at design is less than the internal load (negative non-ventilation slope). Reduce ventilation or lock the non-ventilation load line before increasing ventilation.',
      };
    }

    return {
      ok: true,
      slope,
      intercept,
      debug,
      design: {
        odb: odbDesign,
        owb: owbDesign,
        pressure: pressureDesign,
        elevation: dc.Elevation_Design,
        ohr: ohrDesign,
        idb,
        insideRH,
        insideHR,
        insideH,
        entering: { edb: edbEntering, ewb: ewbEntering, ehr: ehrEntering, stRatio: stRatioAtEntering },
        capacitySensibleAtTest,
      },
      warnings: dc.warning ? [dc.warning] : [],
    };
  } catch (err) {
    return { ok: false, errorMessage: err?.message || String(err) };
  }
}

// Internal engine entry point: parse all form inputs, load JSON data,
// compute the load line, then loop over temperature bins for both
// Candidate and Standard units.  Stores results in module-level state
// (_lastRunPhase1, _lastRunInputs) for consumption by exportBinCalcsJson.
export async function runBinCalcs(form, opts = {}) {
  try {
    if (!form) {
      return '<div class="graph"><h2>Client-side engine error</h2><pre>Missing form</pre></div>';
    }

    const totalCap = _asNumber(form, 'txtTotalCap');
    if (!Number.isFinite(totalCap) || totalCap <= 0) {
      return (
        '<div class="graph"><h2>Client-side engine error</h2><pre>' +
        _escapeHtml("Invalid 'txtTotalCap' (expected kBtuh > 0). Current value: " + String(_val(form, 'txtTotalCap'))) +
        '</pre></div>'
      );
    }

    const cfm = (totalCap / 12) * DEFAULTS.dblCFMperTon;

    const bfnPick_C = _pickNumber([
      { source: 'txtBFn_kw_C', value: _val(form, 'txtBFn_kw_C') },
      { source: 'txtBFn_kw_C_NotAdvanced', value: _val(form, 'txtBFn_kw_C_NotAdvanced') },
      { source: 'txtBFn_kw_C_hidden', value: _val(form, 'txtBFn_kw_C_hidden') },
      { source: 'tdBFn_kw_C_default', value: _textById(form, 'tdBFn_kw_C_default') },
      { source: 'setFanPowerValuesAndDefaults()', value: _fanPowerDefaultFromPage(totalCap) },
      { source: 'defaults_advancedControls.asp', value: _fanPowerDefaultFromAspDefaults(totalCap) },
    ]);

    const bfnPick_S = _pickNumber([
      { source: 'txtBFn_kw_S', value: _val(form, 'txtBFn_kw_S') },
      { source: 'txtBFn_kw_S_NotAdvanced', value: _val(form, 'txtBFn_kw_S_NotAdvanced') },
      { source: 'txtBFn_kw_S_hidden', value: _val(form, 'txtBFn_kw_S_hidden') },
      { source: 'tdBFn_kw_S_default', value: _textById(form, 'tdBFn_kw_S_default') },
      { source: 'setFanPowerValuesAndDefaults()', value: _fanPowerDefaultFromPage(totalCap) },
      { source: 'defaults_advancedControls.asp', value: _fanPowerDefaultFromAspDefaults(totalCap) },
    ]);

    const bfnKwForA0_C = bfnPick_C.value;
    const bfnKwForA0_S = bfnPick_S.value;

    if (!Number.isFinite(bfnKwForA0_C) || !Number.isFinite(bfnKwForA0_S)) {
      return (
        '<div class="graph"><h2>Client-side engine error</h2><pre>' +
        _escapeHtml(
          "Invalid blower-fan kW (candidate). Could not find 'txtBFn_kw_C' and no default cell value was available."
        ) +
        '</pre></div>'
      );
    }

    const stPick_C = _pickNumber([
      { source: 'cmbST_Ratio_C', value: _val(form, 'cmbST_Ratio_C') },
      { source: 'txtST_Ratio_C_hidden', value: _val(form, 'txtST_Ratio_C_hidden') },
      { source: 'defaults_advancedControls.asp', value: DEFAULTS.ST_Ratio_AtTest_C },
    ]);

    const stPick_S = _pickNumber([
      { source: 'cmbST_Ratio_S', value: _val(form, 'cmbST_Ratio_S') },
      { source: 'txtST_Ratio_S_hidden', value: _val(form, 'txtST_Ratio_S_hidden') },
      { source: 'defaults_advancedControls.asp', value: DEFAULTS.ST_Ratio_AtTest_C },
    ]);

    const stAtTest_C = stPick_C.value;
    const stAtTest_S = stPick_S.value;
    if (!Number.isFinite(stAtTest_C) || stAtTest_C <= 0 || stAtTest_C >= 1) {
      return (
        '<div class="graph"><h2>Client-side engine error</h2><pre>' +
        _escapeHtml("Invalid 'cmbST_Ratio_C' (expected 0-1). Current value: " + String(_val(form, 'cmbST_Ratio_C'))) +
        '</pre></div>'
      );
    }
    if (!Number.isFinite(stAtTest_S) || stAtTest_S <= 0 || stAtTest_S >= 1) {
      return (
        '<div class="graph"><h2>Client-side engine error</h2><pre>' +
        _escapeHtml("Invalid 'cmbST_Ratio_S' (expected 0-1). Current value: " + String(_val(form, 'cmbST_Ratio_S'))) +
        '</pre></div>'
      );
    }

    const ehrAtTest = Phr_wb(80, 67, 29.921);
    const grossTotCapBtuhForA0_C = (totalCap + bfnKwForA0_C * dblKWtoKBTUH) * 1000;
    const bpfRes_C = BPF_ADP_SC(grossTotCapBtuhForA0_C, stAtTest_C, cfm, 80, ehrAtTest, 29.921);
    if (bpfRes_C?.errorMessage) {
      return '<div class="graph"><h2>Client-side engine error</h2><pre>' + _escapeHtml('Candidate Unit BPF: ' + bpfRes_C.errorMessage) + '</pre></div>';
    }
    const a0Bpf_C = A0_FromBPF(bpfRes_C.bpf, cfm);

    const grossTotCapBtuhForA0_S = (totalCap + bfnKwForA0_S * dblKWtoKBTUH) * 1000;
    const bpfRes_S = BPF_ADP_SC(grossTotCapBtuhForA0_S, stAtTest_S, cfm, 80, ehrAtTest, 29.921);
    if (bpfRes_S?.errorMessage) {
      return '<div class="graph"><h2>Client-side engine error</h2><pre>' + _escapeHtml('Standard Unit BPF: ' + bpfRes_S.errorMessage) + '</pre></div>';
    }
    const a0Bpf_S = A0_FromBPF(bpfRes_S.bpf, cfm);

    const nPick = _pickNumber([
      { source: 'cmbN_Affinity', value: _val(form, 'cmbN_Affinity') },
      { source: 'defaults_advancedControls.asp', value: DEFAULTS.N_Affinity },
    ]);
    const nAffinity = nPick.value;
    if (!Number.isFinite(nAffinity) || nAffinity <= 0) {
      return (
        '<div class="graph"><h2>Client-side engine error</h2><pre>' +
        _escapeHtml("Invalid 'cmbN_Affinity'. Current value: " + String(_val(form, 'cmbN_Affinity'))) +
        '</pre></div>'
      );
    }

    const doe2Checked = _checkedOrDefault(form, 'chkDOE2_Curves', DEFAULTS.DOE2_Curves === 'on');
    const doe2Source = _el(form, 'chkDOE2_Curves') ? 'chkDOE2_Curves' : 'defaults_advancedControls.asp';

    const specificRtu_C = _valOrDefault(form, 'cmbSpecific_RTU_C', DEFAULTS.Specific_RTU_C) || DEFAULTS.Specific_RTU_C;
    const specificRtuSource_C = _el(form, 'cmbSpecific_RTU_C') ? 'cmbSpecific_RTU_C' : 'defaults_advancedControls.asp';

    // Legacy behavior: "Specific Candidate Unit" applies only to Candidate; Standard remains None.
    const specificRtu_S = 'None';
    const specificRtuSource_S = 'legacy';

    const stageState = new StageState();
    stageState.ResetToFullLoad();

    const advancedControlsChecked = _checked(form, 'chkAdvancedControls');

    const commonSysOpts = {
      totalCap,
      cfm,
      nAffinity,
      doe2Curves: doe2Checked ? 'DOE2' : 'Carrier',
    };

    const objSD_C = _systemFromForm(form, 'C', {
      ...commonSysOpts,
      specificRtu: specificRtu_C,
      stAtTest: stAtTest_C,
      a0Bpf: a0Bpf_C,
      advancedControlsChecked,
    });
    const objSD_S = _systemFromForm(form, 'S', {
      ...commonSysOpts,
      specificRtu: specificRtu_S,
      stAtTest: stAtTest_S,
      a0Bpf: a0Bpf_S,
      advancedControlsChecked,
    });

    // Spreadsheet modeling (if provided).
    if (objSD_C.Spreadsheet) {
      objSD_C.ParseAndModel('pANDm');
      const ssAux = parseFloat(String(objSD_C?.SSD?.AuxilaryPower ?? ''));
      if (Number.isFinite(ssAux)) objSD_C.Aux_kw = ssAux;
    }
    if (objSD_S.Spreadsheet) {
      objSD_S.ParseAndModel('pANDm');
      const ssAux = parseFloat(String(objSD_S?.SSD?.AuxilaryPower ?? ''));
      if (Number.isFinite(ssAux)) objSD_S.Aux_kw = ssAux;
    }

    const state = _val(form, 'cmbState');
    const city = _val(form, 'cmbCityName2');
    const schedule = _val(form, 'cmbSchedule');

    const nUnits = _asNumberOr(form, 'txtNUnits', 1);

    const ventUnitsLabel = _valOrDefault(form, 'cmbVentilationUnits', DEFAULTS.VentilationUnits);
    const ventUnits = ventUnitsLabel && ventUnitsLabel.toUpperCase().includes('CFM') ? 'CFM' : '%';
    const ventValuePick = advancedControlsChecked
      ? _pickNumber([
          { source: 'txtVentilationValue', value: _val(form, 'txtVentilationValue') },
          { source: 'txtVentilationValue_NotAdvanced', value: _val(form, 'txtVentilationValue_NotAdvanced') },
          { source: 'defaults_advancedControls.asp', value: DEFAULTS.VentilationValue },
        ])
      : _pickNumber([
          { source: 'txtVentilationValue_NotAdvanced', value: _val(form, 'txtVentilationValue_NotAdvanced') },
          { source: 'defaults_advancedControls.asp', value: DEFAULTS.VentilationValue },
        ]);
    const ventilationValueRaw = ventValuePick.value;

    const trackOhr = _checkedOrDefault(form, 'chkTrackOHR', DEFAULTS.TrackOHR === 'on') ? 'on' : '';
    const irhPick = _pickNumber([
      { source: 'cmbIRH', value: _val(form, 'cmbIRH') },
      { source: 'txtIRH_hidden', value: _val(form, 'txtIRH_hidden') },
      { source: 'defaults_advancedControls.asp', value: DEFAULTS.IRH_pct },
    ]);
    const irhPct = irhPick.value;
    const irh = irhPct / 100;

    const oversizePick = _pickNumber([
      { source: 'cmbOversizePercent', value: _val(form, 'cmbOversizePercent') },
      { source: 'default', value: 0 },
    ]);
    const oversizePct = oversizePick.value ?? 0;
    const oversizingLRF = oversizePct / 100 + 1;

    const sandIPick = advancedControlsChecked
      ? _pickNumber([
          { source: 'txtSI_Fraction', value: _val(form, 'txtSI_Fraction') },
          { source: 'txtSI_Fraction_NotAdvanced', value: _val(form, 'txtSI_Fraction_NotAdvanced') },
          { source: 'defaults_advancedControls.asp', value: DEFAULTS.SandI_fraction },
        ])
      : _pickNumber([
          { source: 'txtSI_Fraction_NotAdvanced', value: _val(form, 'txtSI_Fraction_NotAdvanced') },
          { source: 'defaults_advancedControls.asp', value: DEFAULTS.SandI_fraction },
        ]);
    let sandIFraction = sandIPick.value;
    if (Number.isFinite(sandIFraction)) sandIFraction = Math.round(sandIFraction * 1000) / 1000;

    const idb = _asNumber(form, 'cmbIDB');
    if (!Number.isFinite(idb)) {
      return (
        '<div class="graph"><h2>Client-side engine error</h2><pre>' +
        _escapeHtml("Invalid 'cmbIDB'. Current value: " + String(_val(form, 'cmbIDB'))) +
        '</pre></div>'
      );
    }

    const idbSetBackRaw = _val(form, 'cmbIDB_SetBack');
    const idbSetBack = (String(idbSetBackRaw || '').trim().toLowerCase() === 'cond. off') ? 0 : Number(idbSetBackRaw);
    const idbUnoccupied = idb + (Number.isFinite(idbSetBack) ? idbSetBack : 0);

    const lockLoadLine = _checked(form, 'chkLockLoadLine') ? 'on' : '';
    const lockedSlope = _val(form, 'txtSlope');
    const lockedIntercept = _val(form, 'txtIntercept');

    // Economics inputs for payback parity.
    const electricityRate = _asNumberOr(form, 'txtElectricityRate', 0.08);
    const discountRatePct = _asNumberOr(form, 'txtDiscountRate_hidden', 5);
    const discountRate = discountRatePct / 100;
    const equipmentLife = _asNumberOr(form, 'cmbEquipmentLife', 15);
    const chartPW = _checked(form, 'chkChartPW');
    const demandMonths = _asNumberOr(form, 'cmbDemandMonths', 4);
    const demandCostPerKW = _asNumberOr(form, 'txtDemandCostPerKW', 0);

    const formValues = {
      Manufacturer: _val(form, 'txtManufacturer') || 'generic',
      TotalCap: totalCap,
      CFM: cfm,
      IDB: idb,
      State: state,
      CityName2: city,
      Schedule: schedule,
      VentilationUnits: ventUnits === 'CFM' ? 'CFM' : '% of Fan Cap.',
      VentilationValue: _num(ventilationValueRaw),
      TrackOHR: trackOhr,
      IRH: irh,
      Oversizing_LoadReductionFactor: oversizingLRF,
      SandI_fraction: sandIFraction,
      LockLoadLine: lockLoadLine,
      Slope: lockLoadLine === 'on' ? lockedSlope : '',
      Intercept: lockLoadLine === 'on' ? lockedIntercept : '',
      // Parity: these hidden inputs are sometimes not part of form.elements; fall back to document.getElementById.
      BM_Slope_hidden: _val(form, 'txtBM_Slope_hidden') || _valById(form, 'txtBM_Slope_hidden'),
      BM_Intercept_hidden: _val(form, 'txtBM_Intercept_hidden') || _valById(form, 'txtBM_Intercept_hidden'),
    };

    const cache = _dataCache;
    const _moduleBase = new URL('.', import.meta.url).href;
    async function loadJsonCached(url) {
      const resolvedUrl = new URL(url, _moduleBase).href;
      if (cache[resolvedUrl]) return cache[resolvedUrl];
      const resp = await fetch(resolvedUrl);
      if (!resp.ok) throw new Error('Failed to load ' + url + ': ' + resp.status);
      const data = await resp.json();
      cache[resolvedUrl] = data;
      return data;
    }

    const [stations, weather] = await Promise.all([
      loadJsonCached('../data/stations.json'),
      loadJsonCached('../data/Tbins_new.json'),
    ]);

    const stationsNorm = (stations || []).map((s) => ({
      City: s?.City ?? s?.city,
      state: s?.state ?? s?.State,
      Temp_DB: s?.Temp_DB ?? s?.temp_db,
      Temp_WB: s?.Temp_WB ?? s?.temp_wb,
      elevation: s?.elevation ?? s?.Elevation,
    }));

    const weatherNorm = (weather || []).map((r) => ({
      state: r?.state ?? r?.State,
      city: r?.city ?? r?.City,
      schedule: r?.schedule ?? r?.Schedule,
      Temp_Outdoor_DB: r?.Temp_Outdoor_DB,
      Temp_Coinc_WB: r?.Temp_Coinc_WB,
      Hours_Cooling: r?.Hours_Cooling,
    }));

    const ll = computeLoadLine({
      formValues,
      candidateSD: objSD_C,
      stageState,
      datasets: { stations: stationsNorm, weather: weatherNorm },
    });

    const aspLoadLine = opts?.aspLoadLine;
    if (
      ll?.ok !== false &&
      formValues.LockLoadLine !== 'on' &&
      aspLoadLine &&
      Number.isFinite(Number(aspLoadLine.slope)) &&
      Number.isFinite(Number(aspLoadLine.intercept))
    ) {
      ll.slope = Number(aspLoadLine.slope);
      ll.intercept = Number(aspLoadLine.intercept);
      if (ll.debug && ll.debug.result) {
        ll.debug.result.slopeUsed = ll.slope;
        ll.debug.result.interceptUsed = ll.intercept;
        ll.debug.result.source = 'asp_meta';
      }
    }

    // Expose load line early so recalcVentilation can use it even if BPF/ADP
    // checks fail later (e.g. S/T = 0.80).  ASP's Establish_SIV does not
    // depend on BPF succeeding — it only needs the load line.
    _lastLoadLine = ll;

    const buildingType =
      _val(form, 'txtBuildingType_hidden') ||
      _val(form, 'cmbBuildingType') ||
      _val(form, 'txtBuildingType_previousvalue_hidden') ||
      '';

    function _defaultValue(name) {
      try {
        const el = _el(form, name);
        if (!el) return undefined;
        // For <select>, defaultValue isn't reliable; use the option marked defaultSelected.
        if (String(el?.tagName || '').toUpperCase() === 'SELECT') {
          const t = _selectDefaultText(el);
          return t ? t : undefined;
        }
        const dv = el.defaultValue;
        return dv === null || dv === undefined ? undefined : String(dv);
      } catch {
        return undefined;
      }
    }

    function _defaultChecked(name) {
      try {
        const el = _el(form, name);
        if (!el) return undefined;
        const dc = el.defaultChecked;
        return typeof dc === 'boolean' ? dc : undefined;
      } catch {
        return undefined;
      }
    }

    const runInputs = {
      totalCap,
      cfm,
      idb,
      state,
      city,
      schedule,
      buildingType,
      advancedControlsChecked,
      lockLoadLine: formValues.LockLoadLine === 'on',
      lockedSlope: formValues.LockLoadLine === 'on' ? lockedSlope : '',
      lockedIntercept: formValues.LockLoadLine === 'on' ? lockedIntercept : '',
      fanControls_C: objSD_C?.FanControls,
      fanControlsSource_C: _el(form, 'cmbFanControls_C') ? 'cmbFanControls_C' : undefined,
      bfnKw_C: objSD_C?.BFn_kw,
      condKw_C: objSD_C?.Cond_kw,
      condSource_C: _el(form, 'txtCond_kw_C') ? 'txtCond_kw_C' : undefined,
      formDefaults: {
        chkAdvancedControls_defaultChecked: _defaultChecked('chkAdvancedControls'),
        tdAdvancedControls_defaultLabel: _textById(form, 'tdAdvancedControls_default'),
        chkEconomizer_C_defaultChecked: _defaultChecked('chkEconomizer_C'),
        chkEconomizer_S_defaultChecked: _defaultChecked('chkEconomizer_S'),
        chkDOE2_Curves_defaultChecked: _defaultChecked('chkDOE2_Curves'),
        chkLockLoadLine_defaultChecked: _defaultChecked('chkLockLoadLine'),
        cmbSpecific_RTU_C_defaultValue: _defaultValue('cmbSpecific_RTU_C'),
        cmbFanControls_C_defaultValue: _defaultValue('cmbFanControls_C') ?? _textById(form, 'tdFanControls_C_default'),
        tdBFn_kw_C_defaultValue: _numFromText(_textById(form, 'tdBFn_kw_C_default')),
        tdCond_kw_C_defaultValue: _numFromText(_textById(form, 'tdCond_kw_C_default')),
        totalCap_kBtuh_default: _numFromText(_textById(form, 'tdTotalCapacity_default')),
        idb_F_default: _numFromText(_textById(form, 'tdIDB_default')),
        nUnits_default: _numFromText(_textById(form, 'tdNUnits_default')),
        ventilationValue_default: _numFromText(_textById(form, 'tdVentilationValue_default')),
        ventilationUnits_default: _textById(form, 'tdVentilationUnits_default'),
        cmbVentilationUnits_defaultValue: _defaultValue('cmbVentilationUnits'),
        txtVentilationValue_defaultValue: _defaultValue('txtVentilationValue'),
        txtVentilationValue_NotAdvanced_defaultValue: _defaultValue('txtVentilationValue_NotAdvanced'),
        cmbTrackOHR_defaultValue: _defaultValue('cmbTrackOHR'),
        cmbIRH_defaultValue: _defaultValue('cmbIRH'),
        txtSI_Fraction_defaultValue: _defaultValue('txtSI_Fraction'),
        txtSI_Fraction_NotAdvanced_defaultValue: _defaultValue('txtSI_Fraction_NotAdvanced'),
        cmbOversizePercent_defaultValue: _defaultValue('cmbOversizePercent'),
        cmbN_Affinity_defaultValue: _defaultValue('cmbN_Affinity'),
        cmbST_Ratio_C_defaultValue: _defaultValue('cmbST_Ratio_C'),
        txtSlope_defaultValue: _defaultValue('txtSlope'),
        txtIntercept_defaultValue: _defaultValue('txtIntercept'),
      },
      raw: {
        // These are the raw form values whose presence/rounding often differs between Advanced ON vs OFF.
        txtVentilationValue: _val(form, 'txtVentilationValue'),
        txtVentilationValue_NotAdvanced: _val(form, 'txtVentilationValue_NotAdvanced'),
        txtSI_Fraction: _val(form, 'txtSI_Fraction'),
        txtSI_Fraction_NotAdvanced: _val(form, 'txtSI_Fraction_NotAdvanced'),

        txtBFn_kw_C: _val(form, 'txtBFn_kw_C'),
        txtBFn_kw_C_NotAdvanced: _val(form, 'txtBFn_kw_C_NotAdvanced'),
        txtBFn_kw_C_hidden: _val(form, 'txtBFn_kw_C_hidden'),

        txtBFn_kw_S: _val(form, 'txtBFn_kw_S'),
        txtBFn_kw_S_NotAdvanced: _val(form, 'txtBFn_kw_S_NotAdvanced'),
        txtBFn_kw_S_hidden: _val(form, 'txtBFn_kw_S_hidden'),

        cmbST_Ratio_C: _val(form, 'cmbST_Ratio_C'),
        txtST_Ratio_C_hidden: _val(form, 'txtST_Ratio_C_hidden'),
        cmbST_Ratio_S: _val(form, 'cmbST_Ratio_S'),
        txtST_Ratio_S_hidden: _val(form, 'txtST_Ratio_S_hidden'),

        chkLockLoadLine: _val(form, 'chkLockLoadLine'),
        txtSlope: _val(form, 'txtSlope'),
        txtIntercept: _val(form, 'txtIntercept'),
        txtBM_Slope_hidden: _val(form, 'txtBM_Slope_hidden'),
        txtBM_Intercept_hidden: _val(form, 'txtBM_Intercept_hidden'),
        txtBM_VentSlopeFraction_hidden: _val(form, 'txtBM_VentSlopeFraction_hidden'),
      },
      nUnits,
      doe2Curves: objSD_C.DOE2_Curves,
      doe2CurvesSource: doe2Source,
      specificRtu_C: objSD_C.Specific_RTU,
      specificRtuSource_C,
      specificRtu_S: objSD_S.Specific_RTU,
      specificRtuSource_S,
      ventilationUnitsLabel: ventUnitsLabel,
      ventilationValue: ventilationValueRaw,
      trackOhr,
      irhPct,
      sandI_fraction: sandIFraction,
      oversizePct,
      oversizingLRF,
      bfnKwForA0_C,
      bfnSource_C: bfnPick_C.source,
      bfnKwForA0_S,
      bfnSource_S: bfnPick_S.source,
      stAtTest_C,
      stSource_C: stPick_C.source,
      stAtTest_S,
      stSource_S: stPick_S.source,
      nAffinity,
      nAffinitySource: nPick.source,
      ventilationValueSource: ventValuePick.source,
      irhSource: irhPick.source,
      oversizeSource: oversizePick.source,
      sandISource: sandIPick.source,
      econCandidate: _checked(form, 'chkEconomizer_C'),
      econStandard: _checked(form, 'chkEconomizer_S'),
      spreadsheet: {
        candidate: _modelDebug(objSD_C),
        standard: _modelDebug(objSD_S),
      },
    };

    _lastRunInputs = runInputs;

    if (!ll?.ok) {
      return '<div class="graph"><h2>Client-side engine error</h2><pre>' + _escapeHtml(ll?.errorMessage || 'Load line computation failed') + '</pre></div>';
    }

    // =====================================================
    // Phase 2: Economizer ON/OFF (per checkbox), staged systems.
    // =====================================================
    const ventCFM = (formValues.VentilationUnits === 'CFM') ? formValues.VentilationValue : (cfm * (formValues.VentilationValue / 100));
    const ventilationFraction = ventCFM / cfm;

    // Run the bin-by-bin energy calculation for one RTU system.
    // Loops over occupied and unoccupied hours, computing economizer
    // contribution, staging, condenser/fan/aux energy, and peak demand.
    // Returns { annualTotal, annualCondenser, annualEFan, annualAux,
    //           peakDemand, binsOcc, binsUnocc, ... }.
    function runOneSystem(objSD) {
      const SP = new StagePair();
      SP.A.ResetToFullLoad();
      SP.SetCondType(objSD);
      SP.EconomizerRunning = false;

      let annualCondenser = 0;
      let annualEFan = 0;
      let annualAux = 0;
      let auxHoursOcc = 0;
      let auxHoursUnocc = 0;
      let annualCondenser_Occ = 0;
      let annualCondenser_Unocc = 0;
      let annualEFan_Occ = 0;
      let annualEFan_Unocc = 0;

      let econHours_Occ = 0;
      let econHours_Unocc = 0;
      const binsEconOcc = new Map();
      const binsEconUnocc = new Map();

      const binsOcc = new Map();
      const binsUnocc = new Map();
      const integratedDebugOcc = [];
      const integratedDebugUnocc = [];
      let peakDemand = 0;
      let totalCoolingHours = 0;
      let sample = null;

      const weatherCity = getWeatherRecords(weatherNorm, state, city, schedule) || [];
      const weatherByTemp = new Map();
      for (const r of weatherCity) {
        weatherByTemp.set(Number(r.Temp_Outdoor_DB), r);
      }

      // Legacy hour accounting: occupied hours drive condenser+fan energy.
      const occHours = getHours(weatherNorm, state, city, schedule, 'Occupied');
      const unoccHours = getHours(weatherNorm, state, city, schedule, 'UnOccupied');
      const allHours = getHours(weatherNorm, state, city, schedule, 'Total');

      for (const [tempKey, occH] of occHours.entries()) {
        const odb = Number(tempKey);
        const hoursOcc = Number(occH);
        if (!Number.isFinite(odb) || !Number.isFinite(hoursOcc) || hoursOcc <= 0) continue;

        const wrec = weatherByTemp.get(odb);
        const owb = Number(wrec?.Temp_Coinc_WB);
        if (!Number.isFinite(owb)) continue;

        const ohr = Phr_wb(odb, owb, ll.design.pressure);
        const insideRH = IRH_Track_OR_Set(trackOhr, ohr, idb, irh, ll.design.pressure);
        const insideHR = Phr_rh(idb, insideRH, ll.design.pressure);

        const nonVent = (ll.slope * (odb - idb)) + ll.intercept;
        const sensVent = SensVentLoad(ventCFM, ohr, odb, idb, ll.design.pressure);
        const totalSens = nonVent + sensVent;

        const runBin = (totalSens > 0) || ((idb - odb) <= 5);

        // If there is no cooling run in this bin, legacy does not include a fan-energy entry
        // for this bin in the cooling energy output tables. (Aux is still annualized separately.)
        if (!runBin) {
          continue;
        }

        // -----------------------------------------------------
        // Economizer logic (legacy-equivalent): compute econ sensible
        // cooling contribution and determine mixed air entering coil.
        // -----------------------------------------------------
        const BOC = { DB: odb, WB: owb, HR: ohr, BP: ll.design.pressure };

        let econoSensCoolingLoad = 0;
        SP.IntegratedState = '';
        SP.EconomizerRunning = true;

        if (objSD.Economizer === 'on') {
          // Temperature-based control (legacy current engine behavior).
          if (odb >= idb || totalSens <= 0) {
            econoSensCoolingLoad = 0;
            SP.EconomizerRunning = false;
          }

          // Match source/legacy behavior for variable-capacity systems: require the outside air to be
          // more than 5F cooler than the setpoint before using economizer cooling.
          // (Occupied: IDB=75 => ODB=70 disables; Unoccupied: IDB=80 => ODB=70 enables.)
          const isVC =
            (typeof objSD?.FanControls === 'string' && objSD.FanControls.slice(0, 1) === 'V') ||
            String(objSD?.Specific_RTU || '') === 'Variable-Speed Compressor';
          if (isVC && ((idb - odb) <= 5)) {
            econoSensCoolingLoad = 0;
            SP.EconomizerRunning = false;
          }
        } else {
          econoSensCoolingLoad = 0;
          SP.EconomizerRunning = false;
        }

        // Economizer sensible cooling load (heat gain, negative when ODB < IDB).
        if (SP.EconomizerRunning && (cfm > ventCFM)) {
          const econFanFraction = FF(objSD, 0, SP, 'C-Off', true, odb, ventilationFraction);
          econoSensCoolingLoad = SensVentLoad(((cfm * econFanFraction) - ventCFM), ohr, odb, idb, ll.design.pressure);
        } else {
          econoSensCoolingLoad = 0;
        }

        // Remaining sensible load after economizer.
        let remaining = totalSens + econoSensCoolingLoad;
        if (remaining < 0) remaining = 0;

        // Mixed air entering coil depends on economizer state.
        let freshAirToCoilCFM = SP.EconomizerRunning ? cfm : ventCFM;
        let mix = Mixer2(objSD, idb, insideHR, odb, ohr, ll.design.pressure, freshAirToCoilCFM, cfm);
        let BCC = { ODB: odb, EWB: mix.EWB, EDB: mix.EDB, EHR: mix.EHR };

        // Reset peaks per-bin.
        const ppCondA = { value: 0 };
        const ppCondBmA = { value: 0 };
        const ppCondB = { value: 0 };
        const ppEFan = { value: 0 };

        // Stage runtimes.
        let integratedUsed = false;
        let integratedRuntime = 0;
        let dblStageLevel = 0;
        const isVC = typeof objSD?.FanControls === 'string' && objSD.FanControls.slice(0, 1) === 'V';
        if (SP.EconomizerRunning && remaining > 0 && !isVC) {
          // Integrated economizer check (economizer + DX). If it fails, fall back to normal DX-only.
          const stageArray = Array.isArray(objSD.StageLevels) ? objSD.StageLevels : [1.0];
          SP.A.CapacityFraction = Number(stageArray[0] ?? 1.0);
          const intRes = _capacityLevelIntegrated_Staged(
            objSD,
            SP,
            BOC,
            nonVent,
            idb,
            ll.design.pressure,
            totalCap,
            cfm,
            ventCFM,
            ventilationFraction
          );

          if (intRes.ok) {
            // Use the integrated runtime for stage A, and force A_only mode.
            SP.pair_mode = 'A_only';
            SP.A.RunTime = intRes.runtime;
            integratedUsed = true;
            integratedRuntime = intRes.runtime;
            SP.A.FlowFraction = FF(objSD, SP.A.CapacityFraction, SP, 'C-On', false, odb, ventilationFraction);
            dblStageLevel = intRes.runtime;
          } else {
            SP.EconomizerRunning = false;
            remaining = totalSens;
            // Legacy: rerun mixing for non-economizer operation when integrated fails.
            freshAirToCoilCFM = ventCFM;
            mix = Mixer2(objSD, idb, insideHR, odb, ohr, ll.design.pressure, freshAirToCoilCFM, cfm);
            BCC = { ODB: odb, EWB: mix.EWB, EDB: mix.EDB, EHR: mix.EHR };
            dblStageLevel = StageLevel(objSD, BCC, remaining < 0 ? 0 : remaining, SP, {
              totalCap,
              cfm,
              pressure: ll.design.pressure,
              ventilationFraction,
            });
          }
        } else {
          dblStageLevel = StageLevel(objSD, BCC, remaining, SP, {
            totalCap,
            cfm,
            pressure: ll.design.pressure,
            ventilationFraction,
          });
        }

        // Condenser energy.
        let eCond = 0;
        let condKw = 0;
        if (SP.pair_mode === 'A_only') {
          condKw = CondenserPower_PL_kW(objSD, SP.A, odb, BCC.EWB, BCC.EDB, ppCondA, totalCap);
          eCond = hoursOcc * condKw * SP.A.RunTime;
          if (SP.A.RunTime === 0) ppCondA.value = 0;
        } else if (SP.pair_mode === 'A_and_BmA') {
          const condA = CondenserPower_PL_kW(objSD, SP.A, odb, BCC.EWB, BCC.EDB, ppCondA, totalCap);
          const condBmA = CondenserPower_PL_kW(objSD, SP.BmA, odb, BCC.EWB, BCC.EDB, ppCondBmA, totalCap);
          condKw = condA + (condBmA * SP.BmA.RunTime);
          eCond = hoursOcc * (condA + (condBmA * SP.BmA.RunTime));
        } else if (SP.pair_mode === 'B_only') {
          condKw = CondenserPower_PL_kW(objSD, SP.B, odb, BCC.EWB, BCC.EDB, ppCondB, totalCap);
          eCond = hoursOcc * condKw * SP.B.RunTime;
        }

        if (integratedUsed) {
          integratedDebugOcc.push({
            odb,
            hours: hoursOcc,
            owb,
            ohr,
            idb,
            insideRH,
            insideHR,
            freshAirToCoilCFM,
            nonVent,
            sensVent,
            totalSens,
            econoSensCoolingLoad,
            remaining,
            pair_mode: SP.pair_mode,
            capFractionA: SP.A?.CapacityFraction,
            flowFractionA: SP.A?.FlowFraction,
            loadFractionA: SP.A?.LoadFraction,
            runTimeA: SP.A?.RunTime,
            bcc_edb: BCC.EDB,
            bcc_ewb: BCC.EWB,
            bcc_ehr: BCC.EHR,
            condKw,
            eCond,
          });
        }

        // Fan energy (match legacy ASP).
        SP.FlowFraction_CompOff = FF(objSD, 0, SP, 'C-Off', SP.EconomizerRunning, odb, ventilationFraction);
        let eFan = 0;
        const fanControls = String(objSD.FanControls || '');
        const isAlwaysOn = fanControls.includes('Always ON');
        const isCycles = fanControls.includes('Cycles With Compressor');

        if (SP.pair_mode === 'A_only') {
          const rtAraw = Math.max(0, Number(SP.A.RunTime) || 0);
          if (isAlwaysOn) {
            if (rtAraw < 1.0) {
              eFan = hoursOcc * (
                FanPower_PL_kW(objSD, SP.A.FlowFraction, ppEFan) * rtAraw +
                FanPower_PL_kW(objSD, SP.FlowFraction_CompOff, ppEFan) * (1.0 - rtAraw)
              );
            } else {
              eFan = hoursOcc * FanPower_PL_kW(objSD, SP.A.FlowFraction, ppEFan) * rtAraw;
            }
          } else if (isCycles) {
            const rtAclamped = Math.min(1, rtAraw);
            eFan = hoursOcc * (
              FanPower_PL_kW(objSD, SP.A.FlowFraction, ppEFan) * rtAclamped +
              FanPower_PL_kW(objSD, SP.FlowFraction_CompOff, ppEFan) * (1.0 - rtAclamped)
            );
          }
        } else if (SP.pair_mode === 'A_and_BmA') {
          const rtB = Math.max(0, Math.min(1, Number(SP.B.RunTime) || 0));
          const rtA = 1.0 - rtB;
          if (isAlwaysOn || isCycles) {
            eFan = hoursOcc * (
              FanPower_PL_kW(objSD, SP.A.FlowFraction, ppEFan) * rtA +
              FanPower_PL_kW(objSD, SP.B.FlowFraction, ppEFan) * rtB
            );
          }
        } else if (SP.pair_mode === 'B_only') {
          const rtBraw = Math.max(0, Number(SP.B.RunTime) || 0);
          if (isAlwaysOn) {
            eFan = hoursOcc * FanPower_PL_kW(objSD, SP.B.FlowFraction, ppEFan) * rtBraw;
          } else if (isCycles) {
            eFan = hoursOcc * FanPower_PL_kW(objSD, SP.B.FlowFraction, ppEFan) * rtBraw;
          }
        }

        // Aux energy: attribute to occupied hours here; unoccupied hours are handled
        // in the setback/unoccupied pass to avoid double-counting or counting bins
        // that don't run during unoccupied.
        const eAux = hoursOcc * objSD.Aux_kw;
        auxHoursOcc += hoursOcc;

        annualCondenser += eCond;
        annualEFan += eFan;
        annualAux += eAux;
        annualCondenser_Occ += eCond;
        annualEFan_Occ += eFan;
        totalCoolingHours += hoursOcc;

        if (SP.EconomizerRunning) econHours_Occ += hoursOcc;
        const prevEconOcc = binsEconOcc.get(odb) || { odb, hours: 0, econHours: 0, econLoad: 0, remaining: 0, integrated: 0 };
        prevEconOcc.hours += hoursOcc;
        prevEconOcc.econHours += SP.EconomizerRunning ? hoursOcc : 0;
        prevEconOcc.econLoad += econoSensCoolingLoad;
        prevEconOcc.remaining += remaining;
        prevEconOcc.integrated += integratedUsed ? integratedRuntime : 0;
        binsEconOcc.set(odb, prevEconOcc);

        // Compute additional per-bin fields for bin tables display.
        const stRatioA = Number(SP.A?.ST_Ratio ?? 0);
        const tcfA = Number(SP.A?.TCap_CF ?? 0);
        const effCfA = Number(SP.A?.Efficiency_CF ?? 0);
        const pcfA = Number(SP.A?.CondPower_CF ?? 0);
        // OCF = PCF / (TCF * (S/T_entering / S/T_test))  (ASP Engine.asp line 1867)
        const stRatioTest = Number(objSD.ST_Ratio_AtTest ?? 0);
        const ocfA = (tcfA !== 0 && stRatioTest !== 0 && stRatioA !== 0)
          ? pcfA / (tcfA * (stRatioA / stRatioTest))
          : 0;
        // Latent load = (remaining / S/T) * (1 - S/T)  (ASP Engine.asp line 1922)
        const latentLoad = (stRatioA > 0 && remaining > 0)
          ? (remaining / stRatioA) * (1 - stRatioA)
          : 0;
        // ERH from entering conditions
        const bcc_erh = (BCC.EDB && BCC.EHR) ? Prh_hr(BCC.EDB, BCC.EHR, ll.design.pressure) : 0;
        // Demand kW for this bin
        const demandKw = (ppCondA.value + ppCondBmA.value + ppCondB.value) + ppEFan.value + objSD.Aux_kw;

        // Bin debug record (occupied)
        const prevOcc = binsOcc.get(odb) || {
          odb,
          hours: 0,
          eFan: 0,
          eCond: 0,
          eAux: 0,
          nonVent: 0,
          sensVent: 0,
          totalSens: 0,
          pair_mode: '',
          rtA: 0,
          rtB: 0,
          lfA: 0,
          ffA: 0,
          capFracA: 0,
          capFracB: 0,
          sensCapA: 0,
          sensCapB: 0,
          remaining: 0,
          bcc_edb: 0,
          bcc_ewb: 0,
          condKw: 0,
          condPowerCF_A: 0,
          effCF_A: 0,
          // Additional fields for bin tables
          owb: 0, ohr: 0, ihr: 0, irh: 0,
          econLoad: 0, latentLoad: 0,
          bcc_ehr: 0, bcc_erh: 0,
          tcf: 0, stRatio: 0, invEcf: 0,
          stageLevel: 0, pcf: 0, ocf: 0,
          demandKw: 0,
          totalHours: 0,
        };
        prevOcc.hours += hoursOcc;
        prevOcc.totalHours = Number(allHours.get(odb) || 0);
        prevOcc.eFan += eFan;
        prevOcc.eCond += eCond;
        prevOcc.eAux += hoursOcc * objSD.Aux_kw;
        prevOcc.nonVent = nonVent;
        prevOcc.sensVent = sensVent;
        prevOcc.totalSens = totalSens;
        prevOcc.pair_mode = String(SP.pair_mode || '');
        prevOcc.rtA = Number(SP.A?.RunTime ?? 0);
        prevOcc.rtB = Number(SP.BmA?.RunTime ?? 0);
        prevOcc.lfA = Number(SP.A?.LoadFraction ?? 0);
        prevOcc.ffA = Number(SP.A?.FlowFraction ?? 0);
        prevOcc.capFracA = Number(SP.A?.CapacityFraction ?? 0);
        prevOcc.capFracB = Number(SP.B?.CapacityFraction ?? 0);
        prevOcc.sensCapA = Number(SP.A?.SensCap_KBtuH ?? 0);
        prevOcc.sensCapB = Number(SP.B?.SensCap_KBtuH ?? 0);
        prevOcc.remaining = remaining;
        prevOcc.bcc_edb = Number(BCC?.EDB ?? 0);
        prevOcc.bcc_ewb = Number(BCC?.EWB ?? 0);
        prevOcc.condKw = Number(condKw ?? 0);
        prevOcc.condPowerCF_A = pcfA;
        prevOcc.effCF_A = effCfA;
        // Additional fields for bin tables
        prevOcc.owb = owb;
        prevOcc.ohr = ohr;
        prevOcc.ihr = insideHR;
        prevOcc.irh = insideRH;
        prevOcc.econLoad = econoSensCoolingLoad;
        prevOcc.latentLoad = latentLoad;
        prevOcc.bcc_ehr = Number(BCC?.EHR ?? 0);
        prevOcc.bcc_erh = bcc_erh;
        prevOcc.tcf = tcfA;
        prevOcc.stRatio = stRatioA;
        prevOcc.invEcf = (effCfA !== 0) ? (1 / effCfA) : 0;
        prevOcc.stageLevel = dblStageLevel;
        prevOcc.pcf = pcfA;
        prevOcc.ocf = ocfA;
        prevOcc.demandKw = demandKw;
        binsOcc.set(odb, prevOcc);

        if (!sample && (eCond > 0 || eFan > 0)) {
          sample = {
            odb,
            hours: hoursOcc,
            pair_mode: SP.pair_mode,
            rtA: SP.A.RunTime,
            rtB: SP.BmA.RunTime,
            lfA: SP.A.LoadFraction,
            ffA: SP.A.FlowFraction,
            ffOff: SP.FlowFraction_CompOff,
            bfn_kw: objSD.BFn_kw,
          };
        }

        if (demandKw > peakDemand) peakDemand = demandKw;
      }

      // =====================================================
      // Legacy setback calcs: process unoccupied hours as if occupied,
      // using the warmer IDB setpoint and forcing fan to "Cycles With Compressor".
      // Equivalent to SetBackCalcs + RunBinCalcs(objUnOccupiedHours_DB, objUnOccupiedHours_Zero_DB).
      // =====================================================
      if (Number.isFinite(idbSetBack) && idbSetBack > 0) {
        const originalFanControls = objSD.FanControls;
        const fcLead = String(originalFanControls || '').substring(0, 1);
        if (fcLead === '1') objSD.FanControls = '1-Spd: Cycles With Compressor';
        if (fcLead === 'N') objSD.FanControls = 'N-Spd: Cycles With Compressor';

        try {
          for (const [tempKey, unoccH] of unoccHours.entries()) {
            const odb = Number(tempKey);
            const hoursUnocc = Number(unoccH);
            if (!Number.isFinite(odb) || !Number.isFinite(hoursUnocc) || hoursUnocc <= 0) continue;

            const wrec = weatherByTemp.get(odb);
            const owb = Number(wrec?.Temp_Coinc_WB);
            if (!Number.isFinite(owb)) continue;

            const ohr = Phr_wb(odb, owb, ll.design.pressure);
            const insideRH = IRH_Track_OR_Set(trackOhr, ohr, idbUnoccupied, irh, ll.design.pressure);
            const insideHR = Phr_rh(idbUnoccupied, insideRH, ll.design.pressure);

            const nonVent = (ll.slope * (odb - idbUnoccupied)) + ll.intercept;
            const sensVent = SensVentLoad(ventCFM, ohr, odb, idbUnoccupied, ll.design.pressure);
            const totalSens = nonVent + sensVent;

            const runBin = (totalSens > 0) || ((idbUnoccupied - odb) <= 5);
            if (!runBin) continue;

            const BOC = { DB: odb, WB: owb, HR: ohr, BP: ll.design.pressure };

            let econoSensCoolingLoad = 0;
            SP.IntegratedState = '';
            SP.EconomizerRunning = true;
            if (objSD.Economizer === 'on') {
              if (odb >= idbUnoccupied || totalSens <= 0) {
                econoSensCoolingLoad = 0;
                SP.EconomizerRunning = false;
              }

              // Variable-capacity economizer gating: require outside air to be >5F cooler than the
              // (unoccupied) setpoint before applying economizer cooling.
              const isVC =
                (typeof objSD?.FanControls === 'string' && objSD.FanControls.slice(0, 1) === 'V') ||
                String(objSD?.Specific_RTU || '') === 'Variable-Speed Compressor';
              if (isVC && ((idbUnoccupied - odb) <= 5)) {
                econoSensCoolingLoad = 0;
                SP.EconomizerRunning = false;
              }
            } else {
              econoSensCoolingLoad = 0;
              SP.EconomizerRunning = false;
            }

            // Economizer sensible cooling load (heat gain, negative when ODB < IDB).
            if (SP.EconomizerRunning && (cfm > ventCFM)) {
              const econFanFraction = FF(objSD, 0, SP, 'C-Off', true, odb, ventilationFraction);
              econoSensCoolingLoad = SensVentLoad(((cfm * econFanFraction) - ventCFM), ohr, odb, idbUnoccupied, ll.design.pressure);
            } else {
              econoSensCoolingLoad = 0;
            }

            let remaining = totalSens + econoSensCoolingLoad;
            if (remaining < 0) remaining = 0;

            let freshAirToCoilCFM = SP.EconomizerRunning ? cfm : ventCFM;
            let mix = Mixer2(objSD, idbUnoccupied, insideHR, odb, ohr, ll.design.pressure, freshAirToCoilCFM, cfm);
            let BCC = { ODB: odb, EWB: mix.EWB, EDB: mix.EDB, EHR: mix.EHR };

            const ppCondA = { value: 0 };
            const ppCondBmA = { value: 0 };
            const ppCondB = { value: 0 };
            const ppEFan = { value: 0 };

            let integratedUsed = false;
            let integratedRuntime = 0;
            let dblStageLevel = 0;
            const isVC = typeof objSD?.FanControls === 'string' && objSD.FanControls.slice(0, 1) === 'V';
            if (SP.EconomizerRunning && remaining > 0 && !isVC) {
              const stageArray = Array.isArray(objSD.StageLevels) ? objSD.StageLevels : [1.0];
              SP.A.CapacityFraction = Number(stageArray[0] ?? 1.0);
              const intRes = _capacityLevelIntegrated_Staged(
                objSD,
                SP,
                BOC,
                nonVent,
                idbUnoccupied,
                ll.design.pressure,
                totalCap,
                cfm,
                ventCFM,
                ventilationFraction
              );

              if (intRes.ok) {
                SP.pair_mode = 'A_only';
                SP.A.RunTime = intRes.runtime;
                integratedUsed = true;
                integratedRuntime = intRes.runtime;
                SP.A.FlowFraction = FF(objSD, SP.A.CapacityFraction, SP, 'C-On', false, odb, ventilationFraction);
                dblStageLevel = intRes.runtime;
              } else {
                SP.EconomizerRunning = false;
                remaining = totalSens;
                freshAirToCoilCFM = ventCFM;
                mix = Mixer2(objSD, idbUnoccupied, insideHR, odb, ohr, ll.design.pressure, freshAirToCoilCFM, cfm);
                BCC = { ODB: odb, EWB: mix.EWB, EDB: mix.EDB, EHR: mix.EHR };
                dblStageLevel = StageLevel(objSD, BCC, remaining < 0 ? 0 : remaining, SP, {
                  totalCap,
                  cfm,
                  pressure: ll.design.pressure,
                  ventilationFraction,
                });
              }
            } else {
              dblStageLevel = StageLevel(objSD, BCC, remaining, SP, {
                totalCap,
                cfm,
                pressure: ll.design.pressure,
                ventilationFraction,
              });
            }

            let eCond = 0;
            let condKw = 0;
            if (SP.pair_mode === 'A_only') {
              condKw = CondenserPower_PL_kW(objSD, SP.A, odb, BCC.EWB, BCC.EDB, ppCondA, totalCap);
              eCond = hoursUnocc * condKw * SP.A.RunTime;
              if (SP.A.RunTime === 0) ppCondA.value = 0;
            } else if (SP.pair_mode === 'A_and_BmA') {
              const condA = CondenserPower_PL_kW(objSD, SP.A, odb, BCC.EWB, BCC.EDB, ppCondA, totalCap);
              const condBmA = CondenserPower_PL_kW(objSD, SP.BmA, odb, BCC.EWB, BCC.EDB, ppCondBmA, totalCap);
              condKw = condA + (condBmA * SP.BmA.RunTime);
              eCond = hoursUnocc * (condA + (condBmA * SP.BmA.RunTime));
            } else if (SP.pair_mode === 'B_only') {
              condKw = CondenserPower_PL_kW(objSD, SP.B, odb, BCC.EWB, BCC.EDB, ppCondB, totalCap);
              eCond = hoursUnocc * condKw * SP.B.RunTime;
            }

            if (integratedUsed) {
              integratedDebugUnocc.push({
                odb,
                hours: hoursUnocc,
                owb,
                ohr,
                idb: idbUnoccupied,
                insideRH,
                insideHR,
                freshAirToCoilCFM,
                nonVent,
                sensVent,
                totalSens,
                econoSensCoolingLoad,
                remaining,
                pair_mode: SP.pair_mode,
                capFractionA: SP.A?.CapacityFraction,
                flowFractionA: SP.A?.FlowFraction,
                loadFractionA: SP.A?.LoadFraction,
                runTimeA: SP.A?.RunTime,
                bcc_edb: BCC.EDB,
                bcc_ewb: BCC.EWB,
                bcc_ehr: BCC.EHR,
                condKw,
                eCond,
              });
            }

            SP.FlowFraction_CompOff = FF(objSD, 0, SP, 'C-Off', SP.EconomizerRunning, odb, ventilationFraction);
            let eFan = 0;
            const fanControls = String(objSD.FanControls || '');
            const isAlwaysOn = fanControls.includes('Always ON');
            const isCycles = fanControls.includes('Cycles With Compressor');

            if (SP.pair_mode === 'A_only') {
              const rtAraw = Math.max(0, Number(SP.A.RunTime) || 0);
              if (isAlwaysOn) {
                if (rtAraw < 1.0) {
                  eFan = hoursUnocc * (
                    FanPower_PL_kW(objSD, SP.A.FlowFraction, ppEFan) * rtAraw +
                    FanPower_PL_kW(objSD, SP.FlowFraction_CompOff, ppEFan) * (1.0 - rtAraw)
                  );
                } else {
                  eFan = hoursUnocc * FanPower_PL_kW(objSD, SP.A.FlowFraction, ppEFan) * rtAraw;
                }
              } else if (isCycles) {
                const rtAclamped = Math.min(1, rtAraw);
                eFan = hoursUnocc * (
                  FanPower_PL_kW(objSD, SP.A.FlowFraction, ppEFan) * rtAclamped +
                  FanPower_PL_kW(objSD, SP.FlowFraction_CompOff, ppEFan) * (1.0 - rtAclamped)
                );
              }
            } else if (SP.pair_mode === 'A_and_BmA') {
              const rtB = Math.max(0, Math.min(1, Number(SP.B.RunTime) || 0));
              const rtA = 1.0 - rtB;
              if (isAlwaysOn || isCycles) {
                eFan = hoursUnocc * (
                  FanPower_PL_kW(objSD, SP.A.FlowFraction, ppEFan) * rtA +
                  FanPower_PL_kW(objSD, SP.B.FlowFraction, ppEFan) * rtB
                );
              }
            } else if (SP.pair_mode === 'B_only') {
              const rtBraw = Math.max(0, Number(SP.B.RunTime) || 0);
              if (isAlwaysOn || isCycles) {
                eFan = hoursUnocc * FanPower_PL_kW(objSD, SP.B.FlowFraction, ppEFan) * rtBraw;
              }
            }

            annualCondenser += eCond;
            annualEFan += eFan;
            annualAux += (hoursUnocc * objSD.Aux_kw);
            auxHoursUnocc += hoursUnocc;
            annualCondenser_Unocc += eCond;
            annualEFan_Unocc += eFan;

            if (SP.EconomizerRunning) econHours_Unocc += hoursUnocc;
            const prevEconUnocc = binsEconUnocc.get(odb) || { odb, hours: 0, econHours: 0, econLoad: 0, remaining: 0, integrated: 0 };
            prevEconUnocc.hours += hoursUnocc;
            prevEconUnocc.econHours += SP.EconomizerRunning ? hoursUnocc : 0;
            prevEconUnocc.econLoad += econoSensCoolingLoad;
            prevEconUnocc.remaining += remaining;
            prevEconUnocc.integrated += integratedUsed ? integratedRuntime : 0;
            binsEconUnocc.set(odb, prevEconUnocc);

            // Compute additional per-bin fields for bin tables display.
            const stRatioA = Number(SP.A?.ST_Ratio ?? 0);
            const tcfA = Number(SP.A?.TCap_CF ?? 0);
            const effCfA = Number(SP.A?.Efficiency_CF ?? 0);
            const pcfA = Number(SP.A?.CondPower_CF ?? 0);
            const stRatioTest = Number(objSD.ST_Ratio_AtTest ?? 0);
            const ocfA = (tcfA !== 0 && stRatioTest !== 0 && stRatioA !== 0)
              ? pcfA / (tcfA * (stRatioA / stRatioTest))
              : 0;
            const latentLoad = (stRatioA > 0 && remaining > 0)
              ? (remaining / stRatioA) * (1 - stRatioA)
              : 0;
            const bcc_erh = (BCC.EDB && BCC.EHR) ? Prh_hr(BCC.EDB, BCC.EHR, ll.design.pressure) : 0;
            const demandKw = (ppCondA.value + ppCondBmA.value + ppCondB.value) + ppEFan.value + objSD.Aux_kw;

            // Bin debug record (unoccupied)
            const prevUnocc = binsUnocc.get(odb) || {
              odb,
              hours: 0,
              eFan: 0,
              eCond: 0,
              eAux: 0,
              nonVent: 0,
              sensVent: 0,
              totalSens: 0,
              pair_mode: '',
              rtA: 0,
              rtB: 0,
              lfA: 0,
              ffA: 0,
              capFracA: 0,
              capFracB: 0,
              sensCapA: 0,
              sensCapB: 0,
              remaining: 0,
              bcc_edb: 0,
              bcc_ewb: 0,
              condKw: 0,
              condPowerCF_A: 0,
              effCF_A: 0,
              // Additional fields for bin tables
              owb: 0, ohr: 0, ihr: 0, irh: 0,
              econLoad: 0, latentLoad: 0,
              bcc_ehr: 0, bcc_erh: 0,
              tcf: 0, stRatio: 0, invEcf: 0,
              stageLevel: 0, pcf: 0, ocf: 0,
              demandKw: 0,
              totalHours: 0,
            };
            prevUnocc.hours += hoursUnocc;
            prevUnocc.totalHours = Number(allHours.get(odb) || 0);
            prevUnocc.eFan += eFan;
            prevUnocc.eCond += eCond;
            prevUnocc.eAux += hoursUnocc * objSD.Aux_kw;
            prevUnocc.nonVent = nonVent;
            prevUnocc.sensVent = sensVent;
            prevUnocc.totalSens = totalSens;
            prevUnocc.pair_mode = String(SP.pair_mode || '');
            prevUnocc.rtA = Number(SP.A?.RunTime ?? 0);
            prevUnocc.rtB = Number(SP.BmA?.RunTime ?? 0);
            prevUnocc.lfA = Number(SP.A?.LoadFraction ?? 0);
            prevUnocc.ffA = Number(SP.A?.FlowFraction ?? 0);
            prevUnocc.capFracA = Number(SP.A?.CapacityFraction ?? 0);
            prevUnocc.capFracB = Number(SP.B?.CapacityFraction ?? 0);
            prevUnocc.sensCapA = Number(SP.A?.SensCap_KBtuH ?? 0);
            prevUnocc.sensCapB = Number(SP.B?.SensCap_KBtuH ?? 0);
            prevUnocc.remaining = remaining;
            prevUnocc.bcc_edb = Number(BCC?.EDB ?? 0);
            prevUnocc.bcc_ewb = Number(BCC?.EWB ?? 0);
            prevUnocc.condKw = Number(condKw ?? 0);
            prevUnocc.condPowerCF_A = pcfA;
            prevUnocc.effCF_A = effCfA;
            // Additional fields for bin tables
            prevUnocc.owb = owb;
            prevUnocc.ohr = ohr;
            prevUnocc.ihr = insideHR;
            prevUnocc.irh = insideRH;
            prevUnocc.econLoad = econoSensCoolingLoad;
            prevUnocc.latentLoad = latentLoad;
            prevUnocc.bcc_ehr = Number(BCC?.EHR ?? 0);
            prevUnocc.bcc_erh = bcc_erh;
            prevUnocc.tcf = tcfA;
            prevUnocc.stRatio = stRatioA;
            prevUnocc.invEcf = (effCfA !== 0) ? (1 / effCfA) : 0;
            prevUnocc.stageLevel = dblStageLevel;
            prevUnocc.pcf = pcfA;
            prevUnocc.ocf = ocfA;
            prevUnocc.demandKw = demandKw;
            binsUnocc.set(odb, prevUnocc);

            if (demandKw > peakDemand) peakDemand = demandKw;
          }
        } finally {
          objSD.FanControls = originalFanControls;
        }
      }

      const annualTotal = annualCondenser + annualEFan + annualAux;
      return {
        annualCondenser,
        annualEFan,
        annualAux,
        auxHoursOcc,
        auxHoursUnocc,
        annualTotal,
        peakDemand,
        totalCoolingHours,
        sample,
        annualCondenser_Occ,
        annualCondenser_Unocc,
        annualEFan_Occ,
        annualEFan_Unocc,
        binsOcc,
        binsUnocc,
        econHours_Occ,
        econHours_Unocc,
        binsEconOcc,
        binsEconUnocc,
        integratedDebugOcc,
        integratedDebugUnocc,
      };
    }

    const econNote =
      'Economizer is enabled per checkbox (Phase 2). Candidate: ' +
      (objSD_C.Economizer === 'on' ? 'ON' : 'OFF') +
      ', Standard: ' +
      (objSD_S.Economizer === 'on' ? 'ON' : 'OFF') +
      '.';
    const resC = runOneSystem(objSD_C);
    const resS = runOneSystem(objSD_S);

    function _renderBinTable(title, binsMap) {
      const rows = Array.from(binsMap.values()).sort((a, b) => a.odb - b.odb);
      let html = '';
      html += '<table class="graph" style="width:auto; margin-top: 1em">';
      html += '<tr><td colspan="4"><b>' + _escapeHtml(title) + '</b></td></tr>';
      html += '<tr><td><b>ODB</b></td><td><b>Hrs</b></td><td><b>E_Fan</b></td><td><b>E_Cond</b></td></tr>';
      for (const r of rows) {
        html += '<tr>' +
          '<td>' + _escapeHtml(Number(r.odb).toFixed(0)) + '</td>' +
          '<td>' + _escapeHtml(Number(r.hours || 0).toFixed(0)) + '</td>' +
          '<td>' + _escapeHtml(Number(r.eFan || 0).toFixed(1)) + '</td>' +
          '<td>' + _escapeHtml(Number(r.eCond || 0).toFixed(1)) + '</td>' +
        '</tr>';
      }
      html += '</table>';
      return html;
    }

    function _renderEconTable(title, binsMap) {
      const rows = Array.from(binsMap.values()).sort((a, b) => a.odb - b.odb);
      let html = '';
      html += '<table class="graph" style="width:auto; margin-top: 1em">';
      html += '<tr><td colspan="5"><b>' + _escapeHtml(title) + '</b></td></tr>';
      html += '<tr><td><b>ODB</b></td><td><b>Hrs</b></td><td><b>Econ Hrs</b></td><td><b>EconLoad</b></td><td><b>IntRT</b></td></tr>';
      for (const r of rows) {
        html += '<tr>' +
          '<td>' + _escapeHtml(Number(r.odb).toFixed(0)) + '</td>' +
          '<td>' + _escapeHtml(Number(r.hours || 0).toFixed(0)) + '</td>' +
          '<td>' + _escapeHtml(Number(r.econHours || 0).toFixed(0)) + '</td>' +
          '<td>' + _escapeHtml(Number(r.econLoad || 0).toFixed(2)) + '</td>' +
          '<td>' + _escapeHtml(Number(r.integrated || 0).toFixed(2)) + '</td>' +
        '</tr>';
      }
      html += '</table>';
      return html;
    }

    const unitsFactor = (Number.isFinite(nUnits) && nUnits > 0) ? nUnits : 1;
    const resC_scaled = {
      annualCondenser: resC.annualCondenser * unitsFactor,
      annualEFan: resC.annualEFan * unitsFactor,
      annualAux: resC.annualAux * unitsFactor,
      annualTotal: resC.annualTotal * unitsFactor,
      peakDemand: resC.peakDemand,
    };
    const resS_scaled = {
      annualCondenser: resS.annualCondenser * unitsFactor,
      annualEFan: resS.annualEFan * unitsFactor,
      annualAux: resS.annualAux * unitsFactor,
      annualTotal: resS.annualTotal * unitsFactor,
      peakDemand: resS.peakDemand,
    };

    _lastRunPhase1 = {
      objSD_C,
      objSD_S,
      resC,
      resS,
      ll,
      nUnits: unitsFactor,
      resC_scaled,
      resS_scaled,
      bpfC: bpfRes_C.bpf,
      bpfS: bpfRes_S.bpf,
      econ: {
        electricityRate,
        discountRate,
        equipmentLife,
        chartPW,
        demandMonths,
        demandCostPerKW,
      },
    };

    const savings_scaled = {
      annualCondenser: resS_scaled.annualCondenser - resC_scaled.annualCondenser,
      annualEFan: resS_scaled.annualEFan - resC_scaled.annualEFan,
      annualAux: resS_scaled.annualAux - resC_scaled.annualAux,
      annualTotal: resS_scaled.annualTotal - resC_scaled.annualTotal,
    };

    return (
      '<div class="graph">' +
      '<h2>Client-side engine (Phase 2: staged, economizer ON/OFF)</h2>' +
      '<p>' + _escapeHtml(econNote) + '</p>' +

      '<table class="graph" style="width:auto">' +
      '<tr><td colspan="3"><b>Annual Energy (kWh)</b></td></tr>' +
      '<tr><td>Units</td><td colspan="2">' + _escapeHtml(unitsFactor) + '</td></tr>' +
      '<tr><td></td><td>Candidate</td><td>Standard</td><td>Savings</td></tr>' +
      '<tr><td>Condenser</td><td>' + _escapeHtml(resC_scaled.annualCondenser.toFixed(0)) + '</td><td>' + _escapeHtml(resS_scaled.annualCondenser.toFixed(0)) + '</td><td>' + _escapeHtml(savings_scaled.annualCondenser.toFixed(0)) + '</td></tr>' +
      '<tr><td>Evap Fan</td><td>' + _escapeHtml(resC_scaled.annualEFan.toFixed(0)) + '</td><td>' + _escapeHtml(resS_scaled.annualEFan.toFixed(0)) + '</td><td>' + _escapeHtml(savings_scaled.annualEFan.toFixed(0)) + '</td></tr>' +
      '<tr><td>Aux</td><td>' + _escapeHtml(resC_scaled.annualAux.toFixed(0)) + '</td><td>' + _escapeHtml(resS_scaled.annualAux.toFixed(0)) + '</td><td>' + _escapeHtml(savings_scaled.annualAux.toFixed(0)) + '</td></tr>' +
      '<tr><td><b>Total</b></td><td><b>' + _escapeHtml(resC_scaled.annualTotal.toFixed(0)) + '</b></td><td><b>' + _escapeHtml(resS_scaled.annualTotal.toFixed(0)) + '</b></td><td><b>' + _escapeHtml(savings_scaled.annualTotal.toFixed(0)) + '</b></td></tr>' +
      '<tr><td colspan="3"><hr/></td></tr>' +
      '<tr><td>Peak Demand (kW)</td><td>' + _escapeHtml(resC.peakDemand.toFixed(2)) + '</td><td>' + _escapeHtml(resS.peakDemand.toFixed(2)) + '</td></tr>' +
      '<tr><td>Total Cooling Hours</td><td>' + _escapeHtml(resC.totalCoolingHours.toFixed(0)) + '</td><td>' + _escapeHtml(resS.totalCoolingHours.toFixed(0)) + '</td></tr>' +
      '<tr><td>FanControls</td><td>' + _escapeHtml(String(objSD_C.FanControls || '')) + '</td><td>' + _escapeHtml(String(objSD_S.FanControls || '')) + '</td></tr>' +
      '<tr><td>EER (input)</td><td>' + _escapeHtml(Number(objSD_C.EER || 0).toFixed(2)) + '</td><td>' + _escapeHtml(Number(objSD_S.EER || 0).toFixed(2)) + '</td></tr>' +
      '<tr><td>BFn_kW (input)</td><td>' + _escapeHtml(Number(objSD_C.BFn_kw || 0).toFixed(3)) + '</td><td>' + _escapeHtml(Number(objSD_S.BFn_kw || 0).toFixed(3)) + '</td></tr>' +
      '<tr><td>Cond_kW (from EER)</td><td>' + _escapeHtml(((totalCap / Math.max(0.001, Number(objSD_C.EER || 0))) - (Number(objSD_C.BFn_kw || 0) + Number(objSD_C.Aux_kw || 0))).toFixed(3)) + '</td><td>' + _escapeHtml(((totalCap / Math.max(0.001, Number(objSD_S.EER || 0))) - (Number(objSD_S.BFn_kw || 0) + Number(objSD_S.Aux_kw || 0))).toFixed(3)) + '</td></tr>' +
      '<tr><td>Cond_kW (input)</td><td>' + _escapeHtml(Number(objSD_C.Cond_kw || 0).toFixed(3)) + '</td><td>' + _escapeHtml(Number(objSD_S.Cond_kw || 0).toFixed(3)) + '</td></tr>' +
      '<tr><td colspan="3"><hr/></td></tr>' +
      '<tr><td>Condenser kWh (Occ)</td><td>' + _escapeHtml(Number(resC.annualCondenser_Occ || 0).toFixed(0)) + '</td><td>' + _escapeHtml(Number(resS.annualCondenser_Occ || 0).toFixed(0)) + '</td></tr>' +
      '<tr><td>Condenser kWh (Unocc)</td><td>' + _escapeHtml(Number(resC.annualCondenser_Unocc || 0).toFixed(0)) + '</td><td>' + _escapeHtml(Number(resS.annualCondenser_Unocc || 0).toFixed(0)) + '</td></tr>' +
      '<tr><td>Evap Fan kWh (Occ)</td><td>' + _escapeHtml(Number(resC.annualEFan_Occ || 0).toFixed(0)) + '</td><td>' + _escapeHtml(Number(resS.annualEFan_Occ || 0).toFixed(0)) + '</td></tr>' +
      '<tr><td>Evap Fan kWh (Unocc)</td><td>' + _escapeHtml(Number(resC.annualEFan_Unocc || 0).toFixed(0)) + '</td><td>' + _escapeHtml(Number(resS.annualEFan_Unocc || 0).toFixed(0)) + '</td></tr>' +
      '</table>' +

      '<div style="display:flex; gap:16px; flex-wrap:wrap">' +
      '<div>' + _renderBinTable('Candidate Bin Energy (Occupied)', resC.binsOcc || new Map()) + _renderBinTable('Candidate Bin Energy (Unoccupied)', resC.binsUnocc || new Map()) + _renderEconTable('Candidate Economizer (Occupied)', resC.binsEconOcc || new Map()) + _renderEconTable('Candidate Economizer (Unoccupied)', resC.binsEconUnocc || new Map()) + '</div>' +
      '<div>' + _renderBinTable('Standard Bin Energy (Occupied)', resS.binsOcc || new Map()) + _renderBinTable('Standard Bin Energy (Unoccupied)', resS.binsUnocc || new Map()) + _renderEconTable('Standard Economizer (Occupied)', resS.binsEconOcc || new Map()) + _renderEconTable('Standard Economizer (Unoccupied)', resS.binsEconUnocc || new Map()) + '</div>' +
      '</div>' +

      '<table class="graph" style="width:auto; margin-top: 1em">' +
      '<tr><td colspan="2"><b>Load Line</b></td></tr>' +
      '<tr><td>Slope</td><td>' + _escapeHtml(ll.slope) + '</td></tr>' +
      '<tr><td>Intercept</td><td>' + _escapeHtml(ll.intercept) + '</td></tr>' +
      '<tr><td colspan="2"><hr/></td></tr>' +
      '<tr><td>BFn kW</td><td>' + _escapeHtml(bfnKwForA0) + ' (' + _escapeHtml(bfnPick.source) + ')</td></tr>' +
      '<tr><td>ST@test</td><td>' + _escapeHtml(stAtTest) + ' (' + _escapeHtml(stPick.source) + ')</td></tr>' +
      '<tr><td>N affinity</td><td>' + _escapeHtml(nAffinity) + ' (' + _escapeHtml(nPick.source) + ')</td></tr>' +
      '<tr><td>Curves</td><td>' + _escapeHtml(objSD_C.DOE2_Curves) + ' (' + _escapeHtml(doe2Source) + ')</td></tr>' +
      '<tr><td>Specific RTU</td><td>' + _escapeHtml(objSD_C.Specific_RTU) + ' (' + _escapeHtml(specificRtuSource) + ')</td></tr>' +
      '<tr><td>TrackOHR</td><td>' + _escapeHtml(trackOhr) + '</td></tr>' +
      '<tr><td>IRH %</td><td>' + _escapeHtml(irhPct) + ' (' + _escapeHtml(irhPick.source) + ')</td></tr>' +
      '<tr><td>Vent Units</td><td>' + _escapeHtml(ventUnitsLabel) + '</td></tr>' +
      '<tr><td>Vent Value</td><td>' + _escapeHtml(ventilationValueRaw) + ' (' + _escapeHtml(ventValuePick.source) + ')</td></tr>' +
      '<tr><td>S&I frac</td><td>' + _escapeHtml(sandIFraction) + ' (' + _escapeHtml(sandIPick.source) + ')</td></tr>' +
      '<tr><td>Oversize %</td><td>' + _escapeHtml(oversizePct) + ' (' + _escapeHtml(oversizePick.source) + ')</td></tr>' +
      '</table>' +

      '</div>'
    );
  } catch (e) {
    return '<div class="graph"><h2>Client-side engine error</h2><pre>' + _escapeHtml(e && (e.stack || e.message || e)) + '</pre></div>';
  }
}
