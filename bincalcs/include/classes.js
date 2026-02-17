export class StagePair {
  constructor() {
    this.A = new StageState();
    this.A.StagePairType = "A";

    this.BmA = new StageState();
    this.BmA.StagePairType = "BmA";

    this.B = new StageState();
    this.B.StagePairType = "B";

    this.pair_mode = "A_only";

    this.FlowFraction_CompOff = 0;
    this.FlowFraction_CompOff_EconOff = 0;
    this.EconomizerRunning = false;

    this.IntegratedState = "";
  }

  SetCondType(objSD) {
    if (typeof objSD?.FanControls === "string" && objSD.FanControls.slice(0, 1) === "V") {
      this.A.CondType = "VC";
      this.BmA.CondType = "NA";
      this.B.CondType = "NA";
    } else {
      this.A.CondType = "Staged";
      this.BmA.CondType = "Staged";
      this.B.CondType = "Staged";
    }
  }
}

export class StageState {
  constructor() {
    this.CondType = undefined;
    this.Level = "NA";
    this.CapacityFraction = "NA";
    this.CapacityFraction_Diff = "NA";
    this.FlowFraction = 0.0;
    this.SensCap_KBtuH = "NA";
    this.LoadFraction = "NA";
    this.RunTime = 0.0;
    this.StagePairType = undefined;
    this.ST_Ratio = "NA";
    this.SDB = "NA";
    this.TCap_CF = "NA";
    this.CondPower_CF = "NA";
    this.Efficiency_CF = "NA";
  }

  ResetToZeroLoad() {
    this.CapacityFraction = 0.0;
    this.FlowFraction = 0.0;
    this.LoadFraction = 0.0;
  }

  ResetToFullLoad() {
    this.CapacityFraction = 1.0;
    this.FlowFraction = 1.0;
    this.LoadFraction = 1.0;
  }
}

export class EnteringConditions {
  constructor() {
    this.ODB = undefined;
    this.EWB = undefined;
    this.EDB = undefined;
    this.EHR = undefined;
  }
}

export class OutdoorConditions {
  constructor() {
    this.DB = undefined;
    this.WB = undefined;
    this.HR = undefined;
    this.BP = undefined;
  }
}

export class SpreadSheetData {
  constructor() {
    this.AirFlow = undefined;
    this.GrossCoolingCapacity = undefined;
    this.NetCoolingCapacity = undefined;
    this.ARI_RatedAirFlow = undefined;
    this.EvaporatorFanPower = undefined;

    this.CondenserPower = undefined;
    this.AuxilaryPower = undefined;
    this.TotalSystemPower = undefined;
    this.EER = undefined;
    this.IEER = undefined;
    this.ST_Ratio = undefined;

    this.VersionOK = undefined;
  }
}

function _isFiniteNumber(x) {
  return typeof x === "number" && Number.isFinite(x);
}

function _solveLinearSystem(A, b) {
  const n = A.length;
  const M = A.map((row, i) => row.slice().concat([b[i]]));

  for (let col = 0; col < n; col++) {
    let pivotRow = col;
    let pivotAbs = Math.abs(M[col][col]);
    for (let r = col + 1; r < n; r++) {
      const v = Math.abs(M[r][col]);
      if (v > pivotAbs) {
        pivotAbs = v;
        pivotRow = r;
      }
    }

    if (!Number.isFinite(pivotAbs) || pivotAbs === 0) {
      throw new Error("Singular matrix");
    }

    if (pivotRow !== col) {
      const tmp = M[col];
      M[col] = M[pivotRow];
      M[pivotRow] = tmp;
    }

    const pivot = M[col][col];
    for (let c = col; c <= n; c++) {
      M[col][c] /= pivot;
    }

    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const factor = M[r][col];
      if (factor === 0) continue;
      for (let c = col; c <= n; c++) {
        M[r][c] -= factor * M[col][c];
      }
    }
  }

  const x = new Array(n);
  for (let i = 0; i < n; i++) x[i] = M[i][n];
  return x;
}

function _evalTerm(term, xVals) {
  let result = 1;
  const factors = term.split("*").map((s) => s.trim()).filter(Boolean);

  for (const factor of factors) {
    if (factor === "1") {
      continue;
    }
    const m = factor.match(/^X(\d+)(?:\^(\d+))?$/);
    if (!m) {
      throw new Error(`Unsupported factor: ${factor}`);
    }
    const idx = parseInt(m[1], 10);
    const pow = m[2] ? parseInt(m[2], 10) : 1;
    const base = xVals[`X${idx}`];
    result *= Math.pow(base, pow);
  }
  return result;
}

class LeastSquaresModel {
  constructor() {
    this.IsSolved = false;
    this.LastSolveError = "";
    this._terms = [];
    this._coefficients = [];
    this._tValues = [];
    this._r2 = null;
    this._residualSE = null;
  }

  fitFromArray(dblDataArray, modelString) {
    try {
      const terms = modelString
        .split("+")
        .map((s) => s.trim())
        .filter(Boolean);

      if (terms.length === 0) {
        throw new Error("No model terms");
      }

      const X = [];
      const Y = [];

      for (const row of dblDataArray) {
        if (!Array.isArray(row) || row.length < 2) continue;

        const y = Number(row[0]);
        if (!_isFiniteNumber(y)) continue;

        const xVals = {};
        for (let j = 1; j < row.length; j++) {
          const v = Number(row[j]);
          xVals[`X${j}`] = v;
        }

        const designRow = new Array(terms.length);
        let ok = true;
        for (let t = 0; t < terms.length; t++) {
          const val = _evalTerm(terms[t], xVals);
          if (!_isFiniteNumber(val)) {
            ok = false;
            break;
          }
          designRow[t] = val;
        }
        if (!ok) continue;

        X.push(designRow);
        Y.push(y);
      }

      if (X.length < terms.length) {
        throw new Error("Insufficient data points");
      }

      const p = terms.length;
      const XtX = Array.from({ length: p }, () => Array.from({ length: p }, () => 0));
      const XtY = Array.from({ length: p }, () => 0);

      for (let i = 0; i < X.length; i++) {
        const xi = X[i];
        const yi = Y[i];
        for (let a = 0; a < p; a++) {
          XtY[a] += xi[a] * yi;
          for (let b = 0; b < p; b++) {
            XtX[a][b] += xi[a] * xi[b];
          }
        }
      }

      // Save XtX before solve (Gaussian elimination mutates the matrix in-place).
      const XtX_copy = XtX.map(r => r.slice());
      const beta = _solveLinearSystem(XtX, XtY);
      this._terms = terms;
      this._coefficients = beta;
      this.IsSolved = true;
      this.LastSolveError = "";

      // Compute R², residual SE, and t-values (matches ASP COM regression output).
      const n = X.length;
      let yMean = 0;
      for (let i = 0; i < n; i++) yMean += Y[i];
      yMean /= n;

      let SSE = 0, SST = 0;
      for (let i = 0; i < n; i++) {
        let yHat = 0;
        for (let j = 0; j < p; j++) yHat += beta[j] * X[i][j];
        SSE += (Y[i] - yHat) * (Y[i] - yHat);
        SST += (Y[i] - yMean) * (Y[i] - yMean);
      }
      this._r2 = SST > 0 ? 1 - SSE / SST : null;
      const dof = n - p;
      const MSE = dof > 0 ? SSE / dof : 0;
      this._residualSE = dof > 0 ? Math.sqrt(MSE) : null;

      // Invert (X'X) to get covariance matrix diagonal for coefficient SEs.
      // Use _solveLinearSystem with identity columns.
      const tVals = new Array(p);
      try {
        for (let j = 0; j < p; j++) {
          const ej = Array.from({ length: p }, (_, k) => k === j ? 1 : 0);
          const col = _solveLinearSystem(XtX_copy.map(r => r.slice()), ej);
          const seBeta = Math.sqrt(Math.max(0, MSE * col[j]));
          tVals[j] = seBeta > 0 ? beta[j] / seBeta : null;
        }
      } catch {
        tVals.fill(null);
      }
      this._tValues = tVals;

      return true;
    } catch (e) {
      this.IsSolved = false;
      this.LastSolveError = e instanceof Error ? e.message : String(e);
      this._terms = [];
      this._coefficients = [];
      this._tValues = [];
      this._r2 = null;
      this._residualSE = null;
      return false;
    }
  }

  PredictValueAlt(xArray) {
    if (!this.IsSolved) {
      throw new Error(this.LastSolveError || "Model not solved");
    }

    const xVals = {};
    for (let i = 0; i < xArray.length; i++) {
      xVals[`X${i + 1}`] = Number(xArray[i]);
    }

    let y = 0;
    for (let i = 0; i < this._terms.length; i++) {
      y += this._coefficients[i] * _evalTerm(this._terms[i], xVals);
    }
    return y;
  }
}

export class SystemProperties {
  constructor() {
    this.SystemName = undefined;
    this.EER = undefined;
    this.UnitCost = undefined;
    this.Maintenance = undefined;
    this.Economizer = undefined;
    this.Spreadsheet = undefined;

    this.BFn_kw = undefined;
    this.Aux_kw = undefined;
    this.Cond_kw = undefined;
    this.BPF = undefined;
    this.A0_BPF = undefined;

    this.Specific_RTU = undefined;
    this.N_Affinity = undefined;
    this.CapacityFraction_Min = undefined;

    this.CondFanPercent = undefined;
    this.N_Stages = undefined;
    this.StageLevels = undefined;
    this.DOE2_Curves = undefined;

    this.FanControls = undefined;
    this.PLDegrFactor = undefined;

    this.ST_Ratio_AtTest = undefined;

    this.SSD = new SpreadSheetData();
    this.SSD.VersionOK = true;

    this.AnnualCost = undefined;

    this.GrossCapacity_KBtu_CA = undefined;
    this.Condenser_kW_CA = undefined;
    this.ST_Ratio_CA = undefined;
    this.EER_PL_CA = undefined;

    this.GrossCapacity_KBtu_Model = new LeastSquaresModel();
    this.Condenser_kW_Model = new LeastSquaresModel();
    this.ST_Ratio_Model = new LeastSquaresModel();
    this.NEER_PL_Model = new LeastSquaresModel();
  }

  Fit_Model(strNameArray, strModelString, strModelDesc, dblDataArray, dfxModel) {
    if (!(dfxModel instanceof LeastSquaresModel)) {
      throw new Error("Fit_Model expects a LeastSquaresModel");
    }
    const ok = dfxModel.fitFromArray(dblDataArray, strModelString);
    if (!ok) {
      return dfxModel.LastSolveError || "Model solve failed";
    }
    return "";
  }

  ParseSpreadsheetData() {
    const text = this.Spreadsheet;
    if (typeof text !== "string" || text === "") {
      return [];
    }
    return text.split(/\r\n|\n|\r/);
  }

  ParseAndModel(strMode) {
    const strVersion_requirement = "V1.2";
    const rows = this.ParseSpreadsheetData();

    this.SSD.VersionOK = true;

    if (!(rows.length > 1 && rows.length < 40)) {
      this.SSD.VersionOK = false;
      return {
        errorMessage: "Spreadsheet data is not in the correct form.",
      };
    }

    const strVersion = rows[1] ?? "";
    if (!String(strVersion).includes(strVersion_requirement)) {
      this.SSD.VersionOK = false;
      return {
        errorMessage: `Spreadsheet version is not ${strVersion_requirement}.`,
      };
    }

    const GrossCapacity_KBtu = Array.from({ length: 18 }, () => ["NA", "NA", "NA"]);
    const Condenser_kW = Array.from({ length: 18 }, () => ["NA", "NA", "NA"]);
    const ST_Ratio = Array.from({ length: 54 }, () => ["NA", "NA", "NA", "NA"]);
    const EER_PL = Array.from({ length: 16 }, () => ["NA", "NA", "NA"]);
    const EER_PL_Normalized = Array.from({ length: 16 }, () => ["NA", "NA", "NA"]);

    let I_OE = 0;
    let I_partload = 0;

    for (let j = 2; j <= rows.length - 2; j++) {
      const row = rows[j];
      const cells = String(row).split("\t");
      const key = String(cells[0] ?? '').trim();

      if (key === "Cap/Cond/ST") {
        const temps = String(cells[1] ?? "").split("/");
        const dblODB = Number(temps[0]);
        const dblEWB = Number(temps[1]);

        GrossCapacity_KBtu[I_OE][0] = cells[2];
        GrossCapacity_KBtu[I_OE][1] = dblODB;
        GrossCapacity_KBtu[I_OE][2] = dblEWB;

        Condenser_kW[I_OE][0] = cells[3];
        Condenser_kW[I_OE][1] = dblODB;
        Condenser_kW[I_OE][2] = dblEWB;

        ST_Ratio[3 * I_OE + 0][0] = cells[4];
        ST_Ratio[3 * I_OE + 0][1] = dblODB;
        ST_Ratio[3 * I_OE + 0][2] = dblEWB;
        ST_Ratio[3 * I_OE + 0][3] = 75.0;

        ST_Ratio[3 * I_OE + 1][0] = cells[5];
        ST_Ratio[3 * I_OE + 1][1] = dblODB;
        ST_Ratio[3 * I_OE + 1][2] = dblEWB;
        ST_Ratio[3 * I_OE + 1][3] = 80.0;

        ST_Ratio[3 * I_OE + 2][0] = cells[6];
        ST_Ratio[3 * I_OE + 2][1] = dblODB;
        ST_Ratio[3 * I_OE + 2][2] = dblEWB;
        ST_Ratio[3 * I_OE + 2][3] = 85.0;

        I_OE += 1;
      } else if (key === "Partload") {
        const loadPct = cells[1];

        EER_PL[4 * I_partload + 0][0] = cells[2];
        EER_PL[4 * I_partload + 0][1] = loadPct;
        EER_PL[4 * I_partload + 0][2] = 95;

        EER_PL[4 * I_partload + 1][0] = cells[3];
        EER_PL[4 * I_partload + 1][1] = loadPct;
        EER_PL[4 * I_partload + 1][2] = 81.5;

        EER_PL[4 * I_partload + 2][0] = cells[4];
        EER_PL[4 * I_partload + 2][1] = loadPct;
        EER_PL[4 * I_partload + 2][2] = 68;

        EER_PL[4 * I_partload + 3][0] = cells[5];
        EER_PL[4 * I_partload + 3][1] = loadPct;
        EER_PL[4 * I_partload + 3][2] = 65;

        I_partload += 1;
      } else if (key === "AirFlow") {
        this.SSD.AirFlow = cells[2];
      } else if (key === "GrossCoolingCapacity") {
        this.SSD.GrossCoolingCapacity = cells[2];
      } else if (key === "NetCoolingCapacity") {
        this.SSD.NetCoolingCapacity = cells[2];
      } else if (key === "ARI_RatedAirFlow") {
        this.SSD.ARI_RatedAirFlow = cells[2];
      } else if (key === "EvaporatorFanPower") {
        this.SSD.EvaporatorFanPower = cells[2];
      } else if (key === "CondenserPower" || key === "CondensorPower") {
        this.SSD.CondenserPower = cells[2];
      } else if (key === "AuxilaryPower" || key === "AuxiliaryPower") {
        this.SSD.AuxilaryPower = cells[2];
      } else if (key === "TotalSystemPower") {
        this.SSD.TotalSystemPower = cells[2];
      } else if (key === "EER") {
        this.SSD.EER = cells[2];
      } else if (key === "IEER") {
        this.SSD.IEER = cells[2];
      } else if (key === "ST_Ratio") {
        this.SSD.ST_Ratio = cells[2];
      }
    }

    if (strMode === "parseOnly") {
      return {
        errorMessage: "",
      };
    }

    const errors = [];

    if (I_OE > 1) {
      const capClean = [];
      for (let j = 0; j < GrossCapacity_KBtu.length; j++) {
        if (GrossCapacity_KBtu[j][0] !== "NA") {
          capClean.push([
            Number(GrossCapacity_KBtu[j][0]),
            Number(GrossCapacity_KBtu[j][1]),
            Number(GrossCapacity_KBtu[j][2]),
          ]);
        }
      }
      this.GrossCapacity_KBtu_CA = capClean;
      if (capClean.length > 2) {
        const err = this.Fit_Model(
          ["GrossCapacity", "ODB", "EWB"],
          "X1 + X1^2 + X2^2 + X1*X2 + X1^2*X2^2",
          "Gross Capacity",
          capClean,
          this.GrossCapacity_KBtu_Model
        );
        if (err) errors.push(err);
      }

      const condClean = [];
      for (let j = 0; j < Condenser_kW.length; j++) {
        if (Condenser_kW[j][0] !== "NA") {
          condClean.push([
            Number(Condenser_kW[j][0]),
            Number(Condenser_kW[j][1]),
            Number(Condenser_kW[j][2]),
          ]);
        }
      }
      this.Condenser_kW_CA = condClean;
      if (condClean.length > 2) {
        const err = this.Fit_Model(
          ["Cond_kW", "ODB", "EWB"],
          "X1^2 + X2^2 + X1*X2 + X1^2*X2^2 + X1^3*X2^3",
          "Condenser Power Model",
          condClean,
          this.Condenser_kW_Model
        );
        if (err) errors.push(err);
      }

      const strClean = [];
      for (let j = 0; j < ST_Ratio.length; j++) {
        const v = ST_Ratio[j][0];
        if (v !== "NA") {
          const vv = Number(v);
          if (Number.isFinite(vv) && vv < 1) {
            strClean.push([
              vv,
              Number(ST_Ratio[j][1]),
              Number(ST_Ratio[j][2]),
              Number(ST_Ratio[j][3]),
            ]);
          }
        }
      }
      this.ST_Ratio_CA = strClean;
      if (strClean.length > 2) {
        const err = this.Fit_Model(
          ["STratio", "ODB", "EWB", "EDB"],
          "X2 + X3 + X2*X3 + X2^2*X3 + X1*X2*X3^2 + X1*X2^2*X3",
          "S/T Ratio",
          strClean,
          this.ST_Ratio_Model
        );
        if (err) errors.push(err);
      }
    }

    if (I_partload > 1) {
      for (let j = 0; j < EER_PL.length; j++) {
        EER_PL_Normalized[j][1] = EER_PL[j][1];
        EER_PL_Normalized[j][2] = EER_PL[j][2];
      }

      for (let j = 0; j <= 3; j++) {
        for (let k = 0; k <= 3; k++) {
          const idx = j + k * 4;
          if (EER_PL[idx][0] !== "NA") {
            if (EER_PL[j][0] !== "NA") {
              const denom = Number(EER_PL[j][0]);
              const num = Number(EER_PL[idx][0]);
              if (Number.isFinite(denom) && denom !== 0 && Number.isFinite(num)) {
                EER_PL_Normalized[idx][0] = num / denom;
              } else {
                EER_PL_Normalized[idx][0] = "NA";
              }
            } else {
              EER_PL_Normalized[idx][0] = "NA";
            }
          } else {
            EER_PL_Normalized[idx][0] = "NA";
          }
        }
      }

      const eerClean = [];
      for (let j = 0; j < EER_PL_Normalized.length; j++) {
        if (EER_PL_Normalized[j][0] !== "NA") {
          eerClean.push([
            Number(EER_PL_Normalized[j][0]),
            Number(EER_PL_Normalized[j][1]),
            Number(EER_PL_Normalized[j][2]),
          ]);
        }
      }
      this.EER_PL_CA = eerClean;
      if (eerClean.length > 2) {
        const err = this.Fit_Model(
          ["NEER", "LF", "ODB"],
          "1 + X1 + X1^2 + X1*X2 + X1^2*X2 + X1^3",
          "Normalized EER",
          eerClean,
          this.NEER_PL_Model
        );
        if (err) errors.push(err);
      }
    }

    return {
      errorMessage: errors.length ? errors.join("; ") : "",
    };
  }

  CompactTheArray(dblInputArray, intGoodPoints, intDimensions) {
    const out = new Array(intGoodPoints);
    for (let j = 0; j < intGoodPoints; j++) {
      out[j] = new Array(intDimensions + 1);
      for (let k = 0; k <= intDimensions; k++) {
        out[j][k] = dblInputArray[j][k];
      }
    }
    return out;
  }

  Predict_ST_Ratio(dblODB, dblEWB, dblEDB) {
    const dblRawPrediction = this.ST_Ratio_Model.PredictValueAlt([dblODB, dblEWB, dblEDB]);

    if (dblRawPrediction > 1) return 1;
    if (dblRawPrediction < 0) return 0;
    return dblRawPrediction;
  }

  Predict_Gross_Capacity_Correction(dblODB, dblEWB) {
    return (
      this.GrossCapacity_KBtu_Model.PredictValueAlt([dblODB, dblEWB]) /
      this.GrossCapacity_KBtu_Model.PredictValueAlt([95, 67])
    );
  }

  Predict_Condenser_Correction(dblODB, dblEWB) {
    return (
      this.Condenser_kW_Model.PredictValueAlt([dblODB, dblEWB]) /
      this.Condenser_kW_Model.PredictValueAlt([95, 67])
    );
  }

  Predict_PartloadFactor(dblLoadPercent, dblODB) {
    let load = dblLoadPercent;
    if (load > 100.0) load = 100.0;

    const dblRawPrediction = this.NEER_PL_Model.PredictValueAlt([load, dblODB]);

    if (dblRawPrediction < 0) return 0;
    return dblRawPrediction;
  }
}

export class TwoDDictionary {
  constructor() {
    this._binArray = Array.from({ length: 100 }, () => Array.from({ length: 41 }, () => "NA"));

    this._col = new Map();
    this._row = new Map();

    this._firstBin = 0;
    this._lastBin = 0;

    this._minRow = 0;
    this._maxRow = 0;
    this._count = 0;

    this.Energy_Annual_BlowerFan = undefined;
    this.Energy_Annual_Aux = undefined;
    this.Energy_Annual_Condenser = undefined;
    this.Energy_Annual_Total = undefined;
    this.SystemName = undefined;
    this.Occupied = undefined;
    this.Peak_Demand_kW = undefined;
    this.Demand_Cost = undefined;
  }

  get Count() {
    return this._count;
  }

  DefineColumnNames(strColNames) {
    this._col = new Map();

    let j = 0;
    for (const name of strColNames) {
      this._col.set(name, j);
      j += 1;
    }

    for (let c = 0; c < this._col.size; c++) {
      this._binArray[0][c] = 0;
      this._binArray[1][c] = 0;
    }
  }

  IntializeBinRange(dblODB_low, dblODB_high) {
    this._row = new Map();

    this._firstBin = dblODB_low;
    this._lastBin = dblODB_high;

    let step = 5;
    if (dblODB_low > dblODB_high) step = -5;

    let r = 2;
    for (let bin = dblODB_low; step > 0 ? bin <= dblODB_high : bin >= dblODB_high; bin += step) {
      this._row.set(bin, r);
      r += 1;
    }
  }

  _rowNameInRange(dblRowName) {
    const minBin = Math.min(this._firstBin, this._lastBin);
    const maxBin = Math.max(this._firstBin, this._lastBin);
    return dblRowName >= minBin && dblRowName <= maxBin;
  }

  setBinValue(strColumnName, strRowName, dblValue_input) {
    if (!this._col.has(strColumnName) || !this._row.has(strRowName)) {
      return;
    }

    if (!this._rowNameInRange(strRowName)) {
      return;
    }

    const col = this._col.get(strColumnName);
    const row = this._row.get(strRowName);

    this._binArray[row][col] = dblValue_input;

    const firstRow = this._binArray[0][col];
    const lastRow = this._binArray[1][col];

    if (firstRow === 0 || row < firstRow) this._binArray[0][col] = row;
    if (lastRow === 0 || row > lastRow) this._binArray[1][col] = row;

    if (this._minRow === 0 || row < this._minRow) this._minRow = row;
    if (this._maxRow === 0 || row > this._maxRow) this._maxRow = row;

    this._count = this._maxRow - this._minRow + 1;
  }

  getBinValue(strColumnName, strRowName) {
    if (!this._col.has(strColumnName) || !this._row.has(strRowName)) {
      return "NA";
    }

    return this._binArray[this._row.get(strRowName)][this._col.get(strColumnName)];
  }

  GetColumn(strColumnName) {
    if (!this._col.has(strColumnName)) return [];

    const col = this._col.get(strColumnName);
    const firstRow = this._binArray[0][col];
    const lastRow = this._binArray[1][col];

    if (!firstRow || !lastRow) return [];

    const out = [];
    for (let r = firstRow; r <= lastRow; r++) {
      out.push(this._binArray[r][col]);
    }
    return out;
  }

  GetColumnAll(strColumnName) {
    if (!this._col.has(strColumnName)) return [];

    const col = this._col.get(strColumnName);

    if (!this._minRow || !this._maxRow) return [];

    const out = [];
    for (let r = this._minRow; r <= this._maxRow; r++) {
      out.push(this._binArray[r][col]);
    }
    return out;
  }
}

export class FormValues {
  constructor() {
    this._values = {};
  }

  static fromObject(values) {
    const fv = new FormValues();
    fv._values = { ...values };
    return fv;
  }

  get Spreadsheet_S() {
    return this._values.Spreadsheet_S;
  }

  get Spreadsheet_C() {
    return this._values.Spreadsheet_C;
  }

  get Advanced() {
    return this._values.Advanced;
  }

  get DemandMonths() {
    return this._values.DemandMonths;
  }

  get DemandCostPerKW() {
    return this._values.DemandCostPerKW;
  }

  get PLDegrFactor_C() {
    return this._values.PLDegrFactor_C;
  }

  get PLDegrFactor_S() {
    return this._values.PLDegrFactor_S;
  }

  get FanControls_C() {
    return this._values.FanControls_C;
  }

  get FanControls_S() {
    return this._values.FanControls_S;
  }

  get BFn_kw_S() {
    return this._values.BFn_kw_S;
  }

  get Aux_kw_S() {
    return this._values.Aux_kw_S;
  }

  get Cond_kw_S() {
    return this._values.Cond_kw_S;
  }

  get BFn_kw_C() {
    return this._values.BFn_kw_C;
  }

  get Aux_kw_C() {
    return this._values.Aux_kw_C;
  }

  get Cond_kw_C() {
    return this._values.Cond_kw_C;
  }

  get CondFanPercent_S() {
    return this._values.CondFanPercent_S;
  }

  get CondFanPercent_C() {
    return this._values.CondFanPercent_C;
  }

  get Schedule() {
    return this._values.Schedule;
  }

  get Economizer_S() {
    return this._values.Economizer_S || "off";
  }

  get Economizer_C() {
    return this._values.Economizer_C || "off";
  }

  get Nstages_S() {
    return this._values.Nstages_S;
  }

  get Nstages_C() {
    return this._values.Nstages_C;
  }

  get UnitCost_Standard() {
    return this._values.UnitCost_Standard;
  }

  get UnitCost_Candidate() {
    return this._values.UnitCost_Candidate;
  }

  get OversizePercent() {
    return this._values.OversizePercent;
  }

  get Oversizing_LoadReductionFactor() {
    return this._values.Oversizing_LoadReductionFactor ?? (this.OversizePercent / 100 + 1);
  }

  get IDB() {
    return this._values.IDB;
  }

  get IDB_SetBack() {
    if (this._values.IDB_SetBack === "Cond. Off") return 0;
    return this._values.IDB_SetBack;
  }

  get Details() {
    return this._values.Details;
  }

  get ChartPW() {
    return this._values.ChartPW;
  }

  get Maintenance_Standard() {
    return this._values.Maintenance_Standard;
  }

  get Maintenance_Candidate() {
    return this._values.Maintenance_Candidate;
  }

  get ElectricityRate() {
    return this._values.ElectricityRate;
  }

  get EER_Standard() {
    return this._values.EER_Standard;
  }

  get EER() {
    return this._values.EER;
  }

  get Manufacturer() {
    return this._values.Manufacturer;
  }

  get Slope() {
    return this._values.Slope;
  }

  get Intercept() {
    return this._values.Intercept;
  }

  get DiscountRate() {
    return this._values.DiscountRate;
  }

  get LockLoadLine() {
    return this._values.LockLoadLine;
  }

  get ST_Ratio_AtTest_C() {
    return this._values.ST_Ratio_AtTest_C;
  }

  get ST_Ratio_AtTest_S() {
    return this._values.ST_Ratio_AtTest_S;
  }

  get SandI_fraction() {
    return this._values.SandI_fraction;
  }

  get EnthalpyControl() {
    return this._values.EnthalpyControl;
  }

  get VentilationUnits() {
    return this._values.VentilationUnits;
  }

  get VentilationValue() {
    return this._values.VentilationValue;
  }

  get VentilationFraction() {
    if (this._values.VentilationUnits === "CFM") {
      const denom = this.CFM;
      let frac = denom ? this._values.VentilationValue / denom : 0;
      if (frac > 1.0) frac = 1.0;
      return frac;
    }
    return (this._values.VentilationValue ?? 0) / 100;
  }

  get TrackOHR() {
    return this._values.TrackOHR;
  }

  get IRH() {
    return this._values.IRH;
  }

  get Units() {
    return this._values.Units;
  }

  get TotalCap() {
    return this._values.TotalCap;
  }

  get CFM() {
    const cfmPerTon = this._values.CFMperTon ?? 400;
    return (this.TotalCap / 12) * cfmPerTon;
  }

  get State() {
    return this._values.State;
  }

  get CityName2() {
    return this._values.CityName2;
  }

  get EquipmentLife() {
    return this._values.EquipmentLife;
  }

  get Specific_RTU_C() {
    return this._values.Specific_RTU_C;
  }

  get N_Affinity() {
    return this._values.N_Affinity;
  }

  get DOE2_Curves() {
    if (this._values.DOE2_Curves === "on") return "DOE2";
    return "Carrier";
  }
}

export class CellTitles {
  constructor() {
    this.Elv = "Elevation at specified location (feet)";
    this.P = "Standard pressure corrected for elevation (inHg)";
    this.ORH = "Outside relative humidity (%)";
    this.IDB = "Inside dry bulb (F)";
    this.IWB = "Inside wet bulb (F)";

    this.ODB = "Outside dry bulb (F)";
    this.OWB = "Outside mean coincident web bulb (F)";
    this.OHR = "Outside humidity ratio";
    this.IHR = "Inside humidity ratio";
    this.IRH = "Inside relative humidity (%)";
    this.Hrs = "Number of hours that this outside condition occurs during the specified schedule in one year at the specified location";
    this.NVLd = "Non-ventilation sensible load (kBtuh)";
    this.VLd = "Ventilation sensible load (kBtuh)";
    this.Ld = "Sum of sensible loads, before considering economizer (kBtuh)";
    this.ECp = "Economizer sensible capacity: derived from fan capacity [excluding ventilation flow] (kBtuh)";
    this.SLdE = "The sensible load after the economizer is considered (kBtuh)";
    this.LLdE = "The latent load after the economizer is considered (kBtuh)";
    this.EDB = "Entering, mixed air dry bulb (F)";
    this.EHR = "Entering, mixed air humidity ratio";
    this.ERH = "Entering, mixed air relative humidity (%)";
    this.EWB = "Entering, mixed air wet bulb (F)";
    this.TCF = "Correction factor applied to ARI-rated gross total capacity; Used in calculating PCF.";
    this.BPF = "ByPass Factor (BPF) adjusted for mass flow at bin conditions.";
    this.ST = "Sensible to total capacity ratio; Used with TCF in calculating sensible capacity.";
    this.ECF = "Inverse (i.e. Out/In, like an EER) of DOE-2 defined Input Efficiency (In/Out); Used in calculating PCF.";

    this.StageLevel_A = "Stage level: for a staged RTU, 2.9 indicates the unit runs at stage level 3 for 90% of the hour and at level 2 for 10%; ";
    this.StageLevel_B = "for a variable-capacity RTU, 0.73 indicates the units is running at 73% of capacity.";
    this.FF_A = "Flow fraction for stage A of the A-B stage pair.";
    this.FF_B = "Flow fraction for stage B of the A-B stage pair.";
    this.FF_Off = "Flow fraction when all condensers are off AND the evaporator fan is on.";

    this.LF1 = "Load fraction for stage 1 (Sensible Load After Economizing / Net Sensible Capacity)";
    this.CP_CF1 = "Part-load correction factor applied to the full-load condenser power for stage 1.";
    this.RT1 = "Runtime for stage 1 = LF1";
    this.LF2 = "Load fraction for stage 2 (Sensible Load After Economizing / Net Sensible Capacity)";
    this.CP_CF2 = "Part-load correction factor applied to the full-load condenser power for stage 2.";
    this.RT2 = "Runtime for stage 2 = LF2";

    this.PCF = "Correction factor applied to ARI-rated condenser power (= TCF/ECF)";
    this.OCF = "Overall system correction factor, proportional to energy usage ( = PCF / (TCF * S/T_CF) ), where S/T_CF = (S/T_@EnteringConditions) / (S/T_@ARI)";

    this.kW = "Peak wattage (kW)";
    this.E_BFan = "Unit's evaporator-fan energy consumption (kWhrs)";
    this.E_Condenser = "Unit's condenser unit (fan and compressor) energy consumption (kWhrs)";
    this.E_Aux = "Unit's auxiliary (electronics) energy consumption (kWhrs)";
  }
}

export function FormatSciNotation(dblInputNumber, intSignificant) {
  return Number(dblInputNumber).toExponential(intSignificant);
}
