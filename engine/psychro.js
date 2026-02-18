/**
 * Psychrometric routines - JavaScript ES Module
 * Converted from psychro.asp (VBScript)
 * 
 * All functions use IP (Imperial) units:
 * - Temperature: °F
 * - Pressure: inches Hg
 * - Humidity ratio: lb water / lb dry air
 * - Enthalpy: Btu / lb dry air
 * - Specific volume: ft³ / lb dry air
 *
 * Converted from psychro.asp (VBScript).  Used by every other engine module.
 * See ARCHITECTURE.md for how this module fits into the engine pipeline.
 */

// Standard pressure at sea level (inches Hg)
export const dblStandardPressure = 29.921;

// Conversion constant
export const dblKWtoKBTUH = 3.412;

//=========================================================
// Specific Volume Functions
//=========================================================

// Specific volume from dry-bulb, humidity ratio, and barometric pressure.
export function Psv_hr(db, hr, bp) {
    return 53.352 * (db + 459.67) * (1 + 1.6078 * hr) / (bp * 144 / 2.0360);
}

export function Psv_rh(db, rh, bp) {
    return Psv_hr(db, Phr_rh(db, rh, bp), bp);
}

export function Psv_wb(db, wb, bp) {
    return Psv_hr(db, Phr_wb(db, wb, bp), bp);
}

//=========================================================
// Dew Point Functions
//=========================================================

// Dew point from humidity ratio and barometric pressure (ASHRAE correlation).
export function Pdp_hr(hr, bp) {
    const pv_inhg = bp * hr / (0.62198 + hr);
    const pv_psi = pv_inhg / 2.0360;
    const alpha = Math.log(pv_psi);

    const a = 100.45;
    const b = 33.193;
    const c = 2.319;
    const d = 0.17074;
    const e = 1.2063;

    const dp_gt32 = a + (b * alpha) + (c * alpha ** 2) + (d * alpha ** 3) + (e * (pv_psi) ** 0.1984);
    const dp_le32 = 90.12 + (26.142 * alpha) + (0.8927 * alpha ** 2);

    if (dp_gt32 > 32) {
        return dp_gt32;
    } else {
        return dp_le32;
    }
}

export function Pdp_rh(db, rh, bp) {
    const hr = Phr_rh(db, rh, bp);
    return Pdp_hr(hr, bp);
}

export function Pdp_wb(db, wb, bp) {
    const hr = Phr_wb(db, wb, bp);
    return Pdp_hr(hr, bp);
}

//=========================================================
// Humidity Ratio (from enthalpy)
//=========================================================

// Humidity ratio from dry-bulb and enthalpy (inverse of Ph_hr).
export function Phr_h(db, h) {
    return (h - 0.240 * db) / (1061 + 0.444 * db);
}

//=========================================================
// Water vapor pressure at saturation (inches Hg)
//=========================================================

// Saturation vapor pressure (inches Hg) from dry-bulb (ASHRAE Hyland-Wexler).
export function Ppvs(db) {
    const db_R = db + 459.67;

    if (db > 32) {
        const pvs_gt32 = Math.exp(-1.044039708e+04 / db_R + -1.12946496e+01 + -2.7022355e-02 * db_R + 1.2890360e-05 * db_R ** 2 + -2.478068e-09 * db_R ** 3 + 6.5459673 * Math.log(db_R));
        return pvs_gt32 * 2.0360;
    } else {
        const pvs_le32 = Math.exp(-1.021416462e+04 / db_R + -4.89350301 + -5.37657944e-03 * db_R + 1.92023769e-07 * db_R ** 2 + 3.55758316e-10 * db_R ** 3 + -9.03446883e-14 * db_R ** 4 + 4.1635019 * Math.log(db_R));
        return pvs_le32 * 2.0360;
    }
}

//=========================================================
// Humidity Ratio at Saturation
//=========================================================

export function Phrs(db, bp) {
    const pws = Ppvs(db);
    return 0.62198 * pws / (bp - pws);
}

//=========================================================
// Humidity Ratio Functions
//=========================================================

// Humidity ratio from dry-bulb, wet-bulb, and barometric pressure.
export function Phr_wb(db, wb, bp) {
    return ((1093 - 0.556 * wb) * Phrs(wb, bp) - 0.240 * (db - wb)) / (1093 + 0.444 * db - wb);
}

export function Phr_rh(db, rh, bp) {
    const vp = rh * Ppvs(db);
    return 0.62198 * vp / (bp - vp);
}

//=========================================================
// Enthalpy (Btu/lbs dry air)
//=========================================================

// Enthalpy (Btu/lb dry air) from dry-bulb and humidity ratio.
export function Ph_hr(db, hr) {
    return (0.240 * db) + hr * (1061 + 0.444 * db);
}

export function Pdb_hr(h, hr) {
    return (h - hr * 1061) / (0.240 + hr * 0.444);
}

export function Ph_wb(db, wb, bp) {
    return Ph_hr(db, Phr_wb(db, wb, bp));
}

export function Ph_rh(db, rh, bp) {
    return (0.240 * db) + Phr_wb(db, Pwb_rh(db, rh, bp), bp) * (1061 + 0.444 * db);
}

//=========================================================
// Relative Humidity (ranges 0 to 1)
//=========================================================

export function Prh_wb(db, wb, bp) {
    return Prh_hr(db, Phr_wb(db, wb, bp), bp);
}

export function Prh_hr(db, hr, bp) {
    const pv = bp * hr / (0.62198 + hr);
    return pv / Ppvs(db);
}

//=========================================================
// Wet Bulb
//=========================================================

export function Pwb_rh(db, rh, bp) {
    return Pwb_hr(db, Phr_rh(db, rh, bp), bp);
}

export function Pwba_hr(db, hr) {
    const h = 0.24 * db + (1061 + 0.444 * db) * hr;
    const y = Math.log(h);
    const wb1 = 30.9185 - 39.682 * y + 20.5841 * y ** 2 - 1.758 * y ** 3;
    const wb2 = 0.6040 + 3.4841 * y + 1.3601 * y ** 2 + 0.9731 * y ** 3;
    if (h < 11.758) {
        return wb2;
    } else {
        return wb1;
    }
}

// Wet-bulb from dry-bulb, humidity ratio, and barometric pressure
// (Newton iteration via Pwbest).
export function Pwb_hr(db, hr, bp) {
    const hr_target = hr;
    let wb_first_guess;
    
    if (db > 0) {
        wb_first_guess = Pwba_hr(db, hr);
    } else {
        wb_first_guess = db;
    }
    
    const wb_delta = 0.02;
    const ok = (db !== null) && (hr !== null) && (bp !== null) && 
               !isNaN(db) && !isNaN(hr) && !isNaN(bp);

    if (!ok) {
        return null;
    } else {
        return Pwbest(db, hr_target, bp, wb_first_guess, wb_delta, 0);
    }
}

function Pwbest(db, hr_target, bp, wb_guess, wb_delta, iter_count) {
    const hr_guess_delta = Phr_wb(db, wb_guess + wb_delta, bp);
    const hr_guess = Phr_wb(db, wb_guess, bp);
    const slope = (hr_guess_delta - hr_guess) / wb_delta;

    // The x_target point, where y=y_target on the line (y-yi)=m(x=xi), is x_target = xi + (y_target - yi)/m
    const wb_next = wb_guess + (hr_target - hr_guess) / slope;
    const wb_diff = Math.abs(wb_next - wb_guess);
    const needsImproving = wb_diff > 0.001;

    // If estimate needs improving, recursively call this function.
    if (iter_count < 10 && needsImproving) {
        return Pwbest(db, hr_target, bp, wb_next, wb_delta, iter_count + 1);
    }

    return wb_next;
}

//=========================================================
// Iterative functions to calculate DP given H (enthalpy at DP)
//=========================================================

export function H_DP_error(dblADPguess, dblH, dblBP) {
    return dblH - Ph_hr(dblADPguess, Phr_rh(dblADPguess, 1.00, dblBP));
}

export function DP_H_BetterGuess(dblADPguess, dblH, dblBP) {
    // Use Newton's method to estimate a better guess.
    const dblCurrentError = H_DP_error(dblADPguess, dblH, dblBP);
    const dblSlope = (dblCurrentError - H_DP_error(dblADPguess + 0.001, dblH, dblBP)) / 0.001;

    // To lose y amount of rise must move delta_x = y/slope from current position.
    return dblADPguess + dblCurrentError / dblSlope;
}

export function DP_H(dblADP_initialguess, dblH, dblBP) {
    let dblADP_PreviousEstimate = dblADP_initialguess;
    let dblADP_Estimate;

    // Repeat until error on estimate less than criteria.
    for (let j = 1; j <= 10; j++) {
        dblADP_Estimate = DP_H_BetterGuess(dblADP_PreviousEstimate, dblH, dblBP);

        if (Math.abs(H_DP_error(dblADP_Estimate, dblH, dblBP)) < 0.001) {
            return dblADP_Estimate;
        }
        dblADP_PreviousEstimate = dblADP_Estimate;
    }

    return dblADP_Estimate;
}

//=========================================================
// Apparatus Dew Point and Bypass factor routines (EnergyPlus)
//=========================================================

export function STratio_EnergyPlus_A0(dblA0, dblGrossTotCap_Stage_Adjusted_Btuh, dblCFM_stage, objStage,
                                       dblEHR, dblEDB, dblBP) {
    // This function uses A0 (at ARI) and the CFM to calculate an adjusted BPF at the given CFM
    const dblEWB = Pwb_hr(dblEDB, dblEHR, dblBP);
    const dblBPF_adjusted = BPF_FromA0(dblA0, dblEDB, dblEWB, dblBP, dblCFM_stage);

    const result = STratio_EnergyPlus(dblBPF_adjusted, dblGrossTotCap_Stage_Adjusted_Btuh, dblCFM_stage, objStage,
                                       dblEHR, dblEDB, dblBP);
    
    // Return both the S/T ratio and the adjusted BPF
    return { stRatio: result, bpfAdjusted: dblBPF_adjusted };
}

export function STratio_EnergyPlus(dblBPF_adjusted, dblGrossTotCap_Stage_Adjusted_Btuh, dblCFM_stage, objStage,
                                    dblEHR, dblEDB, dblBP) {
    // WARNING: dblBPF_adjusted is NOT adjusted in this function for CFM.
    const dblMassFlowDryAir_lbsPerHour = dblCFM_stage * 60 / Psv_hr(dblEDB, dblEHR, dblBP);

    const dblH_Entering = Ph_hr(dblEDB, dblEHR);
    const dblH_ADP = dblH_Entering - ((dblGrossTotCap_Stage_Adjusted_Btuh / dblMassFlowDryAir_lbsPerHour) / (1 - dblBPF_adjusted));

    const dblADP_DB = DP_H(dblEDB, dblH_ADP, dblBP);  // Iterative, dblEDB is starting guess
    const dblADP_HR = Phrs(dblADP_DB, dblBP);

    const dblH_Tin_Hadp = Ph_hr(dblEDB, dblADP_HR);  // The elbow

    // Use the ADP and the elbow to calculate the S/T = Delta_H_sensible/Delta_H_total
    let dblSTratio = (dblH_Tin_Hadp - dblH_ADP) / (dblH_Entering - dblH_ADP);

    // Crop the results.
    if (dblSTratio > 1) {
        return 1;
    } else if (dblSTratio < 0) {
        return 0;
    } else {
        return dblSTratio;
    }
}

export function A0_FromBPF(dblBPF_AHRI, dblCFM_AHRI) {
    // Use BP Factor and flow-rate at AHRI conditions to calculate the A0 factor.
    // BPF = exp(-A0/mdot)
    // So A0 = -mdot*log(BPF)
    const dblEHR = Phr_wb(80, 67, dblStandardPressure);
    const dblMassFlowDryAir_lbsPerHour = dblCFM_AHRI * 60 / Psv_hr(80, dblEHR, dblStandardPressure);

    return -dblMassFlowDryAir_lbsPerHour * Math.log(dblBPF_AHRI);
}

export function BPF_FromA0(dblA0_BPF, dblEDB, dblEWB, dblBP, dblCFM_Stage) {
    // Now if you know the A0 factor in the BPF coil model, then you can calculate the
    // BPF for new conditions as affected by CFM and density (i.e., a different mass flow rate).
    // BPF = exp(-A0/mdot)
    const dblEHR = Phr_wb(dblEDB, dblEWB, dblStandardPressure);
    const dblMassFlowDryAir_lbsPerHour = dblCFM_Stage * 60 / Psv_hr(dblEDB, dblEHR, dblBP);

    return Math.exp(-dblA0_BPF / dblMassFlowDryAir_lbsPerHour);
}

//=========================================================
// Apparatus Dew Point and Bypass factor routines
//=========================================================

// Apparatus Dew Point (ADP) via marching along the condition line from
// supply toward saturation.  Returns { adp, hrAtADP, errorMessage }.
export function ADP_main(dblSDB, dblSHR, dblEDB, dblEHR, dblBP) {
    // Returns { adp, hrAtADP, errorMessage }
    const dblTemperatureStep = 0.2;
    let strErrorMessage = "";

    // First do some quick checks on the inputs.
    if (dblSHR < 0) {
        return { adp: null, hrAtADP: null, errorMessage: "ADP: dblSHR < 0" };
    }

    // March toward the dewpoint at step increments
    const dblSlope = (dblEHR - dblSHR) / (dblEDB - dblSDB);

    // Initialize
    let dblADP_candidate = dblSDB;
    let dblHR_candidate = dblSHR;
    let dblADP_candidate_previous, dblHR_candidate_previous;
    let dblDP_Delta = 1000;
    let dblDP_Delta_previous;

    for (let intN = 1; intN < 1000; intN++) {
        dblADP_candidate_previous = dblADP_candidate;
        dblADP_candidate = dblSDB - intN * dblTemperatureStep;

        dblHR_candidate_previous = dblHR_candidate;
        dblHR_candidate = dblSHR - intN * dblTemperatureStep * dblSlope;

        // Calculate the dewpoint at this candidate humidity ratio.
        const dblDP_at_candidate_hr = Pdp_hr(dblHR_candidate, dblBP);

        dblDP_Delta_previous = dblDP_Delta;
        dblDP_Delta = dblADP_candidate - dblDP_at_candidate_hr;

        // Check to see if time to stop (too cold). Now you can interpolate
        if (dblDP_Delta < 0) {
            // Interpolate to find where dblDP_Delta = 0.
            const adp = dblADP_candidate + (Math.abs(dblDP_Delta) / (Math.abs(dblDP_Delta) + Math.abs(dblDP_Delta_previous))) *
                                           (dblADP_candidate_previous - dblADP_candidate);

            const hrAtADP = dblHR_candidate + (Math.abs(dblDP_Delta) / (Math.abs(dblDP_Delta) + Math.abs(dblDP_Delta_previous))) *
                                              (dblHR_candidate_previous - dblHR_candidate);

            return { adp, hrAtADP, errorMessage: "" };
        }

        // Check for bad stuff
        if (dblSDB > dblEDB) {
            return { adp: null, hrAtADP: null, errorMessage: "ADP: dblSDB > dblEDB, INT=" + intN };
        } else if (dblSHR > dblEHR) {
            return { adp: null, hrAtADP: null, errorMessage: "ADP: dblSHR > dblEHR, INT=" + intN };
        } else if (dblDP_Delta > dblDP_Delta_previous) {
            return { adp: null, hrAtADP: null, errorMessage: "ADP: dblDP_Delta > dblDP_Delta_previous, not finding saturation point, INT=" + intN };
        } else if (intN === 999) {
            strErrorMessage = "ADP: Do loop never met exit conditions, INT=" + intN;
        }
    }

    return { adp: null, hrAtADP: null, errorMessage: strErrorMessage };
}

//=========================================================
// Supply Conditions (Closed Form)
//=========================================================

// Supply air conditions (closed-form) from total capacity, S/T ratio,
// airflow, and entering conditions.  Returns { shr, sdb, errorMessage }.
export function SupplyCond_CF(dblTotalCap_Btuh, dblSTratio, dblCFM, dblEDB, dblEHR, dblBP) {
    // Returns { shr, sdb, errorMessage }
    
    // Analyze in two steps following Example 3 in Chapter 6 of ASHRAE handbook.
    const dblLatentCap_Btuh = (1 - dblSTratio) * dblTotalCap_Btuh;
    const dblMassFlowDryAir_lbsPerHour = dblCFM * 60 / Psv_hr(dblEDB, dblEHR, dblBP);

    const dblDelta_W = dblLatentCap_Btuh / (dblMassFlowDryAir_lbsPerHour * (1061 + 0.444 * dblEDB));
    const dblSHR = dblEHR - dblDelta_W;

    // Next calculate the supply temperature.
    const dblSensibleCap_Btuh = dblTotalCap_Btuh * dblSTratio;
    const dblDelta_T = dblSensibleCap_Btuh / (dblMassFlowDryAir_lbsPerHour * (0.24 + 0.444 * dblSHR));
    const dblSDB = dblEDB - dblDelta_T;

    // Do some checking on the result.
    let strErrorMessage = "";
    if (dblSDB < Pdp_hr(dblSHR, dblBP)) {
        strErrorMessage = "Supply: SDB is less than dewpoint for the calculated SHR (it's on the other side of the saturation curve).";
    } else if (dblSHR < 0) {
        strErrorMessage = "Supply: SHR is less than 0.";
    }

    return { shr: dblSHR, sdb: dblSDB, errorMessage: strErrorMessage };
}

export function SupplyCond_better(dblTotalCap_Btuh, dblSTratio, dblCFM, dblEDB, dblEHR, dblBP) {
    // This function includes the effects of the condensate.
    // Returns { shr, sdb, errorMessage }

    // First get the estimates of the supply conditions.
    let result = SupplyCond_CF(dblTotalCap_Btuh, dblSTratio, dblCFM, dblEDB, dblEHR, dblBP);

    // Estimate the enthalpy of the exiting condensate
    const dblMassFlowDryAir_lbsPerHour = dblCFM * 60 / Psv_hr(dblEDB, dblEHR, dblBP);
    const dblMassFlowWater_lbsPerHour = dblMassFlowDryAir_lbsPerHour * (dblEHR - result.shr);
    const dblh_exitingwater = result.sdb - 32;
    const h_flux_condensate = dblMassFlowWater_lbsPerHour * dblh_exitingwater;

    // Now increase the capacity by this amount and recalculate the supply conditions
    return SupplyCond_CF(dblTotalCap_Btuh + h_flux_condensate, dblSTratio, dblCFM, dblEDB, dblEHR, dblBP);
}

export function TotalCapacity(dblCFM, dblBP, dblEDB, dblEHR, dblSDB, dblSHR) {
    // This routine includes the condensate term in the energy balance.
    const dblMassFlowDryAir_lbsPerHour = dblCFM * 60 / Psv_hr(dblEDB, dblEHR, dblBP);
    const dblh_exitingwater = dblSDB - 32;
    const dblMassFlowWater_lbsPerHour = dblMassFlowDryAir_lbsPerHour * (dblEHR - dblSHR);

    return dblMassFlowDryAir_lbsPerHour * (Ph_hr(dblEDB, dblEHR) - Ph_hr(dblSDB, dblSHR))
           - dblMassFlowWater_lbsPerHour * dblh_exitingwater;
}

export function BPF(dblEDB, dblSDB, dblADP) {
    // Bypass factor
    return (dblSDB - dblADP) / (dblEDB - dblADP);
}

// Compute BPF from capacity and S/T ratio: supply conditions → ADP → BPF.
// Returns { bpf, errorMessage }.
export function BPF_ADP_SC(dblTotalCap_Btuh, dblSTratio, dblCFM, dblEDB, dblEHR, dblBP) {
    // Calc supply conditions (SC) from capacity and S/T, then ADP, and finally BPF
    // Returns { bpf, errorMessage }

    const supplyResult = SupplyCond_CF(dblTotalCap_Btuh, dblSTratio, dblCFM, dblEDB, dblEHR, dblBP);
    if (supplyResult.errorMessage !== "") {
        return { bpf: null, errorMessage: supplyResult.errorMessage };
    }

    const adpResult = ADP_main(supplyResult.sdb, supplyResult.shr, dblEDB, dblEHR, dblBP);
    if (adpResult.errorMessage !== "") {
        return { bpf: null, errorMessage: adpResult.errorMessage };
    }

    return { bpf: BPF(dblEDB, supplyResult.sdb, adpResult.adp), errorMessage: "" };
}

export function SC_ADP(dblTotalCap_Btuh, dblSTratio, dblCFM, dblEDB, dblEHR, dblBP) {
    // Calc supply conditions (SC) from capacity and S/T, then ADP.
    // Returns { sdb, shr, adp, hrAtADP, errorMessage }

    const supplyResult = SupplyCond_CF(dblTotalCap_Btuh, dblSTratio, dblCFM, dblEDB, dblEHR, dblBP);
    if (supplyResult.errorMessage !== "") {
        return { sdb: null, shr: null, adp: null, hrAtADP: null, errorMessage: supplyResult.errorMessage };
    }

    const adpResult = ADP_main(supplyResult.sdb, supplyResult.shr, dblEDB, dblEHR, dblBP);
    if (adpResult.errorMessage !== "") {
        return { sdb: null, shr: null, adp: null, hrAtADP: null, errorMessage: adpResult.errorMessage };
    }

    return {
        sdb: supplyResult.sdb,
        shr: supplyResult.shr,
        adp: adpResult.adp,
        hrAtADP: adpResult.hrAtADP,
        errorMessage: ""
    };
}
