/**
 * Performance Module - JavaScript ES Module
 * Converted from performance_module.asp (VBScript)
 * 
 * Module for sharing capacity and performance related routines.
 * 
 * Dependencies: psychro.js
 */

import {
    Phr_wb, Psv_hr, Ph_hr, Prh_hr, Pwb_hr,
    SupplyCond_CF, STratio_EnergyPlus_A0,
    dblStandardPressure, dblKWtoKBTUH
} from './psychro.js';

function MaxOfTwoValues(dblCandidate, dblCurrentMax) {
    return dblCandidate > dblCurrentMax ? dblCandidate : dblCurrentMax;
}

function MinOfTwoValues(dblCandidate, dblCurrentMin) {
    return dblCandidate < dblCurrentMin ? dblCandidate : dblCurrentMin;
}

//=========================================================
// Temperature Conversion
//=========================================================

export function C_from_F(dblFar) {
    return (5.0 / 9.0) * (dblFar - 32.0);
}

//=========================================================
// Capacity Curves
//=========================================================

/**
 * Total Capacity Correction factor based on operating conditions.
 * @param {Object} objSD - System data object
 * @param {string} strMaker - Manufacturer identifier
 * @param {number} dblNetTotCap_AHRI_KBtuh - Net total capacity at AHRI conditions (kBtuh)
 * @param {Object} objStage - Stage object with CapacityFraction, FlowFraction, StagePairType
 * @param {number} ODB - Outdoor dry bulb temperature (°F)
 * @param {number} EWB - Entering wet bulb temperature (°F)
 * @returns {number} Capacity correction factor
 */
export function Tot_Capacity_Correction(objSD, strMaker, dblNetTotCap_AHRI_KBtuh, objStage, ODB, EWB) {
    let dblFF;
    let Tot_Capacity_Correction_temp, Tot_Capacity_Correction_flow;
    let C0, C1, C2, C3, C4, C5;
    let ODB_C, EWB_C;

    // Flow affects total capacity only in the Advanced Controls case.
    if (objSD.Specific_RTU === "Advanced Controls") {
        dblFF = objStage.FlowFraction;
    } else {
        dblFF = 1.0;
    }

    if (objSD.Specific_RTU === "Variable-Speed Compressor") {
        // Daikin Unit.
        Tot_Capacity_Correction_temp = -2.39099158422236 + (0.0517875312429516 * EWB) + (-0.000262093586694283 * EWB ** 2) + (0.0298244056511833 * ODB) + (-0.000184849593859541 * ODB ** 2) + (-7.40270588964858E-06 * EWB * ODB);
        Tot_Capacity_Correction_flow = 1.00;

    } else if (objSD.Specific_RTU === "Three Stages") {
        if (objStage.CapacityFraction === 1.00) {
            C0 = 2.16908023;
            C1 = -0.04741753;
            C2 = 0.00054899;
            C3 = 0.00394090;
            C4 = -0.00001554;
            C5 = -0.00010878;
        } else if (objStage.CapacityFraction === 0.60) {
            C0 = 2.62596194;
            C1 = -0.05969721;
            C2 = 0.00064792;
            C3 = 0.00367327;
            C4 = -0.00001356;
            C5 = -0.00011917;
        } else if (objStage.CapacityFraction === 0.40) {
            C0 = 2.68873560;
            C1 = -0.06470293;
            C2 = 0.00067946;
            C3 = 0.00676853;
            C4 = -0.00001991;
            C5 = -0.00013382;
        } else {
            console.warn("No match in Tot_Capacity_Correction. CF=" + objStage.CapacityFraction);
            return 1.0;
        }

        Tot_Capacity_Correction_temp = C0 + C1 * EWB + C2 * EWB ** 2 + C3 * ODB + C4 * ODB ** 2 + C5 * EWB * ODB;
        Tot_Capacity_Correction_flow = 1.00;

    } else if (objSD.Specific_RTU === "None" || objSD.Specific_RTU === "Advanced Controls") {
        if (objSD.DOE2_Curves === "DOE2") {
            // DOE-2 Curves
            if (dblNetTotCap_AHRI_KBtuh >= 60) {
                // Default PSZ curve (30 Tons; 360 kBtuh)
                Tot_Capacity_Correction_temp = 0.87403018 + (-0.0011416 * EWB) + (0.00017110 * EWB ** 2) + (-0.00295700 * ODB) + (0.00001018 * ODB ** 2) + (-0.00005917 * EWB * ODB);
                Tot_Capacity_Correction_flow = 0.47278589 + (1.24334150 * dblFF) + (-1.03870550 * dblFF ** 2) + (0.32257813 * dblFF ** 3);
            } else {
                // Default RESYS curve (3 Tons; 36 kBtuh)
                Tot_Capacity_Correction_temp = 0.60034040 + (0.00228726 * EWB) + (-0.00001280 * EWB ** 2) + (0.00138975 * ODB) + (-0.00008060 * ODB ** 2) + (0.00014125 * EWB * ODB);
                Tot_Capacity_Correction_flow = 0.80 + (0.20 * dblFF);
            }

        } else if (objSD.DOE2_Curves === "Carrier") {
            // Carrier substitute for DOE-2 curves (Centigrade).
            C0 = 0.75030980;
            C1 = 0.01611210;
            C2 = 0.00081690;
            C3 = -0.00357190;
            C4 = -0.00018740;
            C5 = -0.00001780;

            ODB_C = C_from_F(ODB);
            EWB_C = C_from_F(EWB);

            Tot_Capacity_Correction_temp = C0 + C1 * EWB_C + C2 * EWB_C ** 2 + C3 * ODB_C + C4 * ODB_C ** 2 + C5 * EWB_C * ODB_C;
            Tot_Capacity_Correction_flow = 1.00;

        } else {
            console.warn("No match in Tot_Capacity_Correction. DOE2_Curves=" + objSD.DOE2_Curves);
            return 1.0;
        }

    } else {
        console.warn("No match in Tot_Capacity_Correction. Specific_RTU=" + objSD.Specific_RTU);
        return 1.0;
    }

    return Tot_Capacity_Correction_temp * Tot_Capacity_Correction_flow;
}

//=========================================================
// Efficiency Curves
//=========================================================

/**
 * Efficiency Correction factor (EIR correction) based on operating conditions.
 * @param {Object} objSD - System data object
 * @param {string} strMaker - Manufacturer identifier
 * @param {Object} objStage - Stage object
 * @param {number} ODB - Outdoor dry bulb temperature (°F)
 * @param {number} EWB - Entering wet bulb temperature (°F)
 * @param {number} EDB - Entering dry bulb temperature (°F)
 * @param {number} totalCap - Total capacity for curve selection (kBtuh)
 * @returns {number} Efficiency correction factor
 */
export function Efficiency_Correction(objSD, strMaker, objStage, ODB, EWB, EDB, totalCap = 60) {
    let dblFF;
    let Efficiency_Correction_temp, Efficiency_Correction_flow;
    let C0, C1, C2, C3, C4, C5, C6, C7, C8, C9, C10;
    let ODB_C, EWB_C, EDB_C, SDB_C;
    let SDB;

    // Flow affects efficiency only in the Advanced Controls case.
    if (objSD.Specific_RTU === "Advanced Controls") {
        dblFF = objStage.FlowFraction;
    } else {
        dblFF = 1.0;
    }

    if (objSD.Specific_RTU === "Variable-Speed Compressor") {
        // Daikin Unit - Centigrade temperatures
        C0 = -0.5966663;
        C1 = 0.24754897;
        C2 = -0.0088454;
        C3 = -0.0036095;
        C4 = -0.0085282;
        C5 = 0.00072631;
        C6 = 0.01271645;
        C7 = -0.0018991;
        C8 = -0.0024883;
        C9 = -0.0008663;
        C10 = 0.00250916;

        // Supply temperature
        SDB = objStage.SDB;

        // Convert to centigrade
        ODB_C = C_from_F(ODB);
        EWB_C = C_from_F(EWB);
        EDB_C = C_from_F(EDB);
        SDB_C = C_from_F(SDB);

        Efficiency_Correction_temp = C0 + C1 * SDB_C + C2 * EWB_C ** 2 + C3 * EDB_C ** 2 + C4 * SDB_C ** 2 + C5 * ODB_C ** 2 + C6 * EWB_C * EDB_C + C7 * EWB_C * ODB_C + C8 * EDB_C * SDB_C + C9 * EDB_C * ODB_C + C10 * SDB_C * ODB_C;
        Efficiency_Correction_flow = 1.00;

    } else if (objSD.Specific_RTU === "Three Stages") {
        if (objStage.CapacityFraction === 1.00) {
            C0 = -1.01155368;
            C1 = 0.05389628;
            C2 = -0.00033881;
            C3 = -0.00389377;
            C4 = 0.00019559;
            C5 = -0.00023034;
        } else if (objStage.CapacityFraction === 0.60) {
            C0 = -1.43280192;
            C1 = 0.06809725;
            C2 = -0.00044252;
            C3 = -0.00527563;
            C4 = 0.00020666;
            C5 = -0.00023623;
        } else if (objStage.CapacityFraction === 0.40) {
            C0 = -1.65119376;
            C1 = 0.07413652;
            C2 = -0.00049982;
            C3 = -0.00567775;
            C4 = 0.00019535;
            C5 = -0.00020417;
        } else {
            console.warn("No match in Efficiency_Correction. CF=" + objStage.CapacityFraction);
            return 1.0;
        }

        Efficiency_Correction_temp = C0 + C1 * EWB + C2 * EWB ** 2 + C3 * ODB + C4 * ODB ** 2 + C5 * EWB * ODB;
        Efficiency_Correction_flow = 1.00;

    } else if (objSD.Specific_RTU === "None" || objSD.Specific_RTU === "Advanced Controls") {
        if (objSD.DOE2_Curves === "DOE2") {
            // DOE-2 Curves
            if (totalCap >= 60) {
                // Default PSZ curve (30 Tons; 360 kBtuh)
                Efficiency_Correction_temp = -1.06393100 + (0.03065843 * EWB) + (-0.00012690 * EWB ** 2) + (0.01542130 * ODB) + (0.00004973 * ODB ** 2) + (-0.00020960 * EWB * ODB);
                Efficiency_Correction_flow = 1.00794840 + (0.34544129 * dblFF) + (-0.69228910 * dblFF ** 2) + (0.33889943 * dblFF ** 3);
            } else {
                // Default RESYS curve (3 Tons; 36 kBtuh)
                Efficiency_Correction_temp = -0.96177870 + (0.04817751 * EWB) + (-0.00023110 * EWB ** 2) + (0.00324392 * ODB) + (0.00014876 * ODB ** 2) + (-0.00029520 * EWB * ODB);
                Efficiency_Correction_flow = 1.156 + (-0.1816 * dblFF) + (0.0256 * dblFF ** 2);
            }

        } else if (objSD.DOE2_Curves === "Carrier") {
            // Carrier substitute for DOE-2 curves (Centigrade).
            C0 = 0.4152146633;
            C1 = 0.0093230741;
            C2 = 0.0002407406;
            C3 = 0.0150246809;
            C4 = 0.0008229240;
            C5 = -0.0018007980;

            ODB_C = C_from_F(ODB);
            EWB_C = C_from_F(EWB);

            Efficiency_Correction_temp = C0 + C1 * EWB_C + C2 * EWB_C ** 2 + C3 * ODB_C + C4 * ODB_C ** 2 + C5 * EWB_C * ODB_C;
            Efficiency_Correction_flow = 1.00;

        } else {
            console.warn("No match in Efficiency_Correction. DOE2_Curves=" + objSD.DOE2_Curves);
            return 1.0;
        }

    } else {
        console.warn("No match in Efficiency_Correction. Specific_RTU=" + objSD.Specific_RTU);
        return 1.0;
    }

    let result = Efficiency_Correction_temp * Efficiency_Correction_flow;

    // Limit to deal with model behavior when entering dry bulb approaches freezing.
    if (result < 0.01) {
        result = 0.01;
    }

    return result;
}

/**
 * Sensible Capacity Correction (DOE-2 relationship)
 * Note: Not currently used - ADP/BPF method gives direct S/T value instead.
 */
export function Sens_Capacity_Correction(ODB, EWB, totalCap = 60) {
    if (totalCap >= 60) {
        return 4.83529620 + (-0.05753070 * EWB) + (0.00006155 * EWB ** 2) + (-0.00526830 * ODB) + (0.00000317 * ODB ** 2) + (0.00003375 * EWB * ODB);
    } else {
        return 6.52756980 + (-0.12613750 * EWB) + (0.00056879 * EWB ** 2) + (0.00907575 * ODB) + (-0.00004830 * ODB ** 2) + (-0.00000875 * EWB * ODB);
    }
}

//=========================================================
// S/T Ratio Calculation
//=========================================================

/**
 * Calculate the sensible to total (S/T) ratio at entering conditions.
 * @returns {Object} { stRatio, bpfAdjusted }
 */
export function ST_Ratio_engine(strMaker, dblNetTotCap_AHRI_KBtuh, dblCFM_AHRI, objStage, objSD, dblTC_Factor,
                                 dblODB, dblEWB, dblEDB, dblBP) {
    // Calculate the sensible to total (S/T) ratio at the entering conditions,
    // by adjusting the BPF that was calculated at AHRI conditions.

    // The fan-heat power needs to be corrected for stage level...
    const dblGrossTotCap_Stage_Adjusted_KBtuh = NetTotCap_Stage_Adjusted_KBtuh(objSD, dblNetTotCap_AHRI_KBtuh, objStage, dblODB, dblEWB, dblEDB) +
                                                 objSD.BFn_kw * dblKWtoKBTUH * (objStage.FlowFraction ** objSD.N_Affinity);

    const dblEHR = Phr_wb(dblEDB, dblEWB, dblBP);

    // Capacity and flow are stage specific here. These are NOT nominal values.
    const dblCFM_stage = dblCFM_AHRI * objStage.FlowFraction;
    
    const result = STratio_EnergyPlus_A0(objSD.A0_BPF, dblGrossTotCap_Stage_Adjusted_KBtuh * 1000, dblCFM_stage, objStage,
                                          dblEHR, dblEDB, dblBP);

    return result;  // { stRatio, bpfAdjusted }
}

//=========================================================
// Capacity Functions (Gross/Net conversions)
//=========================================================

/**
 * Special function for the Daikin unit - part-load capacity factor.
 */
export function VS_Compressor_PL_CapFactor(dblCapacityFraction) {
    const dblCF = dblCapacityFraction;
    return -0.036240442 + (1.275963118 * dblCF) + (-0.288191485 * dblCF ** 2) + (0.048967033 * dblCF ** 3);
}

export function VS_Compressor_PL_EIRFactor(dblCapacityFraction) {
    const dblCF = dblCapacityFraction;
    return -0.120651751 + (9.026346984 * dblCF) + (-15.86922715 * dblCF ** 2) + (7.966206349 * dblCF ** 3);
}

export function VS_Compressor_PL_PowerFactor(dblCapacityFraction) {
    const dblCF = dblCapacityFraction;
    if (dblCF > 0) {
        return VS_Compressor_PL_CapFactor(dblCF) * VS_Compressor_PL_EIRFactor(dblCF);
    }
    // The regressions for the Daikin don't quite give a clean zero, so force it here.
    return 0;
}

/**
 * Net Total Capacity at stage level, adjusted for operating conditions.
 */
export function NetTotCap_Stage_Adjusted_KBtuh(objSD, dblNetTotCap_AHRI_KBtuh, objStage, dblODB, dblEWB, dblEDB) {
    // Full-capacity fan energy.
    const dblFanEnergy_AHRI_KBtuh = objSD.BFn_kw * dblKWtoKBTUH;

    // Remove the influence of the fan energy (add to Net) to get a gross capacity
    const dblGrossTotCap_AHRI_KBtuh = dblNetTotCap_AHRI_KBtuh + dblFanEnergy_AHRI_KBtuh;

    // Use spreadsheet model if available.
    let dblGrossTotCap_Correction;
    if (objSD.GrossCapacity_KBtu_Model && objSD.GrossCapacity_KBtu_Model.IsSolved) {
        dblGrossTotCap_Correction = objSD.Predict_Gross_Capacity_Correction(dblODB, dblEWB);
    } else {
        dblGrossTotCap_Correction = Tot_Capacity_Correction(objSD, "generic", dblNetTotCap_AHRI_KBtuh, objStage, dblODB, dblEWB);
    }

    // Store for use in tables and plots.
    if (objStage) {
        objStage.TCap_CF = dblGrossTotCap_Correction;
    }

    const dblGrossTotCap_Adjusted_KBtuh = dblGrossTotCap_AHRI_KBtuh * dblGrossTotCap_Correction;

    // For the Daikin unit use special part-load function. Don't use simple CF scaling.
    let dblPartLoad_CapFactor;
    if (objSD.Specific_RTU === "Variable-Speed Compressor") {
        dblPartLoad_CapFactor = VS_Compressor_PL_CapFactor(objStage.CapacityFraction);
    } else {
        dblPartLoad_CapFactor = objStage.CapacityFraction;
    }

    return (dblGrossTotCap_Adjusted_KBtuh * dblPartLoad_CapFactor) - (dblFanEnergy_AHRI_KBtuh * (objStage.FlowFraction ** objSD.N_Affinity));
}

/**
 * Net Sensible Capacity at stage level, adjusted for operating conditions.
 */
export function NetSenCap_Stage_Adjusted_KBtuh(objSD, dblNetTotCap_AHRI_KBtuh, dblCFM, objStage, dblTC_Factor,
                                                dblODB, dblEWB, dblEDB, dblBP) {
    // To get a stage-level fan-heat power this must be modified by flowfraction^N_Affinity
    const dlbFanEnergy_Stage_KBtuh = objSD.BFn_kw * dblKWtoKBTUH * (objStage.FlowFraction ** objSD.N_Affinity);

    // Remove the influence of the fan energy (add to Net) to get a gross capacity
    const dblGrossTotCap_Stage_Adjusted_KBtuh = NetTotCap_Stage_Adjusted_KBtuh(objSD, dblNetTotCap_AHRI_KBtuh, objStage, dblODB, dblEWB, dblEDB) +
                                                 dlbFanEnergy_Stage_KBtuh;

    // Use spreadsheet model if available.
    let dblST_Ratio_AtEntering;
    if (objSD.ST_Ratio_Model && objSD.ST_Ratio_Model.IsSolved) {
        dblST_Ratio_AtEntering = objSD.Predict_ST_Ratio(dblODB, dblEWB, dblEDB);
    } else {
        const stResult = ST_Ratio_engine("generic", dblNetTotCap_AHRI_KBtuh, dblCFM, objStage, objSD, dblTC_Factor, dblODB, dblEWB, dblEDB, dblBP);
        dblST_Ratio_AtEntering = stResult.stRatio;
    }

    // Put this S/T in the stage object for reporting.
    if (objStage) {
        objStage.ST_Ratio = dblST_Ratio_AtEntering;
    }

    // Get the supply conditions for use with the Daikin power calculations.
    const dblEHR = Phr_wb(dblEDB, dblEWB, dblBP);
    const supplyResult = SupplyCond_CF(dblGrossTotCap_Stage_Adjusted_KBtuh * 1000, dblST_Ratio_AtEntering, dblCFM * objStage.FlowFraction,
                                        dblEDB, dblEHR, dblBP);

    // Put this supply temp in the stage object for use in the power calculations (needed for Daikin unit).
    if (objStage) {
        objStage.SDB = supplyResult.sdb;
    }

    const dblGrossSenCap_Stage_Adjusted_KBtuh = dblGrossTotCap_Stage_Adjusted_KBtuh * dblST_Ratio_AtEntering;

    // Now put the stage-level fan effects back in (subtract from Gross) to get a NET stage capacity.
    return dblGrossSenCap_Stage_Adjusted_KBtuh - dlbFanEnergy_Stage_KBtuh;
}

//=========================================================
// Humidity Tracking
//=========================================================

/**
 * Limit inside relative humidity to 20-65% range.
 */
export function IRH_limited(dblIRH_tracking) {
    if (dblIRH_tracking > 0.65) {
        return 0.65;
    } else if (dblIRH_tracking < 0.20) {
        return 0.20;
    } else {
        return dblIRH_tracking;
    }
}

/**
 * Calculate inside relative humidity - either tracking outdoor or fixed.
 */
export function IRH_Track_OR_Set(strTrackMode, dblOHR, dblIDB_setpoint, dblIRH_setpoint, dblPressure) {
    if (strTrackMode === "on") {
        // If you let the IHR track the OHR, the indoor humidity can be calculated to be:
        const dblIHR = dblOHR;
        const dblInsideRelativeHumidity_tracking = Prh_hr(dblIDB_setpoint, dblIHR, dblPressure);

        // Check the inside relative humidity to see if it's within a limited range.
        return IRH_limited(dblInsideRelativeHumidity_tracking);
    } else {
        // Set it to the fixed IRH setpoint
        return dblIRH_setpoint;
    }
}

//=========================================================
// Ventilation Calculations
//=========================================================

/**
 * Sensible heat gain corresponding to the change of dry-bulb temperature for given airflow rate.
 * ASHRAE equation 29 in Chapter 26 of Fundamentals Handbook 1989.
 * @returns {number} Sensible ventilation load (kBtuh)
 */
export function SensVentLoad(dblCFM, dblOHR, dblODB, dblIDB, dblPressure) {
    return dblCFM * (60 / Psv_hr(dblODB, dblOHR, dblPressure)) * (0.24 + (dblOHR * 0.444)) * (dblODB - dblIDB) / 1000;
}

/**
 * Calculate CFM needed to produce a known rise in ventilation load per degree F.
 * Inverse of SensVentLoad equation.
 */
export function CFM_VentSlope(dblVentSlope_kBtuh_Per_F, dblOHR, dblODB, dblPressure) {
    return dblVentSlope_kBtuh_Per_F * (Psv_hr(dblODB, dblOHR, dblPressure) / 60) * (1 / (0.24 + (dblOHR * 0.444))) * 1000;
}

//=========================================================
// Mixed Air Calculations
//=========================================================

/**
 * Calculate the mixed air conditions entering the coils.
 * Based on ASHRAE p6.18: Adiabatic Mixing of Two Streams of Moist Air
 * @returns {Object} { EHR, EDB, EWB }
 */
export function Mixer2(objSD, dblIDB, dblIHR, dblODB, dblOHR, dblBP,
                       dblVentilation_CFM, dblEvapFanFlow_CFM) {
    // The enthalpy of the two streams.
    const dblOH = Ph_hr(dblODB, dblOHR);
    const dblIH = Ph_hr(dblIDB, dblIHR);

    // Assume that if the ventilation flow exceeds the evaporator fan flow it must be coming into the building via another fan.
    let dblCFM_RoomAir, dblCFM_VentAir;
    if (dblVentilation_CFM > dblEvapFanFlow_CFM) {
        dblCFM_RoomAir = 0;
        dblCFM_VentAir = dblEvapFanFlow_CFM;
    } else {
        dblCFM_RoomAir = dblEvapFanFlow_CFM - dblVentilation_CFM;
        dblCFM_VentAir = dblVentilation_CFM;
    }

    // Calculate DryAir Mass Flows.
    const dbl_MassFlow_DA_RoomAir_lbsPerMin = dblCFM_RoomAir / Psv_hr(dblIDB, dblIHR, dblBP);
    const dbl_MassFlow_DA_VentAir_lbsPerMin = dblCFM_VentAir / Psv_hr(dblODB, dblOHR, dblBP);

    // Sum the two streams to get total dry air mass flow.
    const dbl_MassFlow_DA_MixedAir_lbsPerMin = dbl_MassFlow_DA_RoomAir_lbsPerMin + dbl_MassFlow_DA_VentAir_lbsPerMin;

    // Calculate the Humidity Ratio of the mixed air stream (mass flow weighted).
    const dblEHR = (dblIHR * dbl_MassFlow_DA_RoomAir_lbsPerMin + dblOHR * dbl_MassFlow_DA_VentAir_lbsPerMin) / dbl_MassFlow_DA_MixedAir_lbsPerMin;

    // Calculate the Enthalpy of the mixed air stream (mass flow weighted)
    const dblEH = (dblIH * dbl_MassFlow_DA_RoomAir_lbsPerMin + dblOH * dbl_MassFlow_DA_VentAir_lbsPerMin) / dbl_MassFlow_DA_MixedAir_lbsPerMin;

    // Invert the ASHRAE equation for H and solve for DB
    const dblEDB_bf = (dblEH - dblEHR * 1061) / (0.24 + dblEHR * 0.444);  // Entering DryBulb BEFORE the FAN (bf)

    // Now add in the fan heat. Assume (1-XX) of the fan motor energy gets into the air stream.
    // 70% fan efficiency and 85% motor efficiency (XX=.70*.85).
    const dblFanAndMotorEfficiency = 0.70 * 0.85;

    // Using ASHRAE equation (30) for H in psychrometric chapter of Fundamentals
    const dblDelta_T = (1000 * dblKWtoKBTUH * objSD.BFn_kw * (1 - dblFanAndMotorEfficiency)) /
                       (60 * dbl_MassFlow_DA_MixedAir_lbsPerMin * (0.240 + 0.444 * dblEHR));

    const dblEDB = dblEDB_bf + dblDelta_T;
    const dblEWB = Pwb_hr(dblEDB, dblEHR, dblBP);

    return { EHR: dblEHR, EDB: dblEDB, EWB: dblEWB };
}

export function CyclingEfficiency(objSD, dblLoadFraction) {
    if (dblLoadFraction < 1.0) {
        return (dblLoadFraction * (objSD.PLDegrFactor / 100)) + (100 - objSD.PLDegrFactor) / 100;
    }
    return 1.0;
}

export function FanPowerFactor_AffinityLaw(objSD, dblLoadFraction) {
    if (dblLoadFraction < 1.0) {
        return Math.pow(dblLoadFraction, objSD.N_Affinity);
    }
    return 1.0;
}

export function FanPower_PL_kW(objSD, dblFlowFraction, peak) {
    const dblFanPower_PL_kW = objSD.BFn_kw * FanPowerFactor_AffinityLaw(objSD, dblFlowFraction);
    if (peak && typeof peak === 'object') {
        peak.value = MaxOfTwoValues(peak.value ?? 0, dblFanPower_PL_kW);
    }
    return dblFanPower_PL_kW;
}

export function CondenserPower_PL_kW(objSD, objStage, dblODB, dblEWB, dblEDB, peak, totalCap) {
    const dblCondenserPowerAtTest = objSD.Cond_kw;

    // Full-load power correction at temperature conditions.
    // Match ASP: if spreadsheet condenser model exists, use it (no flow correction) instead of generic curves.
    let dblEfficiency_CF;
    let dblCondenserPower_CF;
    if (objSD?.Condenser_kW_Model?.IsSolved && typeof objSD.Predict_Condenser_Correction === 'function') {
        // ASP uses:
        //   Efficiency_Correction_Spreadsheet = Predict_Condenser_Correction / Predict_Gross_Capacity_Correction
        //   CondPower_CF = Predict_Gross_Capacity_Correction * Efficiency_Correction_Spreadsheet
        // which simplifies to Predict_Condenser_Correction.
        dblEfficiency_CF = 1.0;
        dblCondenserPower_CF = objSD.Predict_Condenser_Correction(dblODB, dblEWB);
    } else {
        dblEfficiency_CF = Efficiency_Correction(objSD, 'generic', objStage, dblODB, dblEWB, dblEDB, totalCap);
        dblCondenserPower_CF = Tot_Capacity_Correction(objSD, 'generic', totalCap, objStage, dblODB, dblEWB) * dblEfficiency_CF;
    }

    // Dump these correction factors for reporting/debug.
    if (objStage) {
        objStage.Efficiency_CF = dblEfficiency_CF;
        objStage.CondPower_CF = dblCondenserPower_CF;
    }

    const dblCond_AtBC_FL_kw = dblCondenserPowerAtTest * dblCondenserPower_CF;

    let dblCond_AtBC_FL_Stage_kw;
    if (objStage?.StagePairType === 'BmA') {
        dblCond_AtBC_FL_Stage_kw = objStage.CapacityFraction_Diff * dblCond_AtBC_FL_kw;
    } else {
        dblCond_AtBC_FL_Stage_kw = objStage.CapacityFraction * dblCond_AtBC_FL_kw;
    }

    let dblCond_AtBC_PL_Stage_kw;

    // Spreadsheet part-load correction (Normalized EER / NEER) overrides cycling degradation.
    if (objSD?.NEER_PL_Model?.IsSolved && typeof objSD.Predict_PartloadFactor === 'function') {
        if ((objStage?.RunTime ?? 0) > 1.0) {
            // Full-load capacity.
            dblCond_AtBC_PL_Stage_kw = dblCond_AtBC_FL_Stage_kw;
        } else {
            const fanControls = String(objSD?.FanControls || '');
            const isVariable = fanControls.slice(0, 1) === 'V';
            const capFracN = Number(objStage?.CapacityFraction);
            const loadFracN = Number(objStage?.LoadFraction);
            const lf = isVariable
                ? (Number.isFinite(capFracN) ? capFracN : 1.0)
                : (Number.isFinite(loadFracN) ? loadFracN : 1.0);

            const plc = objSD.Predict_PartloadFactor(lf * 100.0, dblODB);
            dblCond_AtBC_PL_Stage_kw = dblCond_AtBC_FL_Stage_kw / Math.max(1e-9, plc);
        }
    } else {
        // Staged condenser: part-load via cycling degradation.
        const fanControls = String(objSD?.FanControls || '');
        const isVariable = fanControls.slice(0, 1) === 'V';

        // Match ASP: for the specific Daikin "Variable-Speed Compressor" unit, use the special part-load
        // power factor polynomial instead of cycling degradation or simple scaling.
        if (isVariable && String(objSD?.Specific_RTU || '') === 'Variable-Speed Compressor') {
            dblCond_AtBC_PL_Stage_kw = dblCond_AtBC_FL_kw * VS_Compressor_PL_PowerFactor(objStage.CapacityFraction);
        } else if (isVariable) {
            // Match ASP: for variable-speed condenser fan units without spreadsheet part-load data,
            // split condenser fan vs compressor power and apply affinity laws to the fan.
            const condFanPct = Number(objSD?.CondFanPercent);
            const condFanFrac = Number.isFinite(condFanPct) ? (condFanPct / 100.0) : 0;

            // ASP defines condenser-fan FL kW at rating conditions (uncorrected).
            const dblCFan_FL_kw = dblCondenserPowerAtTest * condFanFrac;
            const capFrac = (typeof objStage?.CapacityFraction === 'number' && Number.isFinite(objStage.CapacityFraction))
                ? objStage.CapacityFraction
                : 1.0;
            const dblCFan_PL_kw = dblCFan_FL_kw * FanPowerFactor_AffinityLaw(objSD, capFrac);

            // Compressor portion uses corrected full-load kW minus the (uncorrected) condenser-fan FL kW,
            // matching the legacy ASP implementation.
            const dblComp_AtBC_FL_kw = dblCond_AtBC_FL_kw - dblCFan_FL_kw;
            const dblComp_AtBC_PL_kw = ((objStage?.RunTime ?? 0) <= 1.0)
                ? (dblComp_AtBC_FL_kw * capFrac)
                : dblComp_AtBC_FL_kw;

            dblCond_AtBC_PL_Stage_kw = dblComp_AtBC_PL_kw + dblCFan_PL_kw;
        } else {
            const loadFracN = Number(objStage?.LoadFraction);
            const dblLoadFraction = Number.isFinite(loadFracN) ? loadFracN : 1.0;
            dblCond_AtBC_PL_Stage_kw = dblCond_AtBC_FL_Stage_kw / CyclingEfficiency(objSD, dblLoadFraction);
        }
    }

    if (peak && typeof peak === 'object') {
        peak.value = MaxOfTwoValues(peak.value ?? 0, dblCond_AtBC_PL_Stage_kw);
    }

    return dblCond_AtBC_PL_Stage_kw;
}

export function StageFractionAboveVent(objSD, ventilationFraction) {
    const dblStageArray = objSD.StageLevels;
    const intMaxStageArrayIndex = dblStageArray.length - 1;
    for (let i = 0; i <= intMaxStageArrayIndex; i++) {
        const dblFlowFraction = dblStageArray[i];
        if (dblFlowFraction > ventilationFraction) return dblFlowFraction;
    }
    return 1.0;
}

export function FF(objSD, dblCapFraction, objSP, strMode, blnEconomizerRunning, dblODB, ventilationFraction) {
    if (objSD.Specific_RTU === 'Advanced Controls') {
        if (strMode === 'C-On') {
            if (dblCapFraction === 0.5) {
                if (dblODB >= 70.0) {
                    return 0.75;
                }
                return 0.90;
            }
            if (dblCapFraction === 1.0) {
                return 0.90;
            }
            return 0;
        }

        if (strMode === 'C-Off') {
            if (blnEconomizerRunning) {
                if (objSP?.IntegratedState === 'Attempting Integrated') {
                    return 0.90;
                }
                return 0.75;
            }

            if (String(objSD.FanControls || '').includes('Cycles With Compressor')) {
                return 0.0;
            }
            if (String(objSD.FanControls || '').includes('Always ON')) {
                return 0.40;
            }
            return 0;
        }

        return 0;
    }

    if (strMode === 'C-On') {
        if (String(objSD.FanControls || '').slice(0, 1) === '1') {
            return 1.0;
        }

        if (blnEconomizerRunning) {
            return dblCapFraction;
        }

        if (objSP?.A?.CondType === 'VC') {
            if (ventilationFraction > dblCapFraction) return ventilationFraction;
            return dblCapFraction;
        }

        if (objSP?.A?.CondType === 'Staged') {
            return MaxOfTwoValues(StageFractionAboveVent(objSD, ventilationFraction), dblCapFraction);
        }

        return 0;
    }

    if (strMode === 'C-Off') {
        if (blnEconomizerRunning) {
            return 1.0;
        }

        if (String(objSD.FanControls || '').includes('Cycles With Compressor')) {
            return 0.0;
        }

        if (String(objSD.FanControls || '').includes('Always ON')) {
            if (objSP?.A?.CondType === 'VC') {
                return ventilationFraction;
            }
            if (objSP?.A?.CondType === 'Staged') {
                if (String(objSD.FanControls || '').slice(0, 1) === '1') {
                    return 1.0;
                }
                return StageFractionAboveVent(objSD, ventilationFraction);
            }
            return 1.0;
        }

        return 0;
    }

    return 0;
}

export function StageLevel(objSD, BCC, dblSensibleLoad_KBtuH, SP, opts) {
    const totalCap = opts?.totalCap;
    const cfm = opts?.cfm;
    const pressure = opts?.pressure;
    const ventilationFraction = opts?.ventilationFraction;

    const dblStageArray = objSD.StageLevels;
    const intMaxStageArrayIndex = dblStageArray.length - 1;

    SP.A.CapacityFraction = dblStageArray[0];
    SP.A.FlowFraction = FF(objSD, SP.A.CapacityFraction, SP, 'C-On', SP.EconomizerRunning, BCC.ODB, ventilationFraction);

    SP.B.ResetToZeroLoad();
    SP.BmA.ResetToZeroLoad();

    SP.A.SensCap_KBtuH = NetSenCap_Stage_Adjusted_KBtuh(objSD, totalCap, cfm, SP.A, 1, BCC.ODB, BCC.EWB, BCC.EDB, pressure);

    // Variable-capacity (VC): treat as modulating capacity at (near) continuous runtime.
    // Match legacy ASP intent for "Variable-Speed Compressor":
    // - Determine required capacity fraction from full-load sensible capacity.
    // - Clamp to system minimum capacity fraction.
    // - Run continuously at that capacity fraction (so fan and condenser power reflect the modulated point).
    const isVC =
        SP?.A?.CondType === 'VC' ||
        (typeof objSD?.FanControls === 'string' && objSD.FanControls.slice(0, 1) === 'V') ||
        String(objSD?.Specific_RTU || '') === 'Variable-Speed Compressor';
    if (isVC) {
        // If there's no sensible load, don't run the compressor.
        if (!(dblSensibleLoad_KBtuH > 0)) {
            SP.A.CapacityFraction = 0.0;
            SP.A.FlowFraction = FF(objSD, 0.0, SP, 'C-Off', SP.EconomizerRunning, BCC.ODB, ventilationFraction);
            SP.A.SensCap_KBtuH = 0.0;
            SP.A.LoadFraction = 0.0;
            SP.A.RunTime = 0.0;
            SP.pair_mode = 'A_only';
            return 0.0;
        }

        const minFrac = Number(objSD?.CapacityFraction_Min ?? 0);

        function _vsSensAtFrac(cf) {
            SP.pair_mode = 'A_only';
            SP.A.CapacityFraction = cf;
            SP.A.FlowFraction = FF(objSD, SP.A.CapacityFraction, SP, 'C-On', SP.EconomizerRunning, BCC.ODB, ventilationFraction);
            SP.A.SensCap_KBtuH = NetSenCap_Stage_Adjusted_KBtuh(objSD, totalCap, cfm, SP.A, 1, BCC.ODB, BCC.EWB, BCC.EDB, pressure);
            return Number(SP.A.SensCap_KBtuH);
        }

        // Full-load sensible capacity at bin conditions.
        // Match ASP VariableCapacityLevel: initialize CF=1 and FlowFraction=1 for this full-load capacity check.
        SP.pair_mode = 'A_only';
        SP.A.CapacityFraction = 1.0;
        SP.A.FlowFraction = 1.0;
        SP.A.SensCap_KBtuH = NetSenCap_Stage_Adjusted_KBtuh(objSD, totalCap, cfm, SP.A, 1, BCC.ODB, BCC.EWB, BCC.EDB, pressure);
        const fullSens = Number(SP.A.SensCap_KBtuH);

        if (Number.isFinite(fullSens) && fullSens > 0) {
            const load = Number(dblSensibleLoad_KBtuH);

            if (!(load > 0)) {
                SP.A.CapacityFraction = 0.0;
                SP.A.FlowFraction = FF(objSD, 0.0, SP, 'C-Off', SP.EconomizerRunning, BCC.ODB, ventilationFraction);
                SP.A.SensCap_KBtuH = 0.0;
                SP.A.LoadFraction = 0.0;
                SP.A.RunTime = 0.0;
                SP.pair_mode = 'A_only';
                return 0.0;
            }

            // Match ASP control flow:
            // - If 0<load<fullSens: iterate CF and run continuously.
            // - If load>=fullSens: run full capacity with runtime possibly > 1.
            if (load < fullSens) {
                // Initialize guess.
                let cf = load / fullSens;
                if (!Number.isFinite(cf)) cf = 0;

                // Newton iteration (ASP parity).
                for (let j = 0; j < 10; j++) {
                    // Clamp within physical bounds.
                    if (cf < 0) cf = 0;
                    if (cf > 1.0) cf = 1.0;

                    const sens0 = _vsSensAtFrac(cf);
                    const err0 = load - (Number.isFinite(sens0) ? sens0 : 0);
                    if (Math.abs(err0) < 0.001) break;

                    const sens1 = _vsSensAtFrac(cf + 0.001);
                    const err1 = load - (Number.isFinite(sens1) ? sens1 : 0);
                    const slope = (err0 - err1) / 0.001;
                    if (!Number.isFinite(slope) || Math.abs(slope) < 1e-9) break;
                    const next = cf + (err0 / slope);
                    if (Number.isFinite(next)) cf = next;
                }

                if (cf < 0) cf = 0;
                if (cf > 1.0) cf = 1.0;

                SP.A.LoadFraction = 'NA';
                SP.A.CapacityFraction = cf;
                SP.A.FlowFraction = FF(objSD, SP.A.CapacityFraction, SP, 'C-On', SP.EconomizerRunning, BCC.ODB, ventilationFraction);
                // ASP sets SensCap to the load after convergence.
                SP.A.SensCap_KBtuH = load;
                SP.A.RunTime = 1.0;

                // Don't allow capacity fraction to drop below minimum.
                if (Number.isFinite(minFrac) && minFrac > 0 && SP.A.CapacityFraction < minFrac) {
                    const rt = SP.A.CapacityFraction / minFrac;
                    const runtime = Math.max(0, Math.min(1, rt));
                    SP.A.RunTime = runtime;
                    SP.A.CapacityFraction = minFrac;
                    SP.A.FlowFraction = FF(objSD, SP.A.CapacityFraction, SP, 'C-On', SP.EconomizerRunning, BCC.ODB, ventilationFraction);
                }

                SP.pair_mode = 'A_only';
                return SP.A.RunTime;
            }

            // load >= fullSens: full capacity, runtime may exceed 1.0.
            SP.A.LoadFraction = 'NA';
            SP.A.CapacityFraction = 1.0;
            SP.A.FlowFraction = 1.0;
            SP.A.RunTime = load / fullSens;
            SP.pair_mode = 'A_only';
            return SP.A.RunTime;
        }
    }
    if (dblSensibleLoad_KBtuH < SP.A.SensCap_KBtuH || String(objSD.N_Stages) === '1') {
        SP.A.LoadFraction = dblSensibleLoad_KBtuH / SP.A.SensCap_KBtuH;
        SP.A.RunTime = SP.A.LoadFraction;
        SP.pair_mode = 'A_only';
        return SP.A.RunTime;
    }

    if (intMaxStageArrayIndex > 0) {
        for (let intStageArrayIndex = 0; intStageArrayIndex <= (intMaxStageArrayIndex - 1); intStageArrayIndex++) {
            SP.A.CapacityFraction = dblStageArray[intStageArrayIndex];
            SP.A.FlowFraction = FF(objSD, SP.A.CapacityFraction, SP, 'C-On', SP.EconomizerRunning, BCC.ODB, ventilationFraction);
            SP.A.SensCap_KBtuH = NetSenCap_Stage_Adjusted_KBtuh(objSD, totalCap, cfm, SP.A, 1, BCC.ODB, BCC.EWB, BCC.EDB, pressure);

            SP.B.CapacityFraction = dblStageArray[intStageArrayIndex + 1];
            SP.B.FlowFraction = FF(objSD, SP.B.CapacityFraction, SP, 'C-On', SP.EconomizerRunning, BCC.ODB, ventilationFraction);
            SP.BmA.FlowFraction = SP.B.FlowFraction;
            SP.B.SensCap_KBtuH = NetSenCap_Stage_Adjusted_KBtuh(objSD, totalCap, cfm, SP.B, 1, BCC.ODB, BCC.EWB, BCC.EDB, pressure);

            SP.BmA.CapacityFraction_Diff = SP.B.CapacityFraction - SP.A.CapacityFraction;
            SP.BmA.CapacityFraction = SP.B.CapacityFraction;
            SP.BmA.SensCap_KBtuH = SP.B.SensCap_KBtuH - SP.A.SensCap_KBtuH;

            if (dblSensibleLoad_KBtuH >= SP.A.SensCap_KBtuH && dblSensibleLoad_KBtuH < SP.B.SensCap_KBtuH) {
                const dblRemainingSensLoad = (dblSensibleLoad_KBtuH - SP.A.SensCap_KBtuH);
                SP.BmA.LoadFraction = dblRemainingSensLoad / SP.BmA.SensCap_KBtuH;
                SP.BmA.RunTime = SP.BmA.LoadFraction;

                SP.B.LoadFraction = 'NA';
                SP.B.RunTime = SP.BmA.RunTime;

                SP.A.LoadFraction = 1.0;
                SP.A.RunTime = 1.0 - SP.BmA.RunTime;

                SP.pair_mode = 'A_and_BmA';
                return (intStageArrayIndex + 1) + SP.B.RunTime;
            }
        }
    }

    if (dblSensibleLoad_KBtuH >= SP.B.SensCap_KBtuH) {
        SP.A.ResetToZeroLoad();
        SP.BmA.ResetToZeroLoad();

        SP.B.RunTime = dblSensibleLoad_KBtuH / SP.B.SensCap_KBtuH;
        if (SP.B.RunTime >= 1.0) {
            SP.B.LoadFraction = 1.0;
        } else {
            SP.B.LoadFraction = SP.B.RunTime;
        }
        SP.pair_mode = 'B_only';
        return intMaxStageArrayIndex + SP.B.RunTime;
    }

    SP.pair_mode = 'A_only';
    return 0;
}
