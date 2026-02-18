/**
 * Database Module - JavaScript ES Module
 * Converted from database_module.asp (VBScript)
 * 
 * Secondary data collection routines for weather and station data.
 * 
 * NOTE: This module assumes weather/station data is provided as JSON arrays
 * rather than queried from an ADO database. Data can be:
 * - Pre-loaded from static JSON files
 * - Fetched from an API endpoint
 * - Cached in IndexedDB
 * 
 * Dependencies: psychro.js
 * 
 * See ARCHITECTURE.md for how this module fits into the engine pipeline.
 */

import { Phr_wb, Pwb_hr, dblStandardPressure } from './psychro.js';

//=========================================================
// Data Structures
//=========================================================

function _normStr(s) {
    return (s ?? '').toString().trim().toUpperCase();
}

function _stationCity(station) {
    return station?.City ?? station?.city ?? '';
}

function _stationState(station) {
    return station?.state ?? station?.State ?? '';
}

function _stationTempDb(station) {
    return station?.Temp_DB ?? station?.temp_db ?? station?.TempDb ?? null;
}

function _stationElevation(station) {
    return station?.elevation ?? station?.Elevation ?? station?.elev ?? null;
}

/**
 * Weather data record structure (matches Tbins_new table columns):
 * {
 *   state: string,
 *   city: string,
 *   schedule: string,
 *   Temp_Outdoor_DB: number,    // index 3 in original 2D array
 *   Temp_Coinc_WB: number,      // index 4 in original 2D array
 *   Hours_Cooling: number
 * }
 * 
 * Station data record structure (matches stations table):
 * {
 *   City: string,
 *   state: string,
 *   Temp_DB: number,
 *   Temp_WB: number,
 *   elevation: number
 * }
 */

//=========================================================
// Weather Data Functions
//=========================================================

/**
 * Filter and sort weather records for a specific location and schedule.
 * Equivalent to GetWeatherRecords Sub.
 * 
 * @param {Array} allWeatherData - Full weather data array
 * @param {string} strStateName - State name
 * @param {string} strCityName - City name
 * @param {string} strScheduleName - Schedule name
 * @returns {Array} Filtered and sorted weather records
 */
export function getWeatherRecords(allWeatherData, strStateName, strCityName, strScheduleName) {
    const stateN = _normStr(strStateName);
    const cityN = _normStr(strCityName);
    const schedN = _normStr(strScheduleName);

    return allWeatherData
        .filter(record => 
            _normStr(record.state) === stateN && 
            _normStr(record.city) === cityN && 
            _normStr(record.schedule) === schedN
        )
        .sort((a, b) => a.Temp_Outdoor_DB - b.Temp_Outdoor_DB);
}

//=========================================================
// Design Conditions Functions
//=========================================================

/**
 * Get design conditions for a city.
 * Equivalent to GetDesignConditions Sub.
 * 
 * @param {Array} stationData - Station data array
 * @param {Array} weatherData - Pre-filtered weather data for the city/schedule
 * @param {string} strCityName - City name
 * @param {number} dblIDB_setpoint - Inside dry bulb setpoint (°F)
 * @param {boolean} blnSuppressWarnings - Suppress warning messages
 * @returns {Object} { ODB_Design, OWB_Design, Elevation_Design, Pressure_Design, warning }
 */
export function getDesignConditions(stationData, weatherData, strCityName, dblIDB_setpoint, blnSuppressWarnings = false) {
    // Find the station record for this city
    const targetCity = _normStr(strCityName);
    const station = stationData.find(s => _normStr(_stationCity(s)) === targetCity);
    
    if (!station) {
        return {
            ODB_Design: null,
            OWB_Design: null,
            Elevation_Design: null,
            Pressure_Design: null,
            warning: `Station not found for city: ${strCityName}`
        };
    }

    let dblODB_Design = Number(_stationTempDb(station));
    const dblODB_Design_previous = dblODB_Design;
    const dblElevation_Design = Number(_stationElevation(station));
    
    let blnIncreasedTheDesignTemp = false;

    // Correct for elevation...
    const dblThePressure_Design = dblStandardPressure * Math.pow((288 - (0.0065 * dblElevation_Design / 3.281)) / 288, 5.256);

    // Interpolate the HR from the weather data for the DB at design conditions.
    // This ensures the second stage is at 100% runtime at design conditions.
    
    let dblODB_BD, dblOWB_BD, dblOHR_BD;
    let dblODB_BD_previous = 0, dblOHR_BD_previous = 0;
    let blnFoundaBinHotterThanDesign = false;
    let dblOHR_Design, dblOWB_Design;

    // Loop through the weather data to find bins surrounding design DB
    for (let i = 0; i < weatherData.length; i++) {
        dblODB_BD = weatherData[i].Temp_Outdoor_DB;
        dblOWB_BD = weatherData[i].Temp_Coinc_WB;
        dblOHR_BD = Phr_wb(dblODB_BD, dblOWB_BD, dblThePressure_Design);

        if (dblODB_BD > dblODB_Design) {
            blnFoundaBinHotterThanDesign = true;
            break;
        }

        // Store the previous values for use in interpolation.
        dblODB_BD_previous = dblODB_BD;
        dblOHR_BD_previous = dblOHR_BD;
    }

    // Check for cases (usually cold climates) where you can't find any bins beyond the design condition.
    let warning = null;
    
    if (blnFoundaBinHotterThanDesign) {
        // Now interpolate
        const dblDB_fraction = (dblODB_Design - dblODB_BD_previous) / (dblODB_BD - dblODB_BD_previous);
        dblOHR_Design = dblDB_fraction * (dblOHR_BD - dblOHR_BD_previous) + dblOHR_BD_previous;

        // Set the WB using this interpolated HR value.
        dblOWB_Design = Pwb_hr(dblODB_Design, dblOHR_Design, dblThePressure_Design);
    } else {
        // If can't find the hotter bin, just use the warmest bin for the design conditions.
        dblODB_Design = dblODB_BD;
        dblOHR_Design = dblOHR_BD;
        dblOWB_Design = Pwb_hr(dblODB_Design, dblOHR_Design, dblThePressure_Design);
        blnIncreasedTheDesignTemp = true;
    }

    if (blnIncreasedTheDesignTemp && !blnSuppressWarnings) {
        warning = `Note: The design temperature for this city has been changed from ${dblODB_Design_previous.toFixed(1)} to ${dblODB_Design.toFixed(1)} to facilitate calculations at design conditions. This change is necessary for a few of the cool-summer cities.`;
    }

    return {
        ODB_Design: dblODB_Design,
        OWB_Design: dblOWB_Design,
        Elevation_Design: dblElevation_Design,
        Pressure_Design: dblThePressure_Design,
        warning: warning
    };
}

//=========================================================
// Hours Functions
//=========================================================

/**
 * Get bin hours for a schedule.
 * Equivalent to GetHours Sub.
 * 
 * @param {Array} weatherData - Full weather data array (all schedules for the city)
 * @param {string} stateName - State name
 * @param {string} cityName - City name  
 * @param {string} strSchedule - Schedule name
 * @param {string} strMode - "Occupied", "UnOccupied", "UnOccupied_ForceToZero", or "Total"
 * @returns {Map} Map of Temp_Outdoor_DB -> Hours
 */
export function getHours(weatherData, stateName, cityName, strSchedule, strMode) {
    const hoursMap = new Map();

    if (strMode === "UnOccupied" || strMode === "UnOccupied_ForceToZero") {
        if (strSchedule === "All week, All day" || strMode === "UnOccupied_ForceToZero") {
            // All hours are zero for unoccupied when schedule is "All week, All day"
            const filtered = weatherData.filter(r => 
                r.state === stateName && 
                r.city === cityName && 
                r.schedule === strSchedule
            );
            filtered.forEach(r => {
                hoursMap.set(r.Temp_Outdoor_DB, 0);
            });
        } else {
            // Calculate delta hours: "All week, All day" minus the specific schedule
            const allWeekData = weatherData.filter(r => 
                r.state === stateName && 
                r.city === cityName && 
                r.schedule === "All week, All day"
            );
            const scheduleData = weatherData.filter(r => 
                r.state === stateName && 
                r.city === cityName && 
                r.schedule === strSchedule
            );

            // Build a map of schedule hours by temp
            const scheduleHoursMap = new Map();
            scheduleData.forEach(r => {
                scheduleHoursMap.set(r.Temp_Outdoor_DB, r.Hours_Cooling);
            });

            // Calculate delta: All week hours - schedule hours
            allWeekData.forEach(r => {
                const temp = r.Temp_Outdoor_DB;
                const allWeekHours = r.Hours_Cooling;
                const schedHours = scheduleHoursMap.get(temp) || 0;
                hoursMap.set(temp, allWeekHours - schedHours);
            });
        }

    } else if (strMode === "Occupied") {
        const filtered = weatherData.filter(r => 
            r.state === stateName && 
            r.city === cityName && 
            r.schedule === strSchedule
        );
        filtered.forEach(r => {
            hoursMap.set(r.Temp_Outdoor_DB, r.Hours_Cooling);
        });

    } else if (strMode === "Total") {
        const filtered = weatherData.filter(r => 
            r.state === stateName && 
            r.city === cityName && 
            r.schedule === "All week, All day"
        );
        filtered.forEach(r => {
            hoursMap.set(r.Temp_Outdoor_DB, r.Hours_Cooling);
        });

    } else {
        console.warn("Warning: Unknown mode in getHours:", strMode);
    }

    return hoursMap;
}

/**
 * Convert hours Map to sorted array for iteration.
 * @param {Map} hoursMap - Map of temp -> hours
 * @returns {Array} Array of { temp, hours } sorted by temp ascending
 */
export function hoursMapToArray(hoursMap) {
    return Array.from(hoursMap.entries())
        .map(([temp, hours]) => ({ temp, hours }))
        .sort((a, b) => a.temp - b.temp);
}

//=========================================================
// City Name Functions
//=========================================================

/**
 * Get city name - validates selected city exists for state, returns first city if not.
 * Equivalent to GetCityName Function.
 * 
 * @param {Array} stationData - Station data array
 * @param {string} strStateNameSelected - Selected state name
 * @param {string} strCityName2Selected - Selected city name
 * @returns {string} Valid city name for the state
 */
export function getCityName(stationData, strStateNameSelected, strCityName2Selected) {
    // Filter stations for the selected state
    const stateN = _normStr(strStateNameSelected);
    const stateStations = stationData.filter(s => _normStr(_stationState(s)) === stateN);
    
    if (stateStations.length === 0) {
        console.warn("No stations found for state:", strStateNameSelected);
        return null;
    }

    const strFirstCityName = _stationCity(stateStations[0]);
    
    // Check if selected city exists in this state
    const cityN = _normStr(strCityName2Selected);
    const foundCity = stateStations.find(s => _normStr(_stationCity(s)) === cityN);
    
    if (foundCity) {
        return strCityName2Selected;
    } else {
        return strFirstCityName;
    }
}

//=========================================================
// Data Loading Utilities
//=========================================================

/**
 * Fetch weather data from a JSON file or API endpoint.
 * @param {string} url - URL to fetch data from
 * @returns {Promise<Array>} Weather data array
 */
export async function fetchWeatherData(url) {
    try {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return await response.json();
    } catch (error) {
        console.error("Error fetching weather data:", error);
        throw error;
    }
}

/**
 * Fetch station data from a JSON file or API endpoint.
 * @param {string} url - URL to fetch data from
 * @returns {Promise<Array>} Station data array
 */
export async function fetchStationData(url) {
    try {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return await response.json();
    } catch (error) {
        console.error("Error fetching station data:", error);
        throw error;
    }
}

/**
 * Get list of unique states from station data.
 * @param {Array} stationData - Station data array
 * @returns {Array<string>} Sorted array of state names
 */
export function getStateList(stationData) {
    const states = [...new Set(stationData.map(s => s.state))];
    return states.sort();
}

/**
 * Get list of cities for a state from station data.
 * @param {Array} stationData - Station data array
 * @param {string} stateName - State name
 * @returns {Array<string>} Sorted array of city names
 */
export function getCityList(stationData, stateName) {
    const cities = stationData
        .filter(s => s.state === stateName)
        .map(s => s.City);
    return cities.sort();
}

/**
 * Get list of unique schedules from weather data.
 * @param {Array} weatherData - Weather data array
 * @returns {Array<string>} Array of schedule names
 */
export function getScheduleList(weatherData) {
    const schedules = [...new Set(weatherData.map(w => w.schedule))];
    return schedules;
}
