// ---- Populate Total Capacity dropdown (36..360) ----
(function() {
   var sel = document.getElementById('cmbTotalCap');
   for (var i = 36; i <= 360; i++) {
      var opt = document.createElement('option');
      var label = (i < 100 ? '0' : '') + (i < 10 ? '0' : '') + i;
      opt.text = label;
      opt.value = label;
      if (i === 84) opt.selected = true;
      sel.add(opt);
   }
})();

// ---- Populate Degradation Factor dropdowns (0..50) ----
(function() {
   var ids = ['cmbPLDegrFactor_C', 'cmbPLDegrFactor_S'];
   for (var k = 0; k < ids.length; k++) {
      var sel = document.getElementById(ids[k]);
      if (!sel) continue;
      for (var i = 0; i <= 50; i++) {
         var opt = document.createElement('option');
         opt.text = String(i);
         opt.value = String(i);
         if (i === 25) opt.selected = true;
         sel.add(opt);
      }
   }
})();

// ---- City data by state (loaded from Stations.json) ----
var cityData = null;

async function loadCityData() {
   try {
      var resp = await fetch('data/stations.json');
      var stations = await resp.json();
      cityData = {};
      for (var i = 0; i < stations.length; i++) {
         var s = stations[i];
         var st = s.state || s.State;
         var ci = s.City || s.city;
         if (!st || !ci) continue;
         if (!cityData[st]) cityData[st] = [];
         if (cityData[st].indexOf(ci) === -1) cityData[st].push(ci);
      }
      // Sort cities within each state
      for (var st in cityData) {
         cityData[st].sort();
      }
   } catch(e) {
      console.error('Failed to load city data:', e);
   }
}

function toTitleCase(s) {
   return s.replace(/\w\S*/g, function(t) {
      return t.charAt(0).toUpperCase() + t.substr(1).toLowerCase();
   });
}

function updateCities() {
   if (!cityData) return;
   var stSel = document.getElementById('cmbState');
   var ciSel = document.getElementById('cmbCityName2');
   var st = stSel.options[stSel.selectedIndex].text;
   var cities = cityData[st] || [];
   ciSel.innerHTML = '';
   for (var i = 0; i < cities.length; i++) {
      var opt = document.createElement('option');
      opt.value = cities[i];
      opt.text = toTitleCase(cities[i]);
      ciSel.add(opt);
   }
   checkForDefaults('cmbCityName2');
}

// ---- Utility functions ----
function theSelectedValueInPullDownBox(boxname) {
   var el = document.forms.HECACParameters[boxname];
   if (!el) return '';
   for (var i = 0; i < el.length; i++) {
      if (el.options[i].selected) return el.options[i].text;
   }
   return '';
}

function selectValueInPullDownBox(boxname, valueToSelect) {
   var dB = document.forms.HECACParameters[boxname];
   if (!dB) return;
   for (var i = 0; i < dB.length; i++) {
      if (dB.options[i].text == valueToSelect) {
         dB.options[i].selected = true;
         return;
      }
   }
}

function cFD(blnInequality, strCellName) {
   var theCell = document.getElementById(strCellName);
   if (!theCell) return;
   theCell.style.backgroundColor = blnInequality ? '#3a3b3e' : '#707276';
}

function cFDW(blnInequality, strCellName) {
   var theCell = document.getElementById(strCellName);
   if (!theCell) return;
   theCell.style.backgroundColor = blnInequality ? '#FAFAD2' : '#fff';
}

function problemInPowerValue(controlName) {
   var warnings = {
      'txtBFn_kw_C':  'The candidate unit Blower Fan (BFn) field must contain a number (e.g. 1.5)',
      'txtBFn_kw_S':  'The standard unit Blower Fan (BFn) field must contain a number (e.g. 1.5)',
      'txtAux_kw_C':  'The candidate unit Auxiliary (Aux) field must contain a number (e.g. 1.5)',
      'txtAux_kw_S':  'The standard unit Auxiliary (Aux) field must contain a number (e.g. 1.5)',
      'txtCond_kw_C': 'The candidate unit Compressor (Comp) field must contain a number (e.g. 1.5)',
      'txtCond_kw_S': 'The standard unit Compressor (Comp) field must contain a number (e.g. 1.5)'
   };
   var msg = warnings[controlName];
   if (!msg) { window.alert('Warning from problemInPowerValue: Cannot find that control name.'); return true; }
   var val = parseFloat(document.getElementById(controlName).value);
   if (!isFinite(val)) { window.alert(msg); return true; }
   return false;
}

function text_notOK(theText, regExpression) {
   if ((theText.length == 0) || (theText.search(regExpression) != -1)) {
      return true;
   } else {
      return false;
   }
}

function countMatches(theText, regExpression) {
   if (theText.search(regExpression) == -1) {
      return 0;
   } else {
      return theText.match(regExpression).length;
   }
}

function checkTextBoxNumber(textBoxName, numberPeriods) {
   var theText = document.forms.HECACParameters[textBoxName].value;
   if (numberPeriods == 'None') {
      return (text_notOK(theText, /[^0-9\.]/) || (countMatches(theText, /\./g) != 0));
   } else if (numberPeriods == 'NoneOrOne') {
      return (text_notOK(theText, /[^0-9\.]/) || (countMatches(theText, /\./g) > 1));
   } else if (numberPeriods == 'One') {
      return (text_notOK(theText, /[^0-9\.]/) || (countMatches(theText, /\./g) != 1));
   } else {
      window.alert('error 1 from checkTextBoxValue');
      return false;
   }
}

function updateEER(strCorS) {
   var myForm = document.forms.HECACParameters;
   var dblCap_kbtuh = myForm['txtTotalCap'].value * 1;
   var dblBFn_kw, dblAux_kw, dblCond_kw, dblEER;
   if (strCorS === 'C') {
      dblBFn_kw  = myForm['txtBFn_kw_C'].value * 1;
      dblAux_kw  = myForm['txtAux_kw_C'].value * 1;
      dblCond_kw = myForm['txtCond_kw_C'].value * 1;
      dblEER = dblCap_kbtuh / (dblBFn_kw + dblAux_kw + dblCond_kw);
      myForm['txtEER'].value = dblEER.toFixed(2);
      checkForDefaults('txtEER');
   } else {
      dblBFn_kw  = myForm['txtBFn_kw_S'].value * 1;
      dblAux_kw  = myForm['txtAux_kw_S'].value * 1;
      dblCond_kw = myForm['txtCond_kw_S'].value * 1;
      dblEER = dblCap_kbtuh / (dblBFn_kw + dblAux_kw + dblCond_kw);
      myForm['txtEER_Standard'].value = dblEER.toFixed(2);
      checkForDefaults('txtEER_Standard');
   }
}

function updateCond_kw(strCorS) {
   var myForm = document.forms.HECACParameters;
   var dblCap_kbtuh = myForm['txtTotalCap'].value * 1;
   var dblBFn_kw, dblAux_kw, dblCond_kw, dblEER;
   if (strCorS === 'C') {
      dblBFn_kw = myForm['txtBFn_kw_C'].value * 1;
      dblAux_kw = myForm['txtAux_kw_C'].value * 1;
      dblEER    = myForm['txtEER'].value * 1;
      dblCond_kw = (dblCap_kbtuh / dblEER) - (dblBFn_kw + dblAux_kw);
      myForm['txtCond_kw_C'].value = dblCond_kw.toFixed(3);
      document.getElementById('tdCond_kw_C_default').textContent = dblCond_kw.toFixed(3);
      checkForDefaults('txtCond_kw_C');
   } else {
      dblBFn_kw = myForm['txtBFn_kw_S'].value * 1;
      dblAux_kw = myForm['txtAux_kw_S'].value * 1;
      dblEER    = myForm['txtEER_Standard'].value * 1;
      dblCond_kw = (dblCap_kbtuh / dblEER) - (dblBFn_kw + dblAux_kw);
      myForm['txtCond_kw_S'].value = dblCond_kw.toFixed(3);
      document.getElementById('tdCond_kw_S_default').textContent = dblCond_kw.toFixed(3);
      checkForDefaults('txtCond_kw_S');
   }
}

// Lightweight client-side parse of spreadsheet data to extract key fields.
// Validate spreadsheet data format (matches ASP classes.asp ParseAndModel lines 690-713).
// Returns an error message string if invalid, or '' if OK.
function validateSpreadsheetData(ssText, unitName) {
   if (!ssText) return 'WARNING: Spreadsheet data for the ' + unitName + ' unit is empty.';
   var norm = ssText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
   var rows = norm.split('\n');
   if (rows.length > 1 && rows.length < 40) {
      var versionRow = rows[1] || '';
      if (versionRow.indexOf('V1.2') === -1) {
         return 'WARNING: Spreadsheet data for the ' + unitName + ' unit is not in the correct form or the spreadsheet version is not V1.2.';
      }
   } else {
      if (rows.length > 39) {
         return 'WARNING: Spreadsheet data for the ' + unitName + ' unit is not in the correct form. It looks like maybe you pasted the data more than once. Clear out the spreadsheet cell and try again.';
      } else {
         return 'WARNING: Spreadsheet data for the ' + unitName + ' unit is not in the correct form. Clear out the spreadsheet cell and try again.';
      }
   }
   return '';
}

// Matches ASP ParseSpreadsheet sub (Controls.asp line 735).
function parseSpreadsheetFields(ssText) {
   var result = {};
   if (!ssText) return result;
   var rows = ssText.split(/\r\n|\n|\r/);
   for (var j = 0; j < rows.length; j++) {
      var cells = rows[j].split('\t');
      var key = (cells[0] || '').trim();
      if (key === 'NetCoolingCapacity') result.netCoolingCapacity = parseFloat(cells[2]);
      else if (key === 'EER') result.eer = parseFloat(cells[2]);
      else if (key === 'EvaporatorFanPower') result.evapFanPower = parseFloat(cells[2]);
      else if (key === 'AuxilaryPower' || key === 'AuxiliaryPower') result.auxPower = parseFloat(cells[2]);
      else if (key === 'CondenserPower' || key === 'CondensorPower') result.condenserPower = parseFloat(cells[2]);
      else if (key === 'ST_Ratio') result.stRatio = parseFloat(cells[2]);
   }
   return result;
}

// Update form fields from parsed spreadsheet data (matches ASP ParseSpreadsheet).
// strCorS = 'C' for candidate, 'S' for standard.
function applySpreadsheetToForm(strCorS, ssFields) {
   var form = document.forms.HECACParameters;
   if (!ssFields || !isFinite(ssFields.netCoolingCapacity)) return;

   // Update total capacity (common to both units)
   var capKBtuh = Math.round(ssFields.netCoolingCapacity / 1000);
   form['txtTotalCap'].value = capKBtuh;
   selectValueInPullDownBox('cmbTotalCap', String(capKBtuh).padStart(3, '0'));
   checkForDefaults('cmbTotalCap');

   // Default power calculation constants (match ASP FanPower_Default / Condenser_Power_kW)
   var dblBFn_slope = 0.0132;
   var dblBFn_int   = -0.2283;
   var dblDefaultBFn = dblBFn_slope * capKBtuh + dblBFn_int;
   var dblDefaultAux = 0;

   // Unit-specific power fields and EER
   if (strCorS === 'C') {
      if (isFinite(ssFields.evapFanPower)) form['txtBFn_kw_C'].value = ssFields.evapFanPower.toFixed(3);
      if (isFinite(ssFields.auxPower)) form['txtAux_kw_C'].value = ssFields.auxPower.toFixed(3);
      if (isFinite(ssFields.condenserPower)) form['txtCond_kw_C'].value = ssFields.condenserPower.toFixed(3);
      if (isFinite(ssFields.eer)) form['txtEER'].value = ssFields.eer.toFixed(1);
      if (isFinite(ssFields.stRatio)) {
         var stVal = Math.round(ssFields.stRatio * 100) / 100;
         selectValueInPullDownBox('cmbST_Ratio_C', stVal.toFixed(2));
         checkForDefaults('cmbST_Ratio_C');
      }
      // If standard unit has no spreadsheet, recalc ONLY its power from new capacity
      if (!document.getElementById('chkSpreadsheetControl_S').checked) {
         var eerS = parseFloat(form['txtEER_Standard'].value);
         var condS = (capKBtuh / eerS) - (dblDefaultBFn + dblDefaultAux);
         form['txtBFn_kw_S'].value = dblDefaultBFn.toFixed(3);
         form['txtAux_kw_S'].value = dblDefaultAux.toFixed(3);
         form['txtCond_kw_S'].value = condS.toFixed(3);
      }
   } else {
      if (isFinite(ssFields.evapFanPower)) form['txtBFn_kw_S'].value = ssFields.evapFanPower.toFixed(3);
      if (isFinite(ssFields.auxPower)) form['txtAux_kw_S'].value = ssFields.auxPower.toFixed(3);
      if (isFinite(ssFields.condenserPower)) form['txtCond_kw_S'].value = ssFields.condenserPower.toFixed(3);
      if (isFinite(ssFields.eer)) form['txtEER_Standard'].value = ssFields.eer.toFixed(1);
      if (isFinite(ssFields.stRatio)) {
         var stVal = Math.round(ssFields.stRatio * 100) / 100;
         selectValueInPullDownBox('cmbST_Ratio_S', stVal.toFixed(2));
         checkForDefaults('cmbST_Ratio_S');
      }
      // If candidate unit has no spreadsheet, recalc ONLY its power from new capacity
      if (!document.getElementById('chkSpreadsheetControl_C').checked) {
         var eerC = parseFloat(form['txtEER'].value);
         var condC = (capKBtuh / eerC) - (dblDefaultBFn + dblDefaultAux);
         form['txtBFn_kw_C'].value = dblDefaultBFn.toFixed(3);
         form['txtAux_kw_C'].value = dblDefaultAux.toFixed(3);
         form['txtCond_kw_C'].value = condC.toFixed(3);
      }
   }

   // Update defaults display and colors
   setFanPowerValuesAndDefaults(capKBtuh, 'OnlyDefaults');
   checkForDefaults('txtEER');
   checkForDefaults('txtEER_Standard');
   checkForDefaults('txtBFn_kw_C');
   checkForDefaults('txtAux_kw_C');
   checkForDefaults('txtCond_kw_C');
   checkForDefaults('txtBFn_kw_S');
   checkForDefaults('txtAux_kw_S');
   checkForDefaults('txtCond_kw_S');
}

// Recalculate ventilation rate after capacity changes (e.g. from spreadsheet data).
// Matches ASP behavior where Establish_SIV runs during InitializeControls.
async function recalcVentilation() {
   try {
      var form = document.forms.HECACParameters;
      var psychro = await import('./engine/psychro.js');
      var perf = await import('./engine/performance_module.js');
      var engine = await import('./engine/engine_module.js?v=14');
      if (!engine || typeof engine.exportBinCalcsJson !== 'function') return;

      // Iterative Establish_SIV matching ASP Controls.asp lines 405-431:
      // Run engine â†’ extract sensibleCapacityDesign â†’ refine ventilation â†’ repeat
      var lockWasChecked = form['chkLockLoadLine'] && form['chkLockLoadLine'].checked;
      if (lockWasChecked) form['chkLockLoadLine'].checked = false;
      var ranClean = false;
      var bpfErrorMsg = '';
      for (var iter = 0; iter < 4; iter++) {
         try { await engine.exportBinCalcsJson(form, {}); } catch (ex) {
            // Capture BPF/ADP error detail from the engine exception
            var exMsg = ex && ex.message ? ex.message : '';
            var bpfMatch = exMsg.match(/Engine error:\s*(.*)/i);
            bpfErrorMsg = bpfMatch ? bpfMatch[1].trim() : exMsg;
            break;
         }
         var ll = engine.getLastLoadLine();
         if (!ll || !ll.debug || !ll.debug.capacity || !ll.debug.design) break;
         var sensCapDesign = ll.debug.capacity.sensibleCapacityDesign;
         if (!Number.isFinite(sensCapDesign) || sensCapDesign <= 0) break;
         ranClean = true;
         var prevVent = parseFloat(form['txtVentilationValue_NotAdvanced'].value);
         await refineVentilationFromCapacity(psychro, perf, sensCapDesign, ll.debug.design);
         var newVent = parseFloat(form['txtVentilationValue_NotAdvanced'].value);
         if (Math.abs(newVent - prevVent) <= 0.1) break;
      }
      if (lockWasChecked) form['chkLockLoadLine'].checked = true;
      // Show or clear the BPF warning area (matches ASP psychro.asp lines 368-370).
      var warnArea = document.getElementById('bpfWarningArea');
      if (!ranClean) {
         // ASP fallback (Controls.asp lines 442-449): when BPF/ADP fails (e.g. S/T=0.80)
         // Establish_SIV sets blnRanClean=False and ventilation falls back to 10%.
         var ventUnits = theSelectedValueInPullDownBox('cmbVentilationUnits') || '';
         var fallback;
         if (ventUnits.toUpperCase().indexOf('CFM') >= 0) {
            var cap = parseFloat(form['txtTotalCap'].value) || 84;
            fallback = String(Math.round(0.10 * (cap / 12) * 400));
         } else {
            fallback = '10';
         }
         form['txtVentilationValue_NotAdvanced'].value = fallback;
         form['txtVentilationValue'].value = fallback;
         if (warnArea) {
            var unitName = 'Candidate';
            if (bpfErrorMsg && bpfErrorMsg.indexOf('Standard Unit') >= 0) unitName = 'Standard';
            var cleanMsg = bpfErrorMsg.replace(/^(Candidate|Standard) Unit BPF:\s*/i, '');
            warnArea.innerHTML = 'WARNING: Try lowering the S/T ratio for the ' + unitName + ' Unit.<br>'
               + 'ERROR Message from BPF calculation: ' + (cleanMsg || 'BPF/ADP calculation failed.');
            warnArea.style.display = '';
         }
      } else {
         if (warnArea) { warnArea.innerHTML = ''; warnArea.style.display = 'none'; }
      }
      // Update the default display to match the computed ventilation (ASP does this
      // because the server re-renders the page with intVentilationValue in both the
      // input and the default TD â€” see Controls.asp lines 1351 and 1355).
      var ventDef = document.getElementById('tdVentilationValue_default');
      if (ventDef) ventDef.textContent = form['txtVentilationValue'].value;
      checkForDefaults('txtVentilationValue');
   } catch (e) {
      console.warn('recalcVentilation failed:', e);
   }
}

function setFanPowerValuesAndDefaults(dblCapacityTotal_kBtuh, strMode) {
   var dblBFn_slope_kW_per_kBtuh = 0.0132;
   var dblBFn_int_kW = -0.2283;
   var dblBFanPower, dblAuxPower, dblCondPower_C, dblCondPower_S;

   dblBFanPower = dblBFn_slope_kW_per_kBtuh * dblCapacityTotal_kBtuh + dblBFn_int_kW;
   dblAuxPower = 0;
   dblCondPower_C = (dblCapacityTotal_kBtuh / document.forms.HECACParameters['txtEER'].value) - (dblBFanPower + dblAuxPower);
   dblCondPower_S = (dblCapacityTotal_kBtuh / document.forms.HECACParameters['txtEER_Standard'].value) - (dblBFanPower + dblAuxPower);

   document.getElementById('tdBFn_kw_C_default').textContent = dblBFanPower.toFixed(3);
   document.getElementById('tdAux_kw_C_default').textContent = dblAuxPower.toFixed(3);
   document.getElementById('tdCond_kw_C_default').textContent = dblCondPower_C.toFixed(3);

   document.getElementById('tdBFn_kw_S_default').textContent = dblBFanPower.toFixed(3);
   document.getElementById('tdAux_kw_S_default').textContent = dblAuxPower.toFixed(3);
   document.getElementById('tdCond_kw_S_default').textContent = dblCondPower_S.toFixed(3);

   if (strMode === 'AlsoFormValues') {
      document.forms.HECACParameters['txtBFn_kw_C'].value = dblBFanPower.toFixed(3);
      document.forms.HECACParameters['txtAux_kw_C'].value = dblAuxPower.toFixed(3);
      document.forms.HECACParameters['txtCond_kw_C'].value = dblCondPower_C.toFixed(3);

      document.forms.HECACParameters['txtBFn_kw_S'].value = dblBFanPower.toFixed(3);
      document.forms.HECACParameters['txtAux_kw_S'].value = dblAuxPower.toFixed(3);
      document.forms.HECACParameters['txtCond_kw_S'].value = dblCondPower_S.toFixed(3);
   }

   checkForDefaults('txtBFn_kw_C');
   checkForDefaults('txtAux_kw_C');
   checkForDefaults('txtCond_kw_C');
   checkForDefaults('txtBFn_kw_S');
   checkForDefaults('txtAux_kw_S');
   checkForDefaults('txtCond_kw_S');
}

function resetPowerDefaults() {
   setFanPowerValuesAndDefaults(document.forms.HECACParameters['txtTotalCap'].value, 'AlsoFormValues');
   var chkC = document.getElementById('chkSpreadsheetControl_C');
   var chkS = document.getElementById('chkSpreadsheetControl_S');
   if (chkC) chkC.checked = false;
   if (chkS) chkS.checked = false;
}

function checkForDefaults(controlName) {
   var sVPDB = theSelectedValueInPullDownBox;
   var form = document.forms.HECACParameters;

   if (controlName === 'cmbBuildingType') {
      cFD(sVPDB(controlName) !== 'Office-Medium', 'tdBuildingType_default');
   } else if (controlName === 'cmbState') {
      cFD(sVPDB(controlName) !== 'MO', 'tdState_default');
   } else if (controlName === 'cmbCityName2') {
      cFD(sVPDB(controlName) !== 'Kansas City', 'tdCity_default');
   } else if (controlName === 'cmbSchedule') {
      cFD(sVPDB(controlName) !== 'M-Fri, 7 a.m. to 7 p.m.', 'tdSchedule_default');
   } else if (controlName === 'cmbIDB') {
      cFD(sVPDB(controlName) !== '75', 'tdIDB_default');
   } else if (controlName === 'cmbIDB_SetBack') {
      cFD(sVPDB(controlName) !== '5', 'tdIDB_SetBack_default');
   } else if (controlName === 'cmbTotalCap') {
      cFD(sVPDB(controlName) !== '084', 'tdTotalCapacity_default');
   } else if (controlName === 'cmbOversizePercent') {
      cFD(sVPDB(controlName) != 0, 'tdOversizePercent_default');
   } else if (controlName === 'txtEER') {
      cFD(form[controlName].value != 12, 'tdEER_default');
   } else if (controlName === 'txtUnitCost') {
      cFD(form[controlName].value != 4.5, 'tdUnitCost_default');
   } else if (controlName === 'txtEER_Standard') {
      cFD(form[controlName].value != 9.0, 'tdEER_Standard_default');
   } else if (controlName === 'txtUnitCost_Standard') {
      cFD(form[controlName].value != 4.0, 'tdUnitCost_Standard_default');
   } else if (controlName === 'txtMaintenance_Standard') {
      cFD(form[controlName].value != 0.0, 'tdMaintenance_Standard_default');
   } else if (controlName === 'txtMaintenance_Candidate') {
      cFD(form[controlName].value != 0.0, 'tdMaintenance_Candidate_default');
   } else if (controlName === 'chkEconomizer_C') {
      cFD(form[controlName].checked !== true, 'tdEconomizer_C_default');
   } else if (controlName === 'chkEconomizer_S') {
      cFD(form[controlName].checked !== true, 'tdEconomizer_S_default');
   } else if (controlName === 'txtElectricityRate') {
      cFD(form[controlName].value != 0.08, 'tdElectricityRate_default');
   } else if (controlName === 'txtDiscountRate') {
      cFD(form[controlName].value != 7.0, 'tdDiscountRate_default');
   } else if (controlName === 'cmbEquipmentLife') {
      cFD(sVPDB('cmbEquipmentLife') != 15, 'tdEquipmentLife_default');
   } else if (controlName === 'txtNUnits') {
      cFD(form[controlName].value != 1.0, 'tdNUnits_default');
   } else if (controlName === 'chkChartPW') {
      cFD(form[controlName].checked !== true, 'tdChartPW_default');
   } else if (controlName === 'chkDetails') {
      cFD(form[controlName].checked === true, 'tdDetails_default');
   } else if (controlName === 'chkAdvancedControls') {
      cFD(form[controlName].checked === true, 'tdAdvancedControls_default');
   } else if (controlName === 'chkLockLoadLine') {
      cFD(form[controlName].checked === true, 'tdLockLoadLine_default');
   } else if (controlName === 'cmbST_Ratio_C') {
      cFD(sVPDB(controlName) !== '0.72', 'tdST_Ratio_default_C');
   } else if (controlName === 'cmbST_Ratio_S') {
      cFD(sVPDB(controlName) !== '0.72', 'tdST_Ratio_default_S');
   } else if (controlName === 'txtBFn_kw_C') {
      var def = document.getElementById('tdBFn_kw_C_default');
      if (def) cFD(form[controlName].value !== def.textContent.trim(), 'tdBFn_kw_C_default');
   } else if (controlName === 'txtAux_kw_C') {
      var def = document.getElementById('tdAux_kw_C_default');
      if (def) cFD(form[controlName].value !== def.textContent.trim(), 'tdAux_kw_C_default');
   } else if (controlName === 'txtCond_kw_C') {
      var def = document.getElementById('tdCond_kw_C_default');
      if (def) cFD(form[controlName].value !== def.textContent.trim(), 'tdCond_kw_C_default');
   } else if (controlName === 'txtBFn_kw_S') {
      var def = document.getElementById('tdBFn_kw_S_default');
      if (def) cFD(form[controlName].value !== def.textContent.trim(), 'tdBFn_kw_S_default');
   } else if (controlName === 'txtAux_kw_S') {
      var def = document.getElementById('tdAux_kw_S_default');
      if (def) cFD(form[controlName].value !== def.textContent.trim(), 'tdAux_kw_S_default');
   } else if (controlName === 'txtCond_kw_S') {
      var def = document.getElementById('tdCond_kw_S_default');
      if (def) cFD(form[controlName].value !== def.textContent.trim(), 'tdCond_kw_S_default');
   } else if (controlName === 'txtVentilationValue') {
      var def = document.getElementById('tdVentilationValue_default');
      if (def) cFD(form[controlName].value !== def.textContent.trim(), 'tdVentilationValue_default');
   } else if (controlName === 'cmbVentilationUnits') {
      cFD(sVPDB('cmbVentilationUnits') !== '% of Fan Cap.', 'tdVentilationUnits_default');
   } else if (controlName === 'cmbN_Affinity') {
      cFD(sVPDB(controlName) !== '2.5', 'tdN_Affinity_default');
   } else if (controlName === 'txtCondFanPercent_C') {
      cFD(form[controlName].value != 9.0, 'tdCondFanPercent_C');
   } else if (controlName === 'txtCondFanPercent_S') {
      cFD(form[controlName].value != 9.0, 'tdCondFanPercent_S');
   } else if (controlName === 'chkTrackOHR') {
      cFD(form[controlName].checked !== true, 'tdTrackOHR_default');
   } else if (controlName === 'cmbIRH') {
      cFD(sVPDB('cmbIRH') !== '60', 'tdIRH_default');
   } else if (controlName === 'cmbNstages_C') {
      cFD(sVPDB(controlName) !== '1', 'tdNStages_default_C');
   } else if (controlName === 'cmbNstages_S') {
      cFD(sVPDB(controlName) !== '1', 'tdNStages_default_S');
   } else if (controlName === 'cmbFanControls_C') {
      cFD(sVPDB(controlName) !== '1-Spd: Always ON', 'tdFanControls_C_default');
   } else if (controlName === 'cmbFanControls_S') {
      cFD(sVPDB(controlName) !== '1-Spd: Always ON', 'tdFanControls_S_default');
   } else if (controlName === 'cmbPLDegrFactor_C') {
      cFD(sVPDB(controlName) !== '25', 'tdPLDegrFactor_C_default');
   } else if (controlName === 'cmbPLDegrFactor_S') {
      cFD(sVPDB(controlName) !== '25', 'tdPLDegrFactor_S_default');
   } else if (controlName === 'cmbSpecific_RTU_C') {
      cFD(sVPDB(controlName) !== 'None', 'tdSpecific_RTU_C_default');
   } else if (controlName === 'cmbDemandMonths') {
      cFD(sVPDB(controlName) !== '0', 'tdDemandMonths_default');
   } else if (controlName === 'txtDemandCostPerKW') {
      cFD(form[controlName].value != 0, 'tdDemandCostPerKW_default');
   } else if (controlName === 'txtSpreadsheetData_C') {
      cFD((form[controlName].value !== ''), 'tdSpreadsheetData_C_default');
      var ssChecked_C = document.getElementById('chkSpreadsheetControl_C').checked;
      cFDW(ssChecked_C, 'tdPLDegrFactor_C_control');
      cFDW(ssChecked_C, 'tdST_Ratio_C_control');
   } else if (controlName === 'txtSpreadsheetData_S') {
      cFD((form[controlName].value !== ''), 'tdSpreadsheetData_S_default');
      var ssChecked_S = document.getElementById('chkSpreadsheetControl_S').checked;
      cFDW(ssChecked_S, 'tdPLDegrFactor_S_control');
      cFDW(ssChecked_S, 'tdST_Ratio_S_control');
   }
}

function onChangeHandler(controlName) {
   if (controlName === 'cmbTotalCap') {
      var strTotalCap = theSelectedValueInPullDownBox('cmbTotalCap');
      var intTotalCap = parseInt(strTotalCap, 10);
      document.forms.HECACParameters['txtTotalCap'].value = intTotalCap;
      if (document.forms.HECACParameters['chkAdvancedControls'].checked) {
         setFanPowerValuesAndDefaults(intTotalCap, 'AlsoFormValues');
      }
      checkForDefaults('cmbTotalCap');
      checkForDefaults('txtEER');
      checkForDefaults('txtUnitCost');
      recalcVentilation(); // ASP: SubmitTheFormToItself â†’ Establish_SIV
   } else if (controlName === 'txtDiscountRate') {
      document.getElementById('txtDiscountRate_hidden').value = document.getElementById('txtDiscountRate').value;
   } else if (controlName === 'chkChartPW') {
      document.getElementById('txtDiscountRate').disabled = !document.getElementById('chkChartPW').checked;
   } else if (controlName === 'cmbBuildingType') {
      document.forms.HECACParameters['txtBuildingType_hidden'].value = theSelectedValueInPullDownBox(controlName);
      checkForDefaults('cmbBuildingType');
      // Repopulate BM fields if Advanced is ON
      if (document.getElementById('chkAdvancedControls').checked) {
         populateBMFieldsFromBuildingType();
      }
      recalcVentilation(); // ASP: GeneralChangeHandlerANDsubmitToSelf â†’ Establish_SIV
   } else if (controlName === 'txtBM_Slope') {
      var v = parseFloat(document.getElementById('txtBM_Slope').value);
      if (!isFinite(v)) {
         window.alert('The Building Model slope must be a number. The slope will be reset to 0.0.');
         document.getElementById('txtBM_Slope').value = 0;
      }
      cFDW(true, 'tdBuildingType_controlscell');
      document.getElementById('txtBM_Slope_hidden').value = document.getElementById('txtBM_Slope').value;
      document.getElementById('txtBM_postingstate').value = 'Pend';
   } else if (controlName === 'txtBM_Intercept') {
      var v = parseFloat(document.getElementById('txtBM_Intercept').value);
      if (!isFinite(v)) {
         window.alert('The Building Model intercept must be a number. The intercept will be reset to 0.0.');
         document.getElementById('txtBM_Intercept').value = 0;
      }
      cFDW(true, 'tdBuildingType_controlscell');
      document.getElementById('txtBM_Intercept_hidden').value = document.getElementById('txtBM_Intercept').value;
      document.getElementById('txtBM_postingstate').value = 'Pend';
   } else if (controlName === 'txtBM_VentSlopeFraction') {
      var v = parseFloat(document.getElementById('txtBM_VentSlopeFraction').value);
      if (!isFinite(v)) {
         window.alert('The Building Model ventilation-slope fraction must be a number.');
         document.getElementById('txtBM_VentSlopeFraction').value = 0.50;
      } else if (v >= 1) {
         window.alert('The Building Model ventilation-slope fraction must be less than 1.');
         document.getElementById('txtBM_VentSlopeFraction').value = 0.99;
      }
      cFDW(true, 'tdBuildingType_controlscell');
      document.getElementById('txtBM_VentSlopeFraction_hidden').value = document.getElementById('txtBM_VentSlopeFraction').value;
      document.getElementById('txtBM_postingstate').value = 'Pend';

   //Fan and other power related input controls (candidate unit)
   } else if (controlName === 'txtBFn_kw_C' || controlName === 'txtAux_kw_C') {
      if (!problemInPowerValue(controlName)) {
         updateEER('C');
         setFanPowerValuesAndDefaults(document.forms.HECACParameters['txtTotalCap'].value, 'OnlyDefaults');
      }
   } else if (controlName === 'txtCond_kw_C') {
      if (!problemInPowerValue(controlName)) {
         updateEER('C');
         document.getElementById('tdCond_kw_C_default').textContent = document.forms.HECACParameters['txtCond_kw_C'].value;
         checkForDefaults('txtCond_kw_C');
      }
   } else if (controlName === 'txtEER') {
      updateCond_kw('C');

   //Fan and other power related input controls (standard unit)
   } else if (controlName === 'txtBFn_kw_S' || controlName === 'txtAux_kw_S') {
      if (!problemInPowerValue(controlName)) {
         updateEER('S');
         setFanPowerValuesAndDefaults(document.forms.HECACParameters['txtTotalCap'].value, 'OnlyDefaults');
      }
   } else if (controlName === 'txtCond_kw_S') {
      if (!problemInPowerValue(controlName)) {
         updateEER('S');
         document.getElementById('tdCond_kw_S_default').textContent = document.forms.HECACParameters['txtCond_kw_S'].value;
         checkForDefaults('txtCond_kw_S');
      }
   } else if (controlName === 'txtEER_Standard') {
      updateCond_kw('S');

   // Auto humidity checkbox â€” toggle IRH dropdown enabled/disabled (ASP does this via SubmitTheFormToItself)
   } else if (controlName === 'chkTrackOHR') {
      var autoChecked = document.forms.HECACParameters['chkTrackOHR'].checked;
      document.getElementById('cmbIRH').disabled = autoChecked;
      recalcVentilation();
   } else if (controlName === 'cmbIRH') {
      // Keep hidden field in sync (ASP reads from cmbIRH or txtIRH_hidden)
      document.getElementById('txtIRH_hidden').value = theSelectedValueInPullDownBox('cmbIRH');
      recalcVentilation();
   } else if (controlName === 'cmbST_Ratio_C' || controlName === 'cmbST_Ratio_S') {
      recalcVentilation();

   //These next two blocks force the degradation factor to zero when variable speed mode is selected.
   } else if (controlName === 'cmbFanControls_C') {
      var theSelectedValue = theSelectedValueInPullDownBox('cmbFanControls_C');
      if (theSelectedValue.substring(0,1) === 'V') {
         //First, save the value in the hidden field.
         if (document.getElementById('cmbPLDegrFactor_C').disabled !== true) {
            document.forms.HECACParameters['txtPLDegrFactor_C'].value = theSelectedValueInPullDownBox('cmbPLDegrFactor_C');
         }
         selectValueInPullDownBox('cmbPLDegrFactor_C', 0);
         document.getElementById('cmbPLDegrFactor_C').disabled = true;
         document.getElementById('txtCondFanPercent_C').disabled = false;
      } else {
         //Use the hidden value to reset the pull down control
         selectValueInPullDownBox('cmbPLDegrFactor_C', document.forms.HECACParameters['txtPLDegrFactor_C'].value);
         document.getElementById('cmbPLDegrFactor_C').disabled = false;
         document.getElementById('txtCondFanPercent_C').disabled = true;
      }
      checkForDefaults('cmbFanControls_C');
      checkForDefaults('cmbPLDegrFactor_C');
   } else if (controlName === 'cmbPLDegrFactor_C') {
      // stash away the selected value in the hidden field.
      document.forms.HECACParameters['txtPLDegrFactor_C'].value = theSelectedValueInPullDownBox('cmbPLDegrFactor_C');
      checkForDefaults(controlName);

   } else if (controlName === 'cmbFanControls_S') {
      if (theSelectedValueInPullDownBox('cmbFanControls_S').substring(0,1) === 'V') {
         //First, save the value in the hidden field.
         if (document.getElementById('cmbPLDegrFactor_S').disabled !== true) {
            document.forms.HECACParameters['txtPLDegrFactor_S'].value = theSelectedValueInPullDownBox('cmbPLDegrFactor_S');
         }
         selectValueInPullDownBox('cmbPLDegrFactor_S', 0);
         document.getElementById('cmbPLDegrFactor_S').disabled = true;
         document.getElementById('txtCondFanPercent_S').disabled = false;
      } else {
         //Use the hidden value to reset the pull down control
         selectValueInPullDownBox('cmbPLDegrFactor_S', document.forms.HECACParameters['txtPLDegrFactor_S'].value);
         document.getElementById('cmbPLDegrFactor_S').disabled = false;
         document.getElementById('txtCondFanPercent_S').disabled = true;
      }
      checkForDefaults('cmbFanControls_S');
      checkForDefaults('cmbPLDegrFactor_S');
   } else if (controlName === 'cmbPLDegrFactor_S') {
      // stash away the selected value in the hidden field.
      document.forms.HECACParameters['txtPLDegrFactor_S'].value = theSelectedValueInPullDownBox('cmbPLDegrFactor_S');
      checkForDefaults(controlName);

   } else if (controlName === 'txtCondFanPercent_C') {
      document.forms.HECACParameters['txtCondFanPercent_C_hidden'].value = document.forms.HECACParameters['txtCondFanPercent_C'].value;
   } else if (controlName === 'txtCondFanPercent_S') {
      document.forms.HECACParameters['txtCondFanPercent_S_hidden'].value = document.forms.HECACParameters['txtCondFanPercent_S'].value;

   } else if (controlName === 'cmbNstages_C') {
      document.forms.HECACParameters['txtNstages_C'].value = theSelectedValueInPullDownBox(controlName);
      checkForDefaults(controlName);
   } else if (controlName === 'cmbNstages_S') {
      document.forms.HECACParameters['txtNstages_S'].value = theSelectedValueInPullDownBox(controlName);
      checkForDefaults(controlName);

   } else if (controlName === 'cmbSpecific_RTU_C') {
      if (theSelectedValueInPullDownBox(controlName) === 'Three Stages') {
         //Set candidate stages control to 3.
         document.forms.HECACParameters['txtNstages_C'].value = '3';
         selectValueInPullDownBox('cmbNstages_C', '3');
         checkForDefaults('cmbNstages_C');
         //Set fan controls
         selectValueInPullDownBox('cmbFanControls_C', 'N-Spd: Always ON');
         checkForDefaults('cmbFanControls_C');
         //Use the hidden value to reset the pull down control for the degradation factor.
         selectValueInPullDownBox('cmbPLDegrFactor_C', document.forms.HECACParameters['txtPLDegrFactor_C'].value);
         document.getElementById('cmbPLDegrFactor_C').disabled = false;
         document.getElementById('txtCondFanPercent_C').disabled = true;

      } else if (theSelectedValueInPullDownBox(controlName) === 'Advanced Controls') {
         //Set candidate stages control to 2.
         document.forms.HECACParameters['txtNstages_C'].value = '2';
         selectValueInPullDownBox('cmbNstages_C', '2');
         checkForDefaults('cmbNstages_C');
         //Set fan controls
         selectValueInPullDownBox('cmbFanControls_C', '1-Spd: Always ON');
         checkForDefaults('cmbFanControls_C');
         //Use the hidden value to reset the pull down control for the degradation factor.
         selectValueInPullDownBox('cmbPLDegrFactor_C', document.forms.HECACParameters['txtPLDegrFactor_C'].value);
         document.getElementById('cmbPLDegrFactor_C').disabled = false;
         document.getElementById('txtCondFanPercent_C').disabled = true;

      } else if (theSelectedValueInPullDownBox(controlName) === 'Variable-Speed Compressor') {
         //Set candidate stages control to 1.
         document.forms.HECACParameters['txtNstages_C'].value = '1';
         selectValueInPullDownBox('cmbNstages_C', '1');
         checkForDefaults('cmbNstages_C');
         //Set fan controls
         selectValueInPullDownBox('cmbFanControls_C', 'V-Spd: Always ON');
         checkForDefaults('cmbFanControls_C');
         //The following block is used to disable the degradation factor, but first save the current value.
         if (document.getElementById('cmbPLDegrFactor_C').disabled !== true) {
            document.forms.HECACParameters['txtPLDegrFactor_C'].value = theSelectedValueInPullDownBox('cmbPLDegrFactor_C');
         }
         selectValueInPullDownBox('cmbPLDegrFactor_C', 0);
         document.getElementById('cmbPLDegrFactor_C').disabled = true;
         document.getElementById('txtCondFanPercent_C').disabled = false;

      } else if (theSelectedValueInPullDownBox(controlName) === 'None') {
         //Set candidate stages control to 1.
         document.forms.HECACParameters['txtNstages_C'].value = '1';
         selectValueInPullDownBox('cmbNstages_C', '1');
         checkForDefaults('cmbNstages_C');
         //Set fan controls
         selectValueInPullDownBox('cmbFanControls_C', '1-Spd: Always ON');
         checkForDefaults('cmbFanControls_C');
         //Use the hidden value to reset the pull down control for the degradation factor.
         selectValueInPullDownBox('cmbPLDegrFactor_C', document.forms.HECACParameters['txtPLDegrFactor_C'].value);
         document.getElementById('cmbPLDegrFactor_C').disabled = false;
         document.getElementById('txtCondFanPercent_C').disabled = true;
      }
      checkForDefaults(controlName);
      checkForDefaults('cmbPLDegrFactor_C');

   // Spreadsheet data textareas â€” just update default colors; user must check box manually
   } else if (controlName === 'txtSpreadsheetData_C') {
      // If textarea is cleared, uncheck the box
      if (document.getElementById('txtSpreadsheetData_C').value === '') {
         document.getElementById('chkSpreadsheetControl_C').checked = false;
      }
   } else if (controlName === 'txtSpreadsheetData_S') {
      if (document.getElementById('txtSpreadsheetData_S').value === '') {
         document.getElementById('chkSpreadsheetControl_S').checked = false;
      }

   // Spreadsheet data checkboxes
   } else if (controlName === 'chkSpreadsheetControl_C') {
      var chk = document.getElementById('chkSpreadsheetControl_C');
      var ta = document.getElementById('txtSpreadsheetData_C');
      if (chk.checked) {
         if (ta.value === '') {
            chk.checked = false;
         } else {
            var errMsg = validateSpreadsheetData(ta.value, 'Candidate');
            var wa = document.getElementById('bpfWarningArea');
            if (errMsg) {
               if (wa) { wa.innerHTML = errMsg; wa.style.display = ''; }
               chk.checked = false;
            } else {
               if (wa) { wa.innerHTML = ''; wa.style.display = 'none'; }
               // Parse and apply spreadsheet data to form fields
               var ssFields = parseSpreadsheetFields(ta.value);
               applySpreadsheetToForm('C', ssFields);
               setTimeout(function() { ta.scrollTop = 0; }, 0);
               recalcVentilation();
            }
         }
      }
      // Update yellow cautions based on final checkbox state
      checkForDefaults('txtSpreadsheetData_C');
   } else if (controlName === 'chkSpreadsheetControl_S') {
      var chk = document.getElementById('chkSpreadsheetControl_S');
      var ta = document.getElementById('txtSpreadsheetData_S');
      if (chk.checked) {
         if (ta.value === '') {
            chk.checked = false;
         } else {
            var errMsg = validateSpreadsheetData(ta.value, 'Standard');
            var wa = document.getElementById('bpfWarningArea');
            if (errMsg) {
               if (wa) { wa.innerHTML = errMsg; wa.style.display = ''; }
               chk.checked = false;
            } else {
               if (wa) { wa.innerHTML = ''; wa.style.display = 'none'; }
               var ssFields = parseSpreadsheetFields(ta.value);
               applySpreadsheetToForm('S', ssFields);
               setTimeout(function() { ta.scrollTop = 0; }, 0);
               recalcVentilation();
            }
         }
      }
      // Update yellow cautions based on final checkbox state
      checkForDefaults('txtSpreadsheetData_S');
   }
}

async function fetchLoadLine() {
   var form = document.forms.HECACParameters;
   if (form['chkLockLoadLine'] && form['chkLockLoadLine'].checked) {
      try {
         // Save ventilation values so we can restore them after the engine run.
         // The iterative loop changes these, but we don't want fetchLoadLine to
         // have side-effects on the form â€” submitToEngine will do its own refinement.
         var savedVentNA = form['txtVentilationValue_NotAdvanced'].value;
         var savedVent = form['txtVentilationValue'].value;
         var savedSI = form['txtSI_Fraction_NotAdvanced'].value;
         var savedSIAdv = form['txtSI_Fraction'] ? form['txtSI_Fraction'].value : '';

         // Temporarily uncheck the lock so the engine computes the normal (unlocked) load line
         form['chkLockLoadLine'].checked = false;
         document.getElementById('txtSlope').value = '';
         document.getElementById('txtIntercept').value = '';

         // Sync hidden fields before running engine
         document.getElementById('txtDiscountRate_hidden').value = document.getElementById('txtDiscountRate').value;
         var capSel = document.getElementById('cmbTotalCap');
         document.getElementById('txtTotalCap').value = parseInt(capSel.options[capSel.selectedIndex].text, 10);
         syncBuildingModelFields();

         // Run the iterative engine loop (matching submitToEngine) so ventilation is refined
         var psychro = await import('./engine/psychro.js');
         var perf = await import('./engine/performance_module.js');
         var engine = await import('./engine/engine_module.js?v=14');
         for (var iter = 0; iter < 4; iter++) {
            await engine.exportBinCalcsJson(form, {});
            var ll = engine.getLastLoadLine();
            if (!ll || !ll.debug || !ll.debug.capacity || !ll.debug.design) break;
            var sensCapDesign = ll.debug.capacity.sensibleCapacityDesign;
            if (!Number.isFinite(sensCapDesign) || sensCapDesign <= 0) break;
            var prevVent = parseFloat(form['txtVentilationValue_NotAdvanced'].value);
            await refineVentilationFromCapacity(psychro, perf, sensCapDesign, ll.debug.design);
            var newVent = parseFloat(form['txtVentilationValue_NotAdvanced'].value);
            if (Math.abs(newVent - prevVent) <= 0.1) break;
         }
         // Final pass with converged ventilation
         await engine.exportBinCalcsJson(form, {});
         var ll = engine.getLastLoadLine();

         // Re-check the lock checkbox
         form['chkLockLoadLine'].checked = true;

         // Store the computed load line
         if (ll && ll.ok && Number.isFinite(ll.slope) && Number.isFinite(ll.intercept)) {
            document.getElementById('txtSlope').value = ll.slope;
            document.getElementById('txtIntercept').value = ll.intercept;
         }

         // Restore ventilation values so submitToEngine starts from the original state
         form['txtVentilationValue_NotAdvanced'].value = savedVentNA;
         form['txtVentilationValue'].value = savedVent;
         form['txtSI_Fraction_NotAdvanced'].value = savedSI;
         if (form['txtSI_Fraction']) form['txtSI_Fraction'].value = savedSIAdv;

         // Disable building-type controls while locked
         document.getElementById('cmbBuildingType').disabled = true;
         if (document.getElementById('txtBM_Slope')) document.getElementById('txtBM_Slope').disabled = true;
         if (document.getElementById('txtBM_Intercept')) document.getElementById('txtBM_Intercept').disabled = true;

         // Update default column
         checkForDefaults('chkLockLoadLine');
      } catch (e) {
         // Ensure checkbox is restored even on error
         form['chkLockLoadLine'].checked = true;
         console.log('fetchLoadLine error:', e);
      }
   } else {
      // Clear the locked slope and intercept
      document.getElementById('txtSlope').value = '';
      document.getElementById('txtIntercept').value = '';

      // Re-enable building-type controls
      document.getElementById('cmbBuildingType').disabled = false;
      if (document.getElementById('txtBM_Slope')) document.getElementById('txtBM_Slope').disabled = false;
      if (document.getElementById('txtBM_Intercept')) document.getElementById('txtBM_Intercept').disabled = false;

      // Update default column
      checkForDefaults('chkLockLoadLine');
   }
}

function toggleAdvanced() {
   checkForDefaults('chkAdvancedControls');
   var show = document.getElementById('chkAdvancedControls').checked;
   var rows = document.querySelectorAll('.advancedRow');
   for (var i = 0; i < rows.length; i++) {
      rows[i].style.display = show ? '' : 'none';
   }
   // Show/hide Building Model fields (Apply button + S/Int/VSF + Power button)
   var bmSpans = document.querySelectorAll('.bmAdvanced');
   for (var j = 0; j < bmSpans.length; j++) {
      bmSpans[j].style.display = show ? '' : 'none';
   }
   // Hide the non-advanced help icon when Advanced is ON
   var bmHide = document.querySelectorAll('.bmAdvancedHide');
   for (var j = 0; j < bmHide.length; j++) {
      bmHide[j].style.display = show ? 'none' : '';
   }
   // Populate power fields and BM visible fields when Advanced is turned ON
   if (show) {
      setFanPowerValuesAndDefaults(document.forms.HECACParameters['txtTotalCap'].value, 'AlsoFormValues');
      populateBMFieldsFromBuildingType();
   } else {
      // Clear the visible and hidden BM fields when Advanced is turned off
      document.getElementById('txtBM_Slope').value = '';
      document.getElementById('txtBM_Intercept').value = '';
      document.getElementById('txtBM_VentSlopeFraction').value = '';
      document.getElementById('txtBM_Slope_hidden').value = '';
      document.getElementById('txtBM_Intercept_hidden').value = '';
      document.getElementById('txtBM_VentSlopeFraction_hidden').value = '';
   }
   // Adjust instructions cell rowspan to accommodate advanced rows
   var instrCell = document.getElementById('tdInstructions');
   if (instrCell) instrCell.rowSpan = show ? 40 : 20;
   // ASP: SubmitTheFormToItself â†’ Establish_SIV recalculates S&I and ventilation
   recalcVentilation();
}

function populateBMFieldsFromBuildingType() {
   var buildingType = theSelectedValueInPullDownBox('cmbBuildingType') || 'Office-Medium';
   var model = buildingLoadModels[buildingType] || buildingLoadModels['Office-Medium'];
   if (!model) return;
   // Only populate if the fields are empty or building type changed
   var slopeEl = document.getElementById('txtBM_Slope');
   var intEl = document.getElementById('txtBM_Intercept');
   var vsfEl = document.getElementById('txtBM_VentSlopeFraction');
   slopeEl.value = Math.round(model.slope * 100) / 100;
   intEl.value = Math.round(model.intercept * 100) / 100;
   vsfEl.value = Math.round(model.ventFrac * 100) / 100;
   // Sync to hidden fields
   document.getElementById('txtBM_Slope_hidden').value = slopeEl.value;
   document.getElementById('txtBM_Intercept_hidden').value = intEl.value;
   document.getElementById('txtBM_VentSlopeFraction_hidden').value = vsfEl.value;
}

function SubmitTheBuildingModelParameters() {
   var slopeVal = parseFloat(document.getElementById('txtBM_Slope').value);
   var intVal = parseFloat(document.getElementById('txtBM_Intercept').value);
   if (slopeVal === 0 && intVal === 0) {
      window.alert('Both the intercept and the slope are zero. Try different values.');
   } else {
      // Sync visible fields to hidden fields
      document.getElementById('txtBM_Slope_hidden').value = document.getElementById('txtBM_Slope').value;
      document.getElementById('txtBM_Intercept_hidden').value = document.getElementById('txtBM_Intercept').value;
      document.getElementById('txtBM_VentSlopeFraction_hidden').value = document.getElementById('txtBM_VentSlopeFraction').value;
      // Clear the pending state and warning color
      document.getElementById('txtBM_postingstate').value = '';
      cFDW(false, 'tdBuildingType_controlscell');
      // Recalculate ventilation using the custom model parameters
      syncBuildingModelFields();
   }
}

function Reload_engine() {
   window.location.reload();
}

function restoreDefaults() {
   // Preserve Advanced Features state across reset
   var advWasChecked = document.getElementById('chkAdvancedControls').checked;
   document.forms.HECACParameters.reset();
   // Restore Advanced Features checkbox to its pre-reset state
   document.getElementById('chkAdvancedControls').checked = advWasChecked;
   // Reset hidden fields
   document.getElementById('txtTotalCap').value = '84';
   document.getElementById('txtBuildingType_hidden').value = 'Office-Medium';
   document.getElementById('txtDiscountRate_hidden').value = '7';
   // Clear locked load line values and re-enable building type
   document.getElementById('txtSlope').value = '';
   document.getElementById('txtIntercept').value = '';
   document.getElementById('cmbBuildingType').disabled = false;
   if (document.getElementById('txtBM_Slope')) document.getElementById('txtBM_Slope').disabled = false;
   if (document.getElementById('txtBM_Intercept')) document.getElementById('txtBM_Intercept').disabled = false;
   // Reset BM visible and hidden fields
   document.getElementById('txtBM_Slope_hidden').value = '';
   document.getElementById('txtBM_Intercept_hidden').value = '';
   document.getElementById('txtBM_VentSlopeFraction_hidden').value = '';
   if (document.getElementById('txtBM_Slope')) document.getElementById('txtBM_Slope').value = '';
   if (document.getElementById('txtBM_Intercept')) document.getElementById('txtBM_Intercept').value = '';
   if (document.getElementById('txtBM_VentSlopeFraction')) document.getElementById('txtBM_VentSlopeFraction').value = '';
   // Reset Total Capacity dropdown to 084
   var capSel = document.getElementById('cmbTotalCap');
   for (var i = 0; i < capSel.options.length; i++) {
      capSel.options[i].selected = (capSel.options[i].value === '084');
   }
   // Reset Degradation Factor dropdowns to 25 (form.reset doesn't work for JS-created options)
   document.getElementById('cmbPLDegrFactor_C').value = '25';
   document.getElementById('cmbPLDegrFactor_S').value = '25';
   document.getElementById('cmbPLDegrFactor_C').disabled = false;
   document.getElementById('cmbPLDegrFactor_S').disabled = false;
   document.forms.HECACParameters['txtPLDegrFactor_C'].value = '25';
   document.forms.HECACParameters['txtPLDegrFactor_S'].value = '25';
   // Reset condenser fan percent (disabled by default; only enabled for V-Spd fan mode)
   document.getElementById('txtCondFanPercent_C').disabled = true;
   document.getElementById('txtCondFanPercent_S').disabled = true;
   document.forms.HECACParameters['txtCondFanPercent_C_hidden'].value = '9.0';
   document.forms.HECACParameters['txtCondFanPercent_S_hidden'].value = '9.0';
   // Reset hidden Nstages fields
   document.forms.HECACParameters['txtNstages_C'].value = '1';
   document.forms.HECACParameters['txtNstages_S'].value = '1';
   // Reset State to MO and repopulate cities, then select Kansas City
   var stSel = document.getElementById('cmbState');
   for (var i = 0; i < stSel.options.length; i++) {
      stSel.options[i].selected = (stSel.options[i].text === 'MO');
   }
   updateCities();
   var ciSel = document.getElementById('cmbCityName2');
   for (var i = 0; i < ciSel.options.length; i++) {
      if (ciSel.options[i].text === 'Kansas City') { ciSel.options[i].selected = true; break; }
   }
   // Show/hide advanced rows based on preserved checkbox state
   var advRows = document.querySelectorAll('.advancedRow');
   for (var i = 0; i < advRows.length; i++) advRows[i].style.display = advWasChecked ? '' : 'none';
   var bmSpans = document.querySelectorAll('.bmAdvanced');
   for (var i = 0; i < bmSpans.length; i++) bmSpans[i].style.display = advWasChecked ? '' : 'none';
   var bmHide = document.querySelectorAll('.bmAdvancedHide');
   for (var i = 0; i < bmHide.length; i++) bmHide[i].style.display = advWasChecked ? 'none' : '';
   document.getElementById('txtBM_postingstate').value = '';
   // Clear any BPF / spreadsheet warning
   var wa = document.getElementById('bpfWarningArea');
   if (wa) { wa.innerHTML = ''; wa.style.display = 'none'; }
   // Reset ventilation value and default display back to 24 (SS data may have changed them)
   document.forms.HECACParameters['txtVentilationValue_NotAdvanced'].value = '24';
   document.forms.HECACParameters['txtVentilationValue'].value = '24';
   document.getElementById('tdVentilationValue_default').textContent = '24';
   cFDW(false, 'tdBuildingType_controlscell');
   if (advWasChecked) {
      populateBMFieldsFromBuildingType();
      setFanPowerValuesAndDefaults(document.forms.HECACParameters['txtTotalCap'].value, 'AlsoFormValues');
   }
   var instrCell = document.getElementById('tdInstructions');
   if (instrCell) instrCell.rowSpan = advWasChecked ? 40 : 20;
   // Reset colors
   var form = document.forms.HECACParameters;
   for (var i = 0; i < form.length; i++) {
      if (form.elements[i].name) checkForDefaults(form.elements[i].name);
   }
   // Show controls, hide results
   returnToControls();
}

function returnToControls() {
   document.getElementById('controlsView').style.display = '';
   document.getElementById('resultsView').style.display = 'none';
   document.getElementById('binChartsToggleLink').style.display = 'none';
}

// ---- FormatNumber: match ASP FormatNumber(value, decimals) ----
function formatNumber(val, decimals) {
   var n = Number(val);
   if (!isFinite(n)) return 'NA';
   var s = Math.abs(n).toFixed(decimals);
   // Add commas
   var parts = s.split('.');
   parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
   s = parts.join('.');
   return (n < 0 ? '-' : '') + s;
}

// ---- UPV: Uniform Present Value factor ----
function UPV(life, dr) {
   if (!dr || dr === 0) return life;
   var a;
   try { a = Math.pow(1 + dr, life); } catch(e) { a = 1; }
   if (!isFinite(a) || a === 0) a = 1;
   return (a - 1) / (dr * a);
}

// ---- NPV function (for ROR iteration) ----
function NPV_calc(annualCost_C, annualCost_S, unitCost_C, unitCost_S, dr, life) {
   var upv = UPV(life, dr);
   var lccC = (unitCost_C * 1000) + (annualCost_C * upv);
   var lccS = (unitCost_S * 1000) + (annualCost_S * upv);
   return lccS - lccC;
}

// ---- ROR Newton iteration ----
function ROR_calc(annualCost_C, annualCost_S, unitCost_C, unitCost_S, eqLife, initialGuess) {
   function npvErr(dr) {
      return 0 - NPV_calc(annualCost_C, annualCost_S, unitCost_C, unitCost_S, dr, eqLife);
   }
   var prev = initialGuess;
   var est, j;
   for (j = 0; j < 10; j++) {
      var errPrev = npvErr(prev);
      var slope = (errPrev - npvErr(prev + 0.0005)) / 0.0005;
      if (slope === 0 || !isFinite(slope)) return null;
      est = prev + errPrev / slope;
      if (Math.abs(npvErr(est)) < 0.01) break;
      prev = est;
   }
   if (est != null && isFinite(est) && j < 10 && est > 0) return 100 * est;
   return null;
}

// ---- Payback Newton iteration ----
function Payback_calc(annualCost_C, annualCost_S, unitCost_C, unitCost_S, dr, initialGuess) {
   function npvErr(life) {
      return 0 - NPV_calc(annualCost_C, annualCost_S, unitCost_C, unitCost_S, dr, life);
   }
   var prev = initialGuess;
   var est, j;
   for (j = 0; j < 10; j++) {
      var errPrev = npvErr(prev);
      var slope = (errPrev - npvErr(prev + 0.0005)) / 0.0005;
      if (slope === 0 || !isFinite(slope)) return null;
      est = prev + errPrev / slope;
      if (Math.abs(npvErr(est)) < 0.01) break;
      prev = est;
   }
   if (est != null && isFinite(est) && j <= 10 && est > 0) return est;
   return null;
}

// ---- Format_NA: match ASP Format_NA ----
function Format_NA(val, decimals) {
   if (val === null || val === undefined || val === 'NA') return 'NA';
   var n = Number(val);
   if (!isFinite(n)) return 'NA';
   return formatNumber(n, decimals);
}

// ---- Build payback chart data table ----
function buildPaybackDataTable(unitCost_C, unitCost_S, annualCost_C, annualCost_S, dr, eqLife, paybackYears) {
   var data = [];
   for (var yr = 0; yr <= eqLife; yr++) {
      var upv = UPV(yr, dr);
      var lccC = (unitCost_C * 1000) + (annualCost_C * upv);
      var lccS = (unitCost_S * 1000) + (annualCost_S * upv);
      data.push([yr, null, Math.round(lccC * 100) / 100, Math.round(lccS * 100) / 100]);
   }
   // Add payback annotation point
   if (paybackYears !== null && isFinite(paybackYears) && paybackYears > 0) {
      data.push([Math.round(paybackYears * 100000000) / 100000000, 'Payback', null, null]);
   }
   return data;
}

// ---- Draw payback chart ----
function drawPaybackChart(dataTable, eqLife, vAxisTitle) {
   var headerRow = [
      ['Age',
       {label: null, id:'aline', type:'string', role:'annotation'},
       'Candidate',
       'Standard']
   ];

   var googleDataTable = google.visualization.arrayToDataTable(headerRow.concat(dataTable));

   var ticks = [];
   for (var t = 0; t <= eqLife; t += 5) ticks.push(t);
   if (ticks[ticks.length - 1] !== eqLife) ticks.push(eqLife);

   var options = {
      hAxis: {title: 'System Life (years)', ticks: ticks, titleTextStyle: {italic: false}},
      vAxis: {title: vAxisTitle, textStyle: {fontSize: 12}, titleTextStyle: {italic: false}},
      annotations: {style: 'line', stem: {color: 'red'}, textStyle: {fontSize: 10}},
      curveType: 'function',
      series: {
         0: {pointShape: 'circle', lineWidth: 2, pointSize: 0, color: 'green'},
         1: {pointShape: 'circle', lineWidth: 2, pointSize: 0, color: 'tan'}
      },
      legend: {position: 'top', textStyle: {fontSize: 12}},
      chartArea: {left: 65, top: 40, width: '84%', height: '75%'}
   };

   var chart = new google.visualization.LineChart(document.getElementById('payback_div'));
   chart.draw(googleDataTable, options);
}


// ---- Quadratic (or linear) regression for predicting total load at setpoint (matches ASP ModelandPredict_TotalLoad) ----
function _predictTotalLoadAtSetpoint(bdata, odbSetpoint) {
   // Filter to bins with positive hours and valid totalSens
   var pts = [];
   for (var i = 0; i < bdata.length; i++) {
      var b = bdata[i];
      if (b.hours > 0 && Number.isFinite(b.totalSens)) pts.push({x: b.odb, y: b.totalSens});
   }
   var n = pts.length;
   if (n < 3) return null; // ASP requires at least 3 points for linear, 4+ for quadratic

   // Quadratic least-squares fit: y = a + b*x + c*x^2
   var sx=0, sx2=0, sx3=0, sx4=0, sy=0, sxy=0, sx2y=0;
   for (var j = 0; j < n; j++) {
      var x = pts[j].x, y = pts[j].y;
      sx += x; sx2 += x*x; sx3 += x*x*x; sx4 += x*x*x*x;
      sy += y; sxy += x*y; sx2y += x*x*y;
   }
   if (n > 3) {
      // Solve 3x3 system for quadratic
      var M = [[n, sx, sx2], [sx, sx2, sx3], [sx2, sx3, sx4]];
      var R = [sy, sxy, sx2y];
      // Gaussian elimination
      for (var col = 0; col < 3; col++) {
         var maxRow = col;
         for (var row = col+1; row < 3; row++) if (Math.abs(M[row][col]) > Math.abs(M[maxRow][col])) maxRow = row;
         var tmp = M[col]; M[col] = M[maxRow]; M[maxRow] = tmp;
         var tr = R[col]; R[col] = R[maxRow]; R[maxRow] = tr;
         if (Math.abs(M[col][col]) < 1e-12) return null;
         for (var row2 = col+1; row2 < 3; row2++) {
            var f = M[row2][col] / M[col][col];
            for (var k = col; k < 3; k++) M[row2][k] -= f * M[col][k];
            R[row2] -= f * R[col];
         }
      }
      var coef = [0,0,0];
      for (var ii = 2; ii >= 0; ii--) {
         var s = R[ii];
         for (var jj = ii+1; jj < 3; jj++) s -= M[ii][jj] * coef[jj];
         coef[ii] = s / M[ii][ii];
      }
      return coef[0] + coef[1]*odbSetpoint + coef[2]*odbSetpoint*odbSetpoint;
   } else {
      // Linear fit: y = a + b*x
      var denom = n*sx2 - sx*sx;
      if (Math.abs(denom) < 1e-12) return null;
      var a = (sy*sx2 - sx*sxy) / denom;
      var bCoef = (n*sxy - sx*sy) / denom;
      return a + bCoef*odbSetpoint;
   }
}

// ---- Draw bin Loads & Hours chart (matches ASP ChartBinCalcs_LoadsAndHours) ----
function drawBinLoadsChart(divId, bdata, title, chartExtra) {
   var el = document.getElementById(divId);
   if (!el || !bdata || bdata.length === 0) return;
   chartExtra = chartExtra || {};

   // Compute setpoint and design dot values
   var setpointOdb = chartExtra.setpointIdb;
   var setpointLoad = null;
   if (Number.isFinite(setpointOdb)) {
      setpointLoad = _predictTotalLoadAtSetpoint(bdata, setpointOdb);
   }

   var designOdb = chartExtra.designOdb;
   var designLoad = chartExtra.designLoad; // sensibleCapacityDesign (= sensLoadAtDesign * LRF)
   var showDesignDot = Number.isFinite(designOdb) && Number.isFinite(designLoad) && !chartExtra.lockLoadLine;

   // Build data: ODB, TotalSens, Remaining, NonVent, TotalHours, ScheduledHours, SetPoint, DesignCond
   var header = [['ODB', 'Total Sens Load', 'Remaining Load', 'Non-Vent Load', 'Total Hours', 'Scheduled Hours', 'Set Point', 'Design Cond.']];
   var rows = [];
   for (var i = 0; i < bdata.length; i++) {
      var b = bdata[i];
      rows.push([b.odb, b.totalSens, b.remaining, b.nonVent, b.totalHours || 0, b.hours, null, null]);
   }
   // Setpoint dot row
   if (setpointLoad !== null) {
      rows.push([setpointOdb, null, null, null, null, null, Math.round(setpointLoad * 10) / 10, null]);
   }
   // Design condition dot row
   if (showDesignDot) {
      rows.push([Math.round(designOdb * 10) / 10, null, null, null, null, null, null, Math.round(designLoad * 10) / 10]);
   }

   var dt = google.visualization.arrayToDataTable(header.concat(rows));

   var options = {
      title: title,
      titleTextStyle: {fontSize: 13, bold: true},
      hAxis: {title: 'Outside Dry Bulb (Â°F)', titleTextStyle: {italic: false}},
      vAxes: {
         0: {title: 'Load (kBtuh)', titleTextStyle: {italic: false}},
         1: {title: 'Hours', titleTextStyle: {italic: false}}
      },
      series: {
         0: {targetAxisIndex: 0, type: 'line', lineWidth: 2, color: '#c00'},
         1: {targetAxisIndex: 0, type: 'line', lineWidth: 2, color: '#090'},
         2: {targetAxisIndex: 0, type: 'line', lineWidth: 2, color: '#888', lineDashStyle: [6, 3]},
         3: {targetAxisIndex: 1, type: 'bars', color: '#7ab'},
         4: {targetAxisIndex: 1, type: 'bars', color: '#9cf'},
         5: {targetAxisIndex: 0, type: 'line', lineWidth: 0, pointSize: 10, pointShape: 'circle', color: '#36f', visibleInLegend: (setpointLoad !== null)},
         6: {targetAxisIndex: 0, type: 'line', lineWidth: 0, pointSize: 10, pointShape: 'circle', color: '#e00', visibleInLegend: showDesignDot}
      },
      legend: {position: 'top', textStyle: {fontSize: 10}},
      chartArea: {left: 60, top: 50, width: '72%', height: '68%'},
      bar: {groupWidth: 12},
      seriesType: 'line'
   };

   var chart = new google.visualization.ComboChart(el);
   chart.draw(dt, options);
}

// ---- Draw bin Performance chart (matches ASP ChartBinCalcs_Performance) ----
function drawBinPerfChart(divId, bdata, title, chartExtra) {
   var el = document.getElementById(divId);
   if (!el || !bdata || bdata.length === 0) return;
   chartExtra = chartExtra || {};

   var header = [['ODB', 'E_Cond', 'E_Fan', 'TCF (%)', 'PCF (%)', 'S/T (%)']];
   var rows = [];
   for (var i = 0; i < bdata.length; i++) {
      var b = bdata[i];
      rows.push([b.odb, b.eCond, b.eFan, b.tcf * 100, b.pcf * 100, b.stRatio * 100]);
   }
   var dt = google.visualization.arrayToDataTable(header.concat(rows));

   // Use shared axis maxima if provided (matching ASP's shared dblYMax_energy / dblYMax_factor)
   var vAxis0 = {title: 'Energy (kWh)', titleTextStyle: {italic: false}, minValue: 0};
   var vAxis1 = {title: 'Correction Factor (%)', titleTextStyle: {italic: false}, minValue: 0};
   if (chartExtra.yMaxEnergy > 0) vAxis0.maxValue = chartExtra.yMaxEnergy;
   if (chartExtra.yMaxFactor > 0) vAxis1.maxValue = chartExtra.yMaxFactor;

   var options = {
      title: title,
      titleTextStyle: {fontSize: 13, bold: true},
      hAxis: {title: 'Outside Dry Bulb (Â°F)', titleTextStyle: {italic: false}},
      vAxes: { 0: vAxis0, 1: vAxis1 },
      series: {
         0: {targetAxisIndex: 0, type: 'bars', color: '#69c'},
         1: {targetAxisIndex: 0, type: 'bars', color: '#fc6'},
         2: {targetAxisIndex: 1, type: 'line', lineWidth: 2, color: '#c00'},
         3: {targetAxisIndex: 1, type: 'line', lineWidth: 2, color: '#090'},
         4: {targetAxisIndex: 1, type: 'line', lineWidth: 2, color: '#00c'}
      },
      legend: {position: 'top', textStyle: {fontSize: 10}},
      chartArea: {left: 60, top: 50, width: '72%', height: '68%'},
      bar: {groupWidth: 12},
      seriesType: 'bars',
      isStacked: false
   };

   var chart = new google.visualization.ComboChart(el);
   chart.draw(dt, options);
}

// ---- Draw all bin charts stored in window._binChartsQueue ----
function drawAllBinCharts() {
   var q = window._binChartsQueue;
   if (!q || !q.length) return;
   for (var i = 0; i < q.length; i++) {
      var c = q[i];
      if (c.type === 'loads') drawBinLoadsChart(c.divId, c.data, c.title, c.chartExtra);
      else if (c.type === 'perf') drawBinPerfChart(c.divId, c.data, c.title, c.chartExtra);
   }
}

// ---- Draw bin charts into the 4x2 grid divs ----
function drawBinChartsGrid() {
   var q = window._binChartsGridQueue;
   if (!q || !q.length) return;
   for (var i = 0; i < q.length; i++) {
      var c = q[i];
      if (c.type === 'loads') drawBinLoadsChart(c.divId, c.data, c.title, c.chartExtra);
      else if (c.type === 'perf') drawBinPerfChart(c.divId, c.data, c.title, c.chartExtra);
   }
}

// ---- Toggle between bin tables view and charts-only 4x2 grid ----
function toggleBinChartsGrid() {
   var tables = document.getElementById('binTablesView');
   var grid = document.getElementById('binChartsGrid');
   var link = document.getElementById('binChartsToggleAnchor');
   if (!tables || !grid) return;
   if (grid.style.display === 'none') {
      tables.style.display = 'none';
      grid.style.display = '';
      drawBinChartsGrid();
      if (link) link.textContent = 'Show bin tables and charts';
   } else {
      grid.style.display = 'none';
      tables.style.display = '';
      if (link) link.textContent = 'Show side-by-side charts';
   }
}

// ============================================================
// Building Load Model lookup (matches ASP Building_Load_Model Sub)
// ============================================================
var cnsBtuPerMJ = 947.817;
function KBtuHperDegF(mjhPerDegC) { return mjhPerDegC * (cnsBtuPerMJ / 1000) * 5 / 9; }
function KBtuH(mjh) { return mjh * cnsBtuPerMJ / 1000; }

var buildingLoadModels = {
   'Apartment-MidRise':      { slope: KBtuHperDegF(12.32),  intercept: KBtuH(128.3), ventFrac: 0.30 },
   'Healthcare-Hospital':    { slope: KBtuHperDegF(55.8),   intercept: KBtuH(1370),  ventFrac: 0.80 },
   'Healthcare-Outpatient':  { slope: KBtuHperDegF(41.8),   intercept: KBtuH(704),   ventFrac: 0.25 },
   'Hotel-Large':            { slope: KBtuHperDegF(149.5),  intercept: KBtuH(1313),  ventFrac: 0.60 },
   'Office-Small':           { slope: KBtuHperDegF(3.46),   intercept: KBtuH(43.4),  ventFrac: 0.32 },
   'Office-Medium':          { slope: KBtuHperDegF(33.7),   intercept: KBtuH(382),   ventFrac: 0.50 },
   'Office-Large':           { slope: KBtuHperDegF(410),    intercept: KBtuH(4250),  ventFrac: 0.69 },
   'Restaurant-FastFood':    { slope: KBtuHperDegF(9.97),   intercept: KBtuH(45.8),  ventFrac: 0.76 },
   'Restaurant-SitDown':     { slope: KBtuHperDegF(17.65),  intercept: KBtuH(98.0),  ventFrac: 0.79 },
   'Retail-StandAlone':      { slope: KBtuHperDegF(26.5),   intercept: KBtuH(245.4), ventFrac: 0.63 },
   'Retail-StripMall':       { slope: KBtuHperDegF(35.4),   intercept: KBtuH(187.2), ventFrac: 0.40 },
   'School-Primary':         { slope: KBtuHperDegF(63.5),   intercept: KBtuH(718),   ventFrac: 0.61 },
   'School-Secondary':       { slope: KBtuHperDegF(95),     intercept: KBtuH(779),   ventFrac: 0.40 },
   'Warehouse':              { slope: KBtuHperDegF(7.1),    intercept: KBtuH(5.7),   ventFrac: 0.21 }
};

// SIV: Calculate S&I fraction from building model (matches ASP SIV Sub)
// Returns { siFraction, ventSlopeFraction, slope, intercept }
// If customModel is provided ({slope, intercept, ventFrac}), use it instead of the default.
function calcSIV(buildingType, odbDesignF, idbSetpointF, customModel) {
   var model;
   if (customModel && isFinite(customModel.slope) && isFinite(customModel.intercept)) {
      model = customModel;
   } else {
      model = buildingLoadModels[buildingType];
      if (!model) model = buildingLoadModels['Office-Medium'];
   }
   var dtF = odbDesignF - idbSetpointF;
   var siFraction = model.intercept / (model.intercept + model.slope * dtF);
   return {
      siFraction: Math.round(siFraction * 1000) / 1000,
      ventSlopeFraction: model.ventFrac,
      slope: model.slope,
      intercept: model.intercept
   };
}

// Sync building-model hidden fields based on current control values.
// When Advanced is ON and custom BM values exist, use them for S&I computation.
// Otherwise, leave BM hidden fields empty so the engine computes the load line itself.
function syncBuildingModelFields() {
   var form = document.forms.HECACParameters;
   var buildingType = theSelectedValueInPullDownBox('cmbBuildingType') || 'Office-Medium';
   var customModel = getCustomBMModel();

   // When Advanced is ON with custom BM values, the custom model affects the
   // S&I fraction and ventilation (computed above/below). We must NOT put the
   // building-model slope/intercept into BM_hidden because the engine would
   // misinterpret them as final load-line values. The engine should compute
   // the load line from S&I + sensible capacity (the normal path).
   // Always clear BM_hidden so the engine computes the load line itself.
   document.getElementById('txtBM_Slope_hidden').value = '';
   document.getElementById('txtBM_Intercept_hidden').value = '';
   document.getElementById('txtBM_VentSlopeFraction_hidden').value = '';

   // S&I fraction is NOT recomputed here because this function doesn't have
   // access to the actual design ODB (which varies by city).  Using a hardcoded
   // ODB (e.g. 95) produces the wrong S&I (0.505 instead of 0.527 for KC).
   // S&I is correctly computed by recalcVentilation() / refineVentilationFromCapacity()
   // which have the real design conditions from the engine.  Those run on page
   // load and on control changes that trigger ASP post-backs (building type,
   // city, state, schedule, capacity, etc.).

   // Keep visible advanced ventilation field in sync with the not-advanced field
   form['txtVentilationValue'].value = form['txtVentilationValue_NotAdvanced'].value;

   // Update building type hidden
   document.getElementById('txtBuildingType_hidden').value = buildingType;
}

// Returns a custom model object if Advanced is ON and the user has entered custom values,
// or null if defaults should be used.
function getCustomBMModel() {
   var advOn = document.getElementById('chkAdvancedControls').checked;
   if (!advOn) return null;
   var slopeEl = document.getElementById('txtBM_Slope');
   var intEl = document.getElementById('txtBM_Intercept');
   var vsfEl = document.getElementById('txtBM_VentSlopeFraction');
   if (!slopeEl || !intEl || !vsfEl) return null;
   var s = parseFloat(slopeEl.value);
   var i = parseFloat(intEl.value);
   var v = parseFloat(vsfEl.value);
   if (!isFinite(s) || !isFinite(i)) return null;
   if (!isFinite(v)) v = 0.5;
   return { slope: s, intercept: i, ventFrac: v };
}

// Refine ventilation and load-line using the actual sensible capacity at design from the engine.
// This replaces the rough totalCap approximation with the engine's own computed value,
// matching what ASP does (ASP uses NetSenCap_Stage_Adjusted_KBtuh at design conditions).
// Also stores the computed load-line slope/intercept in BM hidden fields so the engine
// uses them directly on the second pass (bypassing the ventilation-dependent calculation).
async function refineVentilationFromCapacity(psychro, perf, sensCapDesign, designInfo) {
   var form = document.forms.HECACParameters;
   var buildingType = theSelectedValueInPullDownBox('cmbBuildingType') || 'Office-Medium';
   var customModel = getCustomBMModel();
   var model = customModel || buildingLoadModels[buildingType] || buildingLoadModels['Office-Medium'];
   var idb = parseInt(theSelectedValueInPullDownBox('cmbIDB'), 10) || 75;
   var totalCap = parseInt(document.getElementById('txtTotalCap').value, 10) || 84;
   var cfm = (totalCap / 12) * 400;
   var oversizePct = parseInt(theSelectedValueInPullDownBox('cmbOversizePercent'), 10) || 0;
   var oversizeLRF = (oversizePct / 100) + 1;

   var odbDesign = designInfo.odbDesign;
   var ohrDesign = designInfo.ohrDesign;
   var pressureDesign = designInfo.pressureDesign;
   if (!Number.isFinite(odbDesign) || !Number.isFinite(ohrDesign) || !Number.isFinite(pressureDesign)) return;

   // S&I fraction from building model using actual design ODB
   // ASP uses the UNROUNDED S&I fraction for the load line slope (line 564),
   // and only rounds to 3 decimals for the form control (line 554).
   var dtDesign = odbDesign - idb;
   if (dtDesign <= 0) return;
   var siFractionRaw = model.intercept / (model.intercept + model.slope * dtDesign);
   form['txtSI_Fraction_NotAdvanced'].value = Math.round(siFractionRaw * 1000) / 1000;
   if (form['txtSI_Fraction']) form['txtSI_Fraction'].value = Math.round(siFractionRaw * 1000) / 1000;

   // Use actual sensible capacity at design from engine
   var sensLoadDesign = sensCapDesign / oversizeLRF;

   // Load line slope (matches ASP Controls.asp line 564, uses unrounded dblSI_fraction)
   var loadLineSlope = (sensLoadDesign - sensLoadDesign * siFractionRaw) / dtDesign;

   // Ventilation slope = loadLineSlope * ventSlopeFraction
   var ventSlope = loadLineSlope * model.ventFrac;

   // CFM from ventilation slope
   var ventCFM = perf.CFM_VentSlope(ventSlope, ohrDesign, odbDesign, pressureDesign);
   var ventPercent = 100 * ventCFM / cfm;

   // Round to 1 decimal (matches ASP round(dblVentilation_percent, 1) at line 584)
   ventPercent = Math.round(ventPercent * 10) / 10;

   form['txtVentilationValue_NotAdvanced'].value = ventPercent;
   // Keep visible advanced field in sync with computed value
   form['txtVentilationValue'].value = ventPercent;
}

// ============================================================
// SUBMIT: Run JS engine and render results
// ============================================================
async function submitToEngine() {
   var form = document.forms.HECACParameters;

   // Check if building model parameters have been edited but not Applied
   var advOn = document.getElementById('chkAdvancedControls').checked;
   if (advOn && document.getElementById('txtBM_postingstate').value === 'Pend') {
      window.alert('The building model parameters have been edited. Please click the Apply button before submitting the form. This will update various parameters on the form that are dependent on the building model.');
      return;
   }

   // ---- Numeric field validation (matches ASP CheckThenSubmit) ----

   // Sync txtTotalCap from combo before validating
   var capSel_v = document.getElementById('cmbTotalCap');
   var dblCapacity = parseInt(capSel_v.options[capSel_v.selectedIndex].text, 10);
   form['txtTotalCap'].value = dblCapacity;

   if (checkTextBoxNumber('txtTotalCap', 'NoneOrOne')) {
      window.alert('The total capacity field must contain a number (e.g. 84 or 84.4)'); return;
   } else if (dblCapacity < 36 || dblCapacity > 360) {
      window.alert('The total capacity must be within the range of 36 to 360 kBtuh.'); return;
   }

   if (advOn && checkTextBoxNumber('txtSI_Fraction', 'NoneOrOne')) {
      window.alert('The S&I fraction must be a positive number (e.g. 0.622)'); return;
   } else if (advOn) {
      var dblSI = form['txtSI_Fraction'].value * 1.0;
      if (dblSI < 0 || dblSI > 1) {
         window.alert('The S&I fraction must be a number between 0 and 1.'); return;
      }
   }

   if (advOn && checkTextBoxNumber('txtVentilationValue', 'NoneOrOne')) {
      window.alert('The ventilation value must be a positive number (e.g. 10.5)'); return;
   }

   if (checkTextBoxNumber('txtElectricityRate', 'One')) {
      window.alert('The electric utility rate field must contain a decimal number (e.g. 0.05)'); return;
   }

   if (checkTextBoxNumber('txtEER', 'NoneOrOne')) {
      window.alert('The candidate unit EER field must contain a number (e.g. 11 or 11.3)'); return;
   } else if (form['txtEER'].value * 1.0 < 5 || form['txtEER'].value * 1.0 > 17) {
      window.alert('The EER of the candidate unit must be within the range of 5.0 to 17.0.'); return;
   }

   if (checkTextBoxNumber('txtEER_Standard', 'NoneOrOne')) {
      window.alert('The standard unit EER field must contain a number (e.g. 11 or 11.3)'); return;
   } else if (form['txtEER_Standard'].value * 1.0 < 5 || form['txtEER_Standard'].value * 1.0 > 17) {
      window.alert('The EER of the standard unit must be within the range of 5.0 to 17.0.'); return;
   }

   if (checkTextBoxNumber('txtUnitCost', 'NoneOrOne')) {
      window.alert('The Unit Cost field (Candidate) must contain a number (e.g. 5 or 5.4)'); return;
   } else if (form['txtUnitCost'].value * 1.0 < 0 || form['txtUnitCost'].value * 1.0 > 100) {
      window.alert('The cost of the candidate unit must be within the range of 0.0 to 100.0 (in units of k$).'); return;
   }

   if (checkTextBoxNumber('txtUnitCost_Standard', 'NoneOrOne')) {
      window.alert('The Unit Cost field (Standard) must contain a number (e.g. 2 or 2.4)'); return;
   } else if (form['txtUnitCost_Standard'].value * 1.0 < 0 || form['txtUnitCost_Standard'].value * 1.0 > 100) {
      window.alert('The cost of the standard unit must be within the range of 0.0 to 100.0 (in units of k$).'); return;
   }

   if (advOn && checkTextBoxNumber('txtCondFanPercent_C', 'NoneOrOne')) {
      window.alert('The Condenser Fan percentage value (Candidate) must be a number (e.g. 9 or 9.5)'); return;
   } else if (advOn && (form['txtCondFanPercent_C'].value * 1.0 < 0 || form['txtCondFanPercent_C'].value * 1.0 > 20)) {
      window.alert('The Condenser Fan percentage (Candidate) must be within the range of 0.0 to 20.0.'); return;
   }

   if (advOn && checkTextBoxNumber('txtCondFanPercent_S', 'NoneOrOne')) {
      window.alert('The Condenser Fan percentage value (Standard) must be a number (e.g. 9 or 9.5)'); return;
   } else if (advOn && (form['txtCondFanPercent_S'].value * 1.0 < 0 || form['txtCondFanPercent_S'].value * 1.0 > 20)) {
      window.alert('The Condenser Fan percentage (Standard) must be within the range of 0.0 to 20.0.'); return;
   }

   if (form['txtUnitCost_Standard'].value * 1.0 == form['txtUnitCost'].value * 1.0) {
      window.alert('The cost of the standard unit must differ from the cost of the candidate unit.'); return;
   }

   if (checkTextBoxNumber('txtMaintenance_Standard', 'NoneOrOne')) {
      window.alert('The annual-maintenance cost field for the standard unit must contain a number (e.g. 100 or 100.5)'); return;
   } else if (form['txtMaintenance_Standard'].value * 1.0 < 0 || form['txtMaintenance_Standard'].value * 1.0 > 5000) {
      window.alert('The annual maintenance cost of the standard unit must be within the range of 0 to 5000.'); return;
   }

   if (checkTextBoxNumber('txtMaintenance_Candidate', 'NoneOrOne')) {
      window.alert('The annual-maintenance cost field for the candidate unit must contain a number (e.g. 100 or 100.5)'); return;
   } else if (form['txtMaintenance_Candidate'].value * 1.0 < 0 || form['txtMaintenance_Candidate'].value * 1.0 > 5000) {
      window.alert('The annual maintenance cost of the candidate unit must be within the range of 0 to 5000.'); return;
   }

   if (checkTextBoxNumber('txtDiscountRate', 'NoneOrOne')) {
      window.alert('The discount rate field must contain a number (e.g. 3 or 3.4)'); return;
   } else if (form['txtDiscountRate'].value * 1.0 <= 0 || form['txtDiscountRate'].value * 1.0 > 20) {
      window.alert('The discount rate must be greater than 0.0 and less than 20.0'); return;
   }

   if (checkTextBoxNumber('txtNUnits', 'None')) {
      window.alert('The number of units field must contain an integer number (e.g. 3)'); return;
   } else if (form['txtNUnits'].value * 1 < 1) {
      window.alert('The number of units field must be set to 1 or greater'); return;
   }

   if (advOn && problemInPowerValue('txtBFn_kw_C'))  return;
   if (advOn && problemInPowerValue('txtBFn_kw_S'))  return;
   if (advOn && problemInPowerValue('txtAux_kw_C'))  return;
   if (advOn && problemInPowerValue('txtAux_kw_S'))  return;
   if (advOn && problemInPowerValue('txtCond_kw_C')) return;
   if (advOn && problemInPowerValue('txtCond_kw_S')) return;

   if (advOn && checkTextBoxNumber('txtDemandCostPerKW', 'NoneOrOne')) {
      window.alert('The demand cost field must contain a number.'); return;
   }

   // Check N-Spd fan mode requires stages >= 2 (matches ASP validation)
   if (advOn) {
      var nSpdMsg = function(unit) {
         return 'Selecting "N-Spd..." in the "E-Fan and Condenser" control for the ' + unit + ' unit requires the number of stages for the ' + unit + ' unit to be 2 or higher.';
      };
      var fanC = theSelectedValueInPullDownBox('cmbFanControls_C') || '';
      var fanS = theSelectedValueInPullDownBox('cmbFanControls_S') || '';
      if (fanC.indexOf('N-Spd') !== -1 && form['cmbNstages_C'].value === '1') {
         window.alert(nSpdMsg('candidate'));
         return;
      }
      if (fanS.indexOf('N-Spd') !== -1 && form['cmbNstages_S'].value === '1') {
         window.alert(nSpdMsg('standard'));
         return;
      }
   }

   if (advOn && form['cmbSpecific_RTU_C'] && form['cmbSpecific_RTU_C'].value !== 'None' && form['txtSpreadsheetData_C'] && form['txtSpreadsheetData_C'].value !== '') {
      window.alert('Spreadsheet data is not allowed if a specific candidate unit has been selected. Please set the "Specific Candidate Unit" field to "None" or clear out the "Spreadsheet Data" field.');
      return;
   }

   // Sync hidden fields
   document.getElementById('txtDiscountRate_hidden').value = document.getElementById('txtDiscountRate').value;

   // Sync building-model hidden fields
   syncBuildingModelFields();

   // Hide controls-page warning so it doesn't show alongside results-page warning
   var wa = document.getElementById('bpfWarningArea');
   if (wa) { wa.innerHTML = ''; wa.style.display = 'none'; }
   // Show a loading indicator
   document.getElementById('controlsView').style.display = 'none';
   document.getElementById('resultsView').style.display = '';
   document.getElementById('resultsContent').innerHTML =
      '<h2>RESULTS</h2><p>Running calculation engine...</p>';

   try {
      // Import modules
      var psychro = await import('./engine/psychro.js');
      var perf = await import('./engine/performance_module.js');

      // Dynamically import the JS engine module
      var engine = await import('./engine/engine_module.js?v=14');
      if (!engine || typeof engine.exportBinCalcsJson !== 'function') {
         throw new Error('exportBinCalcsJson() not available in engine_module.js');
      }

      // ASP flow: Controls.asp refines ventilation via Establish_SIV on page
      // load (and on post-backs that call SubmitTheFormToItself), then Submit
      // sends the form directly to Engine.asp with no further refinement.
      // The JS equivalent is recalcVentilation(), which runs on page load and
      // on control changes that trigger ASP post-backs.  At submit time we
      // just run the engine once with whatever ventilation is in the form.
      var jsJson = await engine.exportBinCalcsJson(form, {});

      // Build the results HTML
      var html = buildResultsHTML(jsJson);
      document.getElementById('resultsContent').innerHTML = html;

      // Show or hide the bin charts toggle link based on whether bin details are present
      var hasBinTables = !!document.getElementById('binTablesView');
      document.getElementById('binChartsToggleLink').style.display = hasBinTables ? '' : 'none';
      // Reset link text and grid state for fresh results
      var toggleAnchor = document.getElementById('binChartsToggleAnchor');
      if (toggleAnchor) toggleAnchor.textContent = 'Show side-by-side charts';

      // Draw the payback chart and bin charts after DOM is updated
      google.charts.setOnLoadCallback(function() {
         if (window._paybackChartData && document.getElementById('payback_div')) {
            drawPaybackChart(
               window._paybackChartData.dataTable,
               window._paybackChartData.eqLife,
               window._paybackChartData.vAxisTitle
            );
         }
         drawAllBinCharts();
      });
      // Also try drawing immediately if Google Charts is already loaded
      if (google.visualization && google.visualization.ComboChart) {
         if (window._paybackChartData && document.getElementById('payback_div')) {
            drawPaybackChart(
               window._paybackChartData.dataTable,
               window._paybackChartData.eqLife,
               window._paybackChartData.vAxisTitle
            );
         }
         drawAllBinCharts();
      }

   } catch(e) {
      var errStr = String(e && (e.message || e));
      // Check for BPF/ADP error â€” show ASP-style warning (psychro.asp lines 368-370)
      var bpfDetail = '';
      var bpfMatch = errStr.match(/Engine error:\s*(.*)/i);
      if (bpfMatch) bpfDetail = bpfMatch[1].trim();
      if (bpfDetail && /\bBPF\b|Supply:|ADP:/i.test(bpfDetail)) {
         var unitName = 'Candidate';
         if (/Standard Unit/i.test(bpfDetail)) unitName = 'Standard';
         var cleanMsg = bpfDetail.replace(/^(Candidate|Standard) Unit BPF:\s*/i, '');
         document.getElementById('resultsContent').innerHTML =
            '<h2>RESULTS</h2>' +
            '<p style="color:#cc0000; font-weight:bold;">' +
            'WARNING: Try lowering the S/T ratio for the ' + unitName + ' Unit.<br>' +
            'ERROR Message from BPF calculation: ' + cleanMsg +
            '</p>';
      } else {
         document.getElementById('resultsContent').innerHTML =
            '<h2>Engine Error</h2><pre>' + String(e && (e.stack || e.message || e)) + '</pre>';
      }
   }
}


// ============================================================
// Build the results HTML from the JS engine JSON output
// ============================================================
function buildResultsHTML(js) {
   var econ = js.economics || {};
   var annual = js.annual || {};
   var annualRaw = js.annual_raw || {};
   var nUnits = js.meta?.nUnits || 1;

   // Read form values for display
   var form = document.forms.HECACParameters;
   var cityName = theSelectedValueInPullDownBox('cmbCityName2');
   var stateName = theSelectedValueInPullDownBox('cmbState');
   var cityHeader = cityName.toUpperCase() + ', ' + stateName;

   var elecRate = econ.electricityRate || 0.08;
   var drRate = econ.discountRate || 0.05;
   var eqLife = econ.equipmentLife || 15;
   var chartPW = econ.chartPW;

   var unitCost_C = econ.candidate?.unitCost || 0;
   var unitCost_S = econ.standard?.unitCost || 0;
   var maint_C = econ.candidate?.maintenance || 0;
   var maint_S = econ.standard?.maintenance || 0;

   // Engine annual energy is already nUnits-scaled; display directly
   var totalC = annual.candidate?.total || 0;
   var totalS = annual.standard?.total || 0;
   var energySavings = totalS - totalC;

   // Per-unit energy for economics (matches ASP: objBD_C.Energy_Annual_Total is per-unit)
   var totalC_perUnit = totalC / nUnits;
   var totalS_perUnit = totalS / nUnits;

   // Per-unit demand cost
   var demandCost_C = econ.candidate?.demandCost || 0;
   var demandCost_S = econ.standard?.demandCost || 0;

   // Per-unit annual operating cost (matches ASP Engine.asp line 2256)
   var annualCost_C = (totalC_perUnit * elecRate) + maint_C + demandCost_C;
   var annualCost_S = (totalS_perUnit * elecRate) + maint_S + demandCost_S;
   var costAnnualSavings = annualCost_S - annualCost_C;
   var capCostSavings = 1000 * (unitCost_C - unitCost_S);

   // Per-unit LCC (matches ASP Engine.asp lines 2290-2296)
   var upv = UPV(eqLife, drRate);
   var lccC = (unitCost_C * 1000) + (annualCost_C * upv);
   var lccS = (unitCost_S * 1000) + (annualCost_S * upv);
   var lccSavings = lccS - lccC;

   // Per-unit annualized cost
   var lcf = (upv !== 0) ? (1 / upv) : 0;
   var annualizedC = lccC * lcf;
   var annualizedS = lccS * lcf;
   var annualizedSavings = annualizedS - annualizedC;

   // Per-unit NPV = LCC_Standard - LCC_Candidate
   var npv = lccSavings;

   // Payback (per-unit, same ratio regardless of nUnits)
   var simplePayback = (costAnnualSavings !== 0) ? (capCostSavings / costAnnualSavings) : -1;

   // Discounted payback via Newton iteration (matches ASP PayBack function)
   function _npvAtLife(dr, life) {
      var u = UPV(life, dr);
      return ((unitCost_S * 1000) + (annualCost_S * u)) - ((unitCost_C * 1000) + (annualCost_C * u));
   }
   var discountedPayback;
   if (simplePayback < 0) {
      discountedPayback = 0; // Immediate
   } else if (simplePayback > 100) {
      discountedPayback = -1;
   } else {
      var prev = simplePayback, est = null, j;
      for (j = 0; j < 10; j++) {
         var errPrev = 0 - _npvAtLife(drRate, prev);
         var errDelta = 0 - _npvAtLife(drRate, prev + 0.0005);
         var slope = (errPrev - errDelta) / 0.0005;
         if (slope === 0 || !Number.isFinite(slope)) break;
         est = prev + errPrev / slope;
         var errEst = 0 - _npvAtLife(drRate, est);
         if (Math.abs(errEst) < 0.01) break;
         prev = est;
      }
      discountedPayback = (est !== null && Number.isFinite(est) && j <= 10 && est > 0) ? est : -1;
   }

   // ROR: find discount rate where NPV = 0
   function _npvAtRate(dr) {
      var u = UPV(eqLife, dr);
      return ((unitCost_S * 1000) + (annualCost_S * u)) - ((unitCost_C * 1000) + (annualCost_C * u));
   }
   var ror = null;
   if (capCostSavings !== 0) {
      var prevR = costAnnualSavings / capCostSavings, estR = null, jR;
      for (jR = 0; jR < 10; jR++) {
         var eP = 0 - _npvAtRate(prevR);
         var eD = 0 - _npvAtRate(prevR + 0.0005);
         var sR = (eP - eD) / 0.0005;
         if (sR === 0 || !Number.isFinite(sR)) break;
         estR = prevR + eP / sR;
         var eE = 0 - _npvAtRate(estR);
         if (Math.abs(eE) < 0.01) break;
         prevR = estR;
      }
      if (estR !== null && Number.isFinite(estR) && jR < 10 && estR > 0) ror = 100 * estR;
   }

   // SIR
   var lccAnnualSavings = costAnnualSavings * upv;
   var sir = (capCostSavings !== 0) ? (lccAnnualSavings / capCostSavings) : 0;

   // Payback display value
   var paybackDisplay;
   if (chartPW) {
      if (discountedPayback === 0) paybackDisplay = 'Immediate';
      else if (discountedPayback < 0 || discountedPayback === -1) paybackDisplay = 'NA';
      else paybackDisplay = Format_NA(discountedPayback, 1);
   } else {
      paybackDisplay = Format_NA(simplePayback, 1);
   }

   // Percent savings helper
   function pctSav(savings, base) {
      if (base === 0) return 'NA';
      return formatNumber(100 * savings / base, 0) + '%';
   }

   var h = '';

   // ---- Show Bin Calculations (rendered first, matching ASP order) ----
   var showDetails = form['chkDetails'].checked;
   var dc = js.designConditions || {};
   var bins = js.bins || {};
   if (showDetails && dc.outdoor && bins.candidate) {
      var oc = dc.outdoor;
      var ic = dc.indoor;
      var ec = dc.entering;
      var cap = dc.capacity;
      var lds = dc.loads;
      var llEq = dc.loadLine;
      var bpfs = dc.bpf;

      h += '<br>';

      if (llEq.locked) {
         // Locked load line: show header + equation instead of full design table (ASP Engine.asp lines 733-735)
         h += "<h1>NON-VENTILATION LOAD LINE <span style='color:#f00 !important'> (LOCKED)</span></h1>";
         h += '<h2>Non-Ventilation Load Equation = ' + Format_NA(llEq.slope, 3) + ' * (ODB - IDB) + ' + Format_NA(llEq.intercept, 1) + '</h2>';
      } else {
         h += '<h1>BIN CALCULATIONS</h1>';
         h += '<h2>Design Conditions --- Candidate Unit</h2>';
         h += '<table>';

         // Row 1: Outdoor Conditions (ASP: ODB, OWB, OHR, ORH, ELV, P)
         h += '<tr>';
         h += "<td class='HeaderDC'>Outdoor Conditions</td>";
         h += "<td class='TitlesDC'>ODB=</td><td class='NormalDataDC'>" + Format_NA(oc.odb, 1) + '</td>';
         h += "<td class='TitlesDC'>OWB=</td><td class='NormalDataDC'>" + Format_NA(oc.owb, 1) + '</td>';
         h += "<td class='TitlesDC'>OHR=</td><td class='NormalDataDC'>" + Format_NA(oc.ohr, 4) + '</td>';
         h += "<td class='TitlesDC'>ORH=</td><td class='NormalDataDC'>" + Format_NA(oc.orh, 1) + '</td>';
         h += "<td class='TitlesDC'>ELV=</td><td class='NormalDataDC'>" + Format_NA(oc.elevation, 0) + '</td>';
         h += "<td class='TitlesDC'>P=</td><td class='NormalDataDC'>" + Format_NA(oc.pressure, 2) + '</td>';
         h += '</tr>';

         // Row 2: Mixed Air Conditions (ASP: EDB, EWB, EHR, ERH, TCF, S/T)
         h += '<tr>';
         h += "<td class='HeaderDC'>Mixed Air Conditions</td>";
         h += "<td class='TitlesDC'>EDB=</td><td class='NormalDataDC'>" + Format_NA(ec.edb, 1) + '</td>';
         h += "<td class='TitlesDC'>EWB=</td><td class='NormalDataDC'>" + Format_NA(ec.ewb, 1) + '</td>';
         h += "<td class='TitlesDC'>EHR=</td><td class='NormalDataDC'>" + Format_NA(ec.ehr, 4) + '</td>';
         h += "<td class='TitlesDC'>ERH=</td><td class='NormalDataDC'>" + Format_NA(ec.erh, 1) + '</td>';
         h += "<td class='TitlesDC'>TCF=</td><td class='NormalDataDC'>" + Format_NA(ec.tcf, 3) + '</td>';
         h += "<td class='TitlesDC'>S/T=</td><td class='NormalDataDC'>" + Format_NA(ec.stRatio, 3) + '</td>';
         h += '</tr>';

         // Row 3: Indoor Conditions (ASP: IDB, IWB, IHR, IRH, BPF-S, BPF-C)
         h += '<tr>';
         h += "<td class='HeaderDC'>Indoor Conditions</td>";
         h += "<td class='TitlesDC'>IDB=</td><td class='NormalDataDC'>" + Format_NA(ic.idb, 1) + '</td>';
         h += "<td class='TitlesDC'>IWB=</td><td class='NormalDataDC'>" + Format_NA(ic.iwb, 1) + '</td>';
         h += "<td class='TitlesDC'>IHR=</td><td class='NormalDataDC'>" + Format_NA(ic.insideHR, 4) + '</td>';
         h += "<td class='TitlesDC'>IRH=</td><td class='NormalDataDC'>" + Format_NA(ic.insideRH * 100, 1) + '</td>';
         h += "<td class='TitlesDC'>BPF-S=</td><td class='NormalDataDC'>" + Format_NA(bpfs.standard, 3) + '</td>';
         h += "<td class='TitlesDC'>BPF-C=</td><td class='NormalDataDC'>" + Format_NA(bpfs.candidate, 3) + '</td>';
         h += '</tr>';

         // Row 4: Equipment Capacity (ASP: RatedSensCap, SupplyFanCFM, VentCFM)
         h += '<tr>';
         h += "<td class='HeaderDC'>Equipment Capacity (kBtuh)</td>";
         h += "<td class='TitlesDC' colspan='3'>Rated Sensible Capacity=</td><td class='NormalDataDC'>" + Format_NA(cap.sensibleAtTest, 1) + '</td>';
         h += "<td class='TitlesDC' colspan='3'>Supply Fan CFM=</td><td class='NormalDataDC'>" + Format_NA(cap.cfm, 0) + '</td>';
         h += "<td class='TitlesDC' colspan='3'>Ventilation CFM=</td><td class='NormalDataDC'>" + Format_NA(cap.ventCFM, 0) + '</td>';
         h += '</tr>';

         // Row 5: Capacity and Load (ASP: SensCap, SensVentLoad, SensNonVentLoad)
         h += '<tr>';
         h += "<td class='HeaderDC'>Capacity and Load (kBtuh)</td>";
         h += "<td class='TitlesDC' colspan='3'>Sensible Capacity=</td><td class='NormalDataDC'>" + Format_NA(cap.sensibleAtDesign, 1) + '</td>';
         h += "<td class='TitlesDC' colspan='3'>Sensible Ventilation Load=</td><td class='NormalDataDC'>" + Format_NA(lds.sensVentLoadDesign, 1) + '</td>';
         h += "<td class='TitlesDC' colspan='3'>Sensible Non-Ventilation Load=</td><td class='NormalDataDC'>" + Format_NA(lds.sensNonVentLoadDesign, 1) + '</td>';
         h += '</tr>';

         // Row 6: Load Equation (ASP: InternalLoadFraction, SensNonVentLoadEquation)
         h += '<tr>';
         h += "<td class='HeaderDC'>Load Equation (kBtuh)</td>";
         h += "<td class='TitlesDC' colspan='3'>Internal Load Fraction=</td><td class='NormalDataDC'>" + Format_NA(lds.sandIfrac, 3) + '</td>';
         h += "<td class='TitlesDC' colspan='4'>Sensible Non-Vent Load Equation=</td>";
         h += "<td class='NormalDataDC' colspan='4'>" + Format_NA(llEq.slope, 3) + ' * (ODB - IDB) + ' + Format_NA(llEq.intercept, 1) + '</td>';
         h += '</tr>';

         h += '</table>';
      }
   }

   // ---- Results (Economics) ----
   h += '<h2>RESULTS</h2>';
   h += '<table>';

   // Header row
   h += '<tr>';
   h += "<th class='Header'>" + cityHeader + "</th>";
   h += "<th class='Header'>Candidate</th>";
   h += "<th class='Header'>Standard</th>";
   h += "<td class='HeaderCenterBold' colspan='2'>Savings</td>";
   h += '</tr>';

   // Annual Energy
   h += '<tr>';
   h += "<td class='TitlesResults'>Annual Energy Consumption (kWhrs)</td>";
   h += "<td class='NormalData'>" + formatNumber(totalC, 0) + "</td>";
   h += "<td class='NormalData'>" + formatNumber(totalS, 0) + "</td>";
   h += "<td class='BoldRed'>" + formatNumber(energySavings, 0) + "</td>";
   h += "<td class='NormalData'>" + pctSav(energySavings, totalS) + "</td>";
   h += '</tr>';

   // Advanced: energy breakdown rows (Condenser, EFan, Aux)
   var advOn = !!(js.meta && js.meta.inputs && js.meta.inputs.advancedControlsChecked);
   if (advOn) {
      var condC = annual.candidate?.condenser || 0;
      var condS = annual.standard?.condenser || 0;
      var efanC = annual.candidate?.efan || 0;
      var efanS = annual.standard?.efan || 0;
      var auxC = annual.candidate?.aux || 0;
      var auxS = annual.standard?.aux || 0;
      var condSav = condS - condC;
      var efanSav = efanS - efanC;
      var auxSav = auxS - auxC;

      h += '<tr>';
      h += "<td class='TitlesResults'>Condenser Unit Energy (kWhrs)</td>";
      h += "<td class='NormalData'>" + formatNumber(condC, 0) + "</td>";
      h += "<td class='NormalData'>" + formatNumber(condS, 0) + "</td>";
      h += "<td class='BoldRed'>" + formatNumber(condSav, 0) + "</td>";
      h += "<td class='NormalData'>" + pctSav(condSav, condS) + "</td>";
      h += '</tr>';

      h += '<tr>';
      h += "<td class='TitlesResults'>Evaporator Fan Energy (kWhrs)</td>";
      h += "<td class='NormalData'>" + formatNumber(efanC, 0) + "</td>";
      h += "<td class='NormalData'>" + formatNumber(efanS, 0) + "</td>";
      h += "<td class='BoldRed'>" + formatNumber(efanSav, 0) + "</td>";
      h += "<td class='NormalData'>" + pctSav(efanSav, efanS) + "</td>";
      h += '</tr>';

      h += '<tr>';
      h += "<td class='TitlesResults'>Aux Electronics Energy (kWhrs)</td>";
      h += "<td class='NormalData'>" + formatNumber(auxC, 0) + "</td>";
      h += "<td class='NormalData'>" + formatNumber(auxS, 0) + "</td>";
      h += "<td class='BoldRed'>" + formatNumber(auxSav, 0) + "</td>";
      h += "<td class='NormalData'>" + pctSav(auxSav, auxS) + "</td>";
      h += '</tr>';
   }

   // Annual Operating Cost
   h += '<tr>';
   h += "<td class='TitlesResults'>Annual Operating Cost ($)</td>";
   h += "<td class='NormalData'>" + formatNumber(annualCost_C * nUnits, 0) + "</td>";
   h += "<td class='NormalData'>" + formatNumber(annualCost_S * nUnits, 0) + "</td>";
   h += "<td class='BoldRed'>" + formatNumber(costAnnualSavings * nUnits, 0) + "</td>";
   h += "<td class='NormalData'>" + pctSav(costAnnualSavings, annualCost_S) + "</td>";
   h += '</tr>';

   // Advanced: Demand Costs
   if (advOn) {
      var demCostC = (econ.candidate?.demandCost || 0) * nUnits;
      var demCostS = (econ.standard?.demandCost || 0) * nUnits;
      var demCostSav = demCostS - demCostC;
      h += '<tr>';
      h += "<td class='TitlesResults'>Demand Costs ($)</td>";
      h += "<td class='NormalData'>" + formatNumber(demCostC, 0) + "</td>";
      h += "<td class='NormalData'>" + formatNumber(demCostS, 0) + "</td>";
      h += "<td class='BoldRed'>" + formatNumber(demCostSav, 0) + "</td>";
      h += "<td class='NormalData'>" + pctSav(demCostSav, demCostS) + "</td>";
      h += '</tr>';
   }

   // Life Cycle Cost
   h += '<tr>';
   h += "<td class='TitlesResults'>" + eqLife + " Year Life Cycle Cost ($)</td>";
   h += "<td class='NormalData'>" + formatNumber(lccC * nUnits, 0) + "</td>";
   h += "<td class='NormalData'>" + formatNumber(lccS * nUnits, 0) + "</td>";
   h += "<td class='BoldRed'>" + formatNumber(lccSavings * nUnits, 0) + "</td>";
   h += "<td class='NormalData'>" + pctSav(lccSavings, lccS) + "</td>";
   h += '</tr>';

   // Annualized Cost
   h += '<tr>';
   h += "<td class='TitlesResults'>Annualized Cost ($)</td>";
   h += "<td class='NormalData'>" + formatNumber(annualizedC * nUnits, 0) + "</td>";
   h += "<td class='NormalData'>" + formatNumber(annualizedS * nUnits, 0) + "</td>";
   h += "<td class='BoldRed'>" + formatNumber(annualizedSavings * nUnits, 0) + "</td>";
   h += "<td class='NormalData'>" + pctSav(annualizedSavings, annualizedS) + "</td>";
   h += '</tr>';

   // Net Present Value
   h += '<tr>';
   h += "<td class='TitlesResults'>Net Present Value ($)</td>";
   h += "<td class='NormalData'>" + formatNumber(npv * nUnits, 0) + "</td>";
   h += "<td style='white-space: nowrap;' colspan='3'>&nbsp;</td>";
   h += '</tr>';

   // Payback
   h += '<tr>';
   if (chartPW) {
      h += "<td class='TitlesResultsBold'>Payback (yrs)</td>";
      h += "<td class='BoldRed'>" + paybackDisplay + "</td>";
   } else {
      h += "<td class='TitlesResultsBold'>Simple Payback (yrs)</td>";
      h += "<td class='BoldRed'>" + paybackDisplay + "</td>";
   }
   h += "<td style='white-space: nowrap;' colspan='3'>&nbsp;</td>";
   h += '</tr>';

   // Rate of Return
   h += '<tr>';
   h += "<td class='TitlesResults'>Rate of Return (%)</td>";
   h += "<td class='NormalData'>" + Format_NA(ror, 2) + "</td>";
   h += "<td style='white-space: nowrap;' colspan='3'>&nbsp;</td>";
   h += '</tr>';

   // SIR
   h += '<tr>';
   h += "<td class='TitlesResults'>Savings to Investment Ratio (SIR)</td>";
   h += "<td class='NormalData'>" + Format_NA(sir, 2) + "</td>";
   h += "<td style='white-space: nowrap;' colspan='3'>&nbsp;</td>";
   h += '</tr>';

   // Locked load line note (ASP Engine.asp lines 2399-2402)
   if (dc.loadLine && dc.loadLine.locked) {
      h += '<tr>';
      h += "<td colspan='5' style='white-space: nowrap;background: #eee; text-align: center;'>Note: the non-ventilation load line is locked</td>";
      h += '</tr>';
   }

   // Payback chart placeholder
   h += '<tr><td colspan="5">';
   h += "<div id='payback_div' style='width: 450px; height: 400px;'></div>";
   h += '</td></tr>';

   h += '</table>';

   // Build payback chart data for later rendering
   var vAxisTitle = chartPW ? 'Discounted Costs ($)' : 'Costs ($)';
   var pbYears = chartPW ? discountedPayback : simplePayback;
   if (pbYears <= 0 || pbYears > eqLife) pbYears = null;
   window._paybackChartData = {
      dataTable: buildPaybackDataTable(unitCost_C, unitCost_S, annualCost_C, annualCost_S, chartPW ? drRate : 0, eqLife, pbYears),
      eqLife: eqLife,
      vAxisTitle: vAxisTitle
   };

   // ---- Parameter Summary ----
   var inp = (js.meta && js.meta.inputs) || {};
   var inpC = inp.candidate || {};
   var inpS = inp.standard || {};

   // WPT_Row helper matching ASP's WPT_Row sub
   function wptRow(name, valC, valS) {
      if (valS === '' || valS === undefined || valS === null) {
         if (valC === '' || valC === undefined || valC === null) valC = 'off';
         return "<tr><td class='TitlesResults'>" + name + "</td><td class='NormalData' colspan='2'>" + valC + "</td></tr>";
      }
      return "<tr><td class='TitlesResults'>" + name + "</td><td class='NormalData'>" + valC + "</td><td class='NormalData'>" + valS + "</td></tr>";
   }

   h += "<div class='graph'>";
   h += '<h2>Parameter Summary</h2>';
   h += '<table>';
   h += "<tr><td class='HeaderRight'>Feature Name</td><td class='HeaderCenter'>Candidate</td><td class='HeaderCenter'>Standard</td></tr>";
   h += wptRow('EER', form['txtEER'].value, form['txtEER_Standard'].value);
   h += wptRow('Unit Cost (k$)', form['txtUnitCost'].value, form['txtUnitCost_Standard'].value);
   h += wptRow('Annual Maintenance Cost ($/year)', form['txtMaintenance_Candidate'].value, form['txtMaintenance_Standard'].value);
   if (advOn) {
      h += wptRow('Specific Candidate Unit', inpC.specificRtu || 'None', 'None');
   }
   h += wptRow('Enable Economizer', form['chkEconomizer_C'].checked ? 'on' : 'off', form['chkEconomizer_S'].checked ? 'on' : 'off');
   if (advOn) {
      h += wptRow('Power -- EFn (kW)', Format_NA(inpC.bfn_kw, 3), Format_NA(inpS.bfn_kw, 3));
      h += wptRow('Power -- Aux (kW)', Format_NA(inpC.aux_kw, 3), Format_NA(inpS.aux_kw, 3));
      h += wptRow('Power -- Cnd (kW)', Format_NA(inpC.cond_kw, 3), Format_NA(inpS.cond_kw, 3));
      h += wptRow('Condenser Fan (%)', String(inpC.condFanPercent != null ? inpC.condFanPercent : 9), String(inpS.condFanPercent != null ? inpS.condFanPercent : 9));
      h += wptRow('Degradation Factor', String(inpC.plDegrFactor != null ? inpC.plDegrFactor : 25), String(inpS.plDegrFactor != null ? inpS.plDegrFactor : 25));
      h += wptRow('Number of stages', String(inpC.nStages || 1), String(inpS.nStages || 1));
      h += wptRow('E-Fan and Condenser', inpC.fanControls || '1-Spd: Always ON', inpS.fanControls || '1-Spd: Always ON');
      h += wptRow('S/T Ratio', Format_NA(inpC.stRatioAtTest, 2), Format_NA(inpS.stRatioAtTest, 2));
      h += wptRow('Spreadsheet data', inpC.spreadsheet ? 'on' : 'off', inpS.spreadsheet ? 'on' : 'off');
   }

   h += "<tr><td class='HeaderCenter'></td><td class='HeaderCenter' colspan='2'>Applies to Both Units</td></tr>";
   if (advOn) {
      h += wptRow('Lock load line', inp.lockLoadLine ? 'on' : 'off', '');
   }
   h += wptRow('Building Type', theSelectedValueInPullDownBox('cmbBuildingType'), '');
   if (advOn) {
      var bmS = inp.bmSlope || '';
      var bmI = inp.bmIntercept || '';
      var bmV = inp.bmVentFrac || '';
      var bmDisplay = (bmS !== '' && bmI !== '' && bmV !== '')
         ? Format_NA(Number(bmS), 2) + ', ' + Format_NA(Number(bmI), 2) + ', ' + Format_NA(Number(bmV), 2)
         : '';
      h += wptRow('Building Model', bmDisplay, '');
   }
   h += wptRow('State, City', stateName + ', ' + theSelectedValueInPullDownBox('cmbCityName2').toUpperCase(), '');
   h += wptRow('Schedule', theSelectedValueInPullDownBox('cmbSchedule'), '');
   var setbackVal = theSelectedValueInPullDownBox('cmbIDB_SetBack');
   var setbackDisplay = (String(setbackVal).trim() === '0') ? 'Cond. Off' : setbackVal;
   h += wptRow('Setpoint Temperature, Setback', theSelectedValueInPullDownBox('cmbIDB') + ', ' + setbackDisplay, '');
   if (advOn) {
      var humTrack = (inp.trackOhr === 'on') ? 'on' : 'off';
      var humRH = (inp.trackOhr === 'on') ? 'NA' : String(inp.irhPct || 60);
      h += wptRow('Auto Humidity, RH Setpoint', humTrack + ', ' + humRH, '');
   }
   h += wptRow('Total Capacity (kBtuh), Oversizing (%)', form['txtTotalCap'].value + ', ' + theSelectedValueInPullDownBox('cmbOversizePercent'), '');
   if (advOn) {
      h += wptRow('Ventilation Rate, Units', String(inp.ventilationValue || '') + ', ' + String(inp.ventilationUnits || '% of Fan Cap.'), '');
      h += wptRow('N for fan energy calcs', String(inp.nAffinity || 2.5), '');
   }
   h += wptRow('Electric Utility Rate ($/kWhrs)', form['txtElectricityRate'].value, '');
   if (advOn) {
      h += wptRow('Demand: Months, Cost ($/kW)', String(econ.demandMonths || 0) + ', ' + String(econ.demandCostPerKW || 0), '');
   }
   h += wptRow('Equipment Life', theSelectedValueInPullDownBox('cmbEquipmentLife'), '');
   h += wptRow('Number of Units', form['txtNUnits'].value, '');

   var drDisplay = chartPW ? drRate : 0;
   h += wptRow('Discounted costs, Rate', (chartPW ? 'on' : 'off') + ', ' + drDisplay, '');

   h += '</table>';
   h += '</div>';

   // ---- Spreadsheet Model Summaries (matches ASP DisplaySpreadsheetAndModel) ----
   var models = (js.meta && js.meta.models) || null;
   if (models) {
      function replaceVarNames(termStr, vars) {
         var s = termStr;
         for (var vi = vars.length; vi >= 1; vi--) {
            s = s.replace(new RegExp('X' + vi, 'g'), vars[vi - 1]);
         }
         return s;
      }
      function renderModelSection(unitLabel, md) {
         if (!md || !md.versionOK) return '';
         var mh = '';
         mh += "<div class='graph'>";
         mh += '<h2>Modeling of Manufacturer\'s Detailed Specification Data (' + unitLabel + ')</h2>';

         // Raw spreadsheet data table (matches ASP ParseSpreadsheetData display)
         if (md.rawSpreadsheet) {
            mh += '<details><summary style="cursor:pointer;font-weight:bold;color:#555;">Raw spreadsheet data from ' + unitLabel + ' Unit</summary>';
            mh += '<table>';
            mh += "<tr><td class='HeaderModel' colspan='7'>Raw spreadsheet data from " + unitLabel + " Unit</td></tr>";
            var ssRows = md.rawSpreadsheet.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
            for (var ri = 0; ri < ssRows.length; ri++) {
               var ssCells = ssRows[ri].split('\t');
               mh += '<tr>';
               for (var ci = 0; ci < ssCells.length; ci++) {
                  mh += "<td class='NormalDataModel'>" + (ssCells[ci] || '') + "</td>";
               }
               mh += '</tr>';
            }
            mh += '</table><br>';
            mh += '</details><br>';
         }

         // Model coefficient tables with t-values, RÂ², and residual SE
         var modelNames = [
            { key: 'grossCapacity', label: 'Gross Total Capacity', vars: ['ODB','EWB','EDB'] },
            { key: 'condenserKW',   label: 'Condenser Power',      vars: ['ODB','EWB','EDB'] },
            { key: 'stRatio',       label: 'Sensible to Total Ratio', vars: ['ODB','EWB','EDB'],
              failMsg: 'Spreadsheet data is insufficient or inappropriate to develop a STR model. The simulation will revert to the ADP/BPF model.' },
            { key: 'neer',          label: 'Normalized EER',       vars: ['LF','ODB'] }
         ];
         for (var mi = 0; mi < modelNames.length; mi++) {
            var mDef = modelNames[mi];
            var m = md.models ? md.models[mDef.key] : null;
            if (m && m.isSolved && m.terms && m.coefficients) {
               mh += '<table>';
               mh += "<tr><td class='HeaderModel' colspan='3'>Model for " + mDef.label + "</td></tr>";
               mh += "<tr><td class='HeaderModel'>Parameter</td><td class='HeaderModel'>Value</td><td class='HeaderModel'>T-Value</td></tr>";
               for (var ti = 0; ti < m.terms.length; ti++) {
                  var termDisplay = replaceVarNames(m.terms[ti], mDef.vars);
                  var coefStr = m.coefficients[ti].toExponential(4);
                  var tStr = (m.tValues && m.tValues[ti] != null) ? Number(m.tValues[ti].toFixed(1)).toLocaleString('en-US', {minimumFractionDigits:1, maximumFractionDigits:1}) : '';
                  mh += "<tr><td class='NormalDataModel'>" + termDisplay + "</td>";
                  mh += "<td class='NormalDataModel'>" + coefStr + "</td>";
                  mh += "<td class='NormalDataModel'>" + tStr + "</td></tr>";
               }
               mh += '</table>';
               // Show model equation (matches ASP ModelDefinitionLocal.Y display)
               var eqParts = [];
               for (var ti = 0; ti < m.terms.length; ti++) {
                  var termDisplay = replaceVarNames(m.terms[ti], mDef.vars);
                  eqParts.push(m.coefficients[ti].toExponential(4) + '*' + termDisplay);
               }
               mh += '<br>' + mDef.label + ' = ' + eqParts.join(' + ');
               // RÂ² and Residual Standard Error (matches ASP lines 607, 610)
               if (m.r2 != null) mh += '<br>R^2 Value = ' + m.r2.toFixed(5);
               if (m.residualSE != null) mh += '<br>Residual Standard Error = ' + m.residualSE.toFixed(5);
               mh += '<br><br>';
            } else {
               var warnText = mDef.failMsg || ('Spreadsheet data is insufficient or inappropriate to develop a ' + mDef.label.toLowerCase() + ' model. The simulation will revert to the standard DOE-2 correction model.');
               mh += "<div style='margin:0 0 1.2em 0'><span style='background-color: yellow'>" + warnText + "</span></div>";
            }
         }
         mh += '</div>';
         return mh;
      }
      var ssOnC = form['chkSpreadsheetControl_C'] && form['chkSpreadsheetControl_C'].checked;
      var ssOnS = form['chkSpreadsheetControl_S'] && form['chkSpreadsheetControl_S'].checked;
      if (models.candidate && ssOnC) h += renderModelSection('Candidate', models.candidate);
      if (models.standard && ssOnS) h += renderModelSection('Standard', models.standard);
   }

   h += '<br><br>';

   // ---- Bin Tables and Charts (after Results and Parameter Summary, matching ASP order) ----
   if (showDetails && dc.outdoor && bins.candidate) {
      window._binChartsQueue = [];
      window._binChartsGridQueue = [];
      var chartIdx = 0;
      var gridIdx = 0;

      // Data for setpoint/design dots on Loads/Hours charts
      var idbSetpoint = dc.indoor ? dc.indoor.idb : 0;
      var idbSetBackVal = theSelectedValueInPullDownBox('cmbIDB_SetBack');
      var idbSetBack = (String(idbSetBackVal).trim().toLowerCase() === 'cond. off') ? 0 : Number(idbSetBackVal);
      if (!Number.isFinite(idbSetBack)) idbSetBack = 0;
      var designOdb = dc.outdoor ? dc.outdoor.odb : 0;
      var designLoad = dc.capacity ? dc.capacity.sensibleAtDesign : 0; // sensibleCapacityDesign = sensLoadAtDesign * LRF

      // Compute shared axis maxima across all bin sets (matching ASP's dblYMax_energy / dblYMax_factor)
      var sharedMaxEnergy = 0.1, sharedMaxFactor = 0.1;
      var allBinArrays = [bins.candidate.occ, bins.candidate.unocc, bins.standard.occ, bins.standard.unocc];
      for (var ai = 0; ai < allBinArrays.length; ai++) {
         var ab = allBinArrays[ai];
         if (!ab) continue;
         for (var aj = 0; aj < ab.length; aj++) {
            if (ab[aj].eCond > sharedMaxEnergy) sharedMaxEnergy = ab[aj].eCond;
            if (ab[aj].eFan > sharedMaxEnergy) sharedMaxEnergy = ab[aj].eFan;
            if (ab[aj].tcf > sharedMaxFactor) sharedMaxFactor = ab[aj].tcf;
         }
      }
      var sharedMaxFactorPct = sharedMaxFactor * 100;

      // Wrap all bin tables + inline charts in a container for toggling
      h += "<div id='binTablesView'>";

      var binSets = [
         { label: 'Candidate', bins: bins.candidate, system: 'candidate' },
         { label: 'Standard', bins: bins.standard, system: 'standard' }
      ];

      for (var si = 0; si < binSets.length; si++) {
         var bset = binSets[si];
         var occBins = bset.bins.occ || [];
         var unoccBins = bset.bins.unocc || [];

         var binGroups = [
            { label: bset.label + ' Unit --- Occupied Hours', data: occBins, occType: 'Occ' },
            { label: bset.label + ' Unit --- Unoccupied Hours', data: unoccBins, occType: 'Unocc' }
         ];

         for (var gi = 0; gi < binGroups.length; gi++) {
            var bg = binGroups[gi];
            var bdata = bg.data;
            if (!bdata || bdata.length === 0) continue;

            var lhDivId = 'binLH_' + chartIdx;
            var epDivId = 'binEP_' + chartIdx;
            chartIdx++;

            h += "<div class='graph'>";

            // Table A: Loads and Conditions
            h += '<h2>BIN Calculations --- ' + bg.label + '</h2>';
            h += "<table class='dataTbl_A'>";
            h += '<tr>';
            h += "<td class='HeaderBC' title='Outside dry bulb (F)'>ODB</td>";
            h += "<td class='HeaderBC' title='Outside mean coincident web bulb (F)'>OWB</td>";
            h += "<td class='HeaderBC' title='Outside humidity ratio'>OHR</td>";
            h += "<td class='HeaderBC' title='Inside humidity ratio'>IHR</td>";
            h += "<td class='HeaderBC' title='Inside relative humidity (%)'>IRH</td>";
            h += "<td class='HeaderBC' title='Number of hours that this outside condition occurs during the specified schedule in one year at the specified location'>Hrs</td>";
            h += "<td class='HeaderBC' title='Non-ventilation sensible load (kBtuh)'>NVSLd</td>";
            h += "<td class='HeaderBC' title='Ventilation sensible load (kBtuh)'>VSLd</td>";
            h += "<td class='HeaderBC' title='Sum of sensible loads, before considering economizer (kBtuh)'>SLd</td>";
            h += "<td class='HeaderBC' title='Economizer sensible capacity: derived from fan capacity [excluding ventilation flow] (kBtuh)'>ESCp</td>";
            h += "<td class='HeaderBC' title='The sensible load after the economizer is considered (kBtuh)'>SLdE</td>";
            h += "<td class='HeaderBC' title='The latent load after the economizer is considered (kBtuh)'>LLdE</td>";
            h += "<td class='HeaderBC' title='Entering, mixed air dry bulb (F)'>EDB</td>";
            h += "<td class='HeaderBC' title='Entering, mixed air humidity ratio'>EHR</td>";
            h += "<td class='HeaderBC' title='Entering, mixed air relative humidity (%)'>ERH</td>";
            h += "<td class='HeaderBC' title='Entering, mixed air wet bulb (F)'>EWB</td>";
            h += "<td class='HeaderBC' title='Correction factor applied to ARI-rated gross total capacity; Used in calculating PCF.'>TCF</td>";
            h += "<td class='HeaderBC' title='Sensible to total capacity ratio; Used with TCF in calculating sensible capacity.'>S/T</td>";
            h += "<td class='HeaderBC' title='Inverse (i.e. Out/In, like an EER) of DOE-2 defined Input Efficiency (In/Out); Used in calculating PCF.'>1/ECF</td>";
            h += '</tr>';

            for (var bi = 0; bi < bdata.length; bi++) {
               var b = bdata[bi];
               h += '<tr>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.odb, 0) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.owb, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.ohr, 4) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.ihr, 4) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(100 * b.irh, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.hours, 0) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.nonVent, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.sensVent, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.totalSens, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.econLoad, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.remaining, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.latentLoad, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.bcc_edb, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.bcc_ehr, 4) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(100 * b.bcc_erh, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.bcc_ewb, 1) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.tcf, 3) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b.stRatio, 3) + '</td>';
               h += "<td class='NormalDataBC'>" + ((b.invEcf !== 0) ? Format_NA(b.invEcf, 3) : '---') + '</td>';
               h += '</tr>';
            }
            h += '</table>';

            // Table B: Performance
            h += '<h2>BIN Calculations --- ' + bg.label + ' (continued)</h2>';
            h += "<table class='dataTbl_B'>";
            h += '<tr>';
            h += "<td class='HeaderBC' title='Outside dry bulb (F)'>ODB</td>";
            h += "<td class='HeaderBC' title='Stage level: for a staged RTU, 2.9 indicates the unit runs at stage level 3 for 90% of the hour and at level 2 for 10%; for a variable-capacity RTU, 0.73 indicates the units is running at 73% of capacity.'>SL</td>";
            h += "<td class='HeaderBC' title='Correction factor applied to ARI-rated condenser power (= TCF/ECF)'>PCF</td>";
            h += "<td class='HeaderBC' title='Overall system correction factor, proportional to energy usage ( = PCF / (TCF * S/T_CF) ), where S/T_CF = (S/T_@EnteringConditions) / (S/T_@ARI)'>OCF</td>";
            h += "<td class='HeaderBC' title='Peak wattage (kW)'>kW</td>";
            h += "<td class='HeaderBC' title='Unit&#39;s auxiliary (electronics) energy consumption (kWhrs)'>E_Aux</td>";
            h += "<td class='HeaderBC' title='Unit&#39;s evaporator-fan energy consumption (kWhrs)'>E_Fan</td>";
            h += "<td class='HeaderBC' title='Unit&#39;s condenser unit (fan and compressor) energy consumption (kWhrs)'>E_Cond</td>";
            h += '</tr>';

            for (var bi2 = 0; bi2 < bdata.length; bi2++) {
               var b2 = bdata[bi2];
               h += '<tr>';
               h += "<td class='NormalDataBC'>" + Format_NA(b2.odb, 0) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b2.stageLevel, 3) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b2.pcf, 3) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b2.ocf, 3) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b2.demandKw, 2) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b2.eAux, 0) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b2.eFan, 0) + '</td>';
               h += "<td class='NormalDataBC'>" + Format_NA(b2.eCond, 0) + '</td>';
               h += '</tr>';
            }
            h += '</table>';

            // Chart divs (Loads/Hours on left, Performance on right)
            h += "<div style='display:flex; gap:16px; flex-wrap:wrap; margin:20px 0 5px 0;'>";
            h += "<div id='" + lhDivId + "' style='width:450px; height:375px;'></div>";
            h += "<div id='" + epDivId + "' style='width:450px; height:375px;'></div>";
            h += "</div>";

            // Queue chart data for drawing after DOM update
            var chartSetpointIdb = (bg.occType === 'Occ') ? idbSetpoint : (idbSetpoint + idbSetBack);
            var loadsEntry = { type: 'loads', divId: lhDivId, data: bdata, title: 'Loads and Hours --- ' + bg.label,
                 chartExtra: { setpointIdb: chartSetpointIdb, designOdb: designOdb, designLoad: designLoad, lockLoadLine: false } };
            var perfEntry = { type: 'perf',  divId: epDivId, data: bdata, title: 'Performance --- ' + bg.label,
                 chartExtra: { yMaxEnergy: sharedMaxEnergy, yMaxFactor: sharedMaxFactorPct } };
            window._binChartsQueue.push(loadsEntry, perfEntry);

            // Queue grid chart entries with system/occType keys for the 4x2 grid
            var gridLhId = 'gridLH_' + bset.system + '_' + bg.occType;
            var gridEpId = 'gridEP_' + bset.system + '_' + bg.occType;
            window._binChartsGridQueue.push(
               { type: 'loads', divId: gridLhId, data: bdata, title: loadsEntry.title, chartExtra: loadsEntry.chartExtra },
               { type: 'perf',  divId: gridEpId, data: bdata, title: perfEntry.title, chartExtra: perfEntry.chartExtra }
            );

            h += '</div>';
         }
      }

      // ---- Definitions Table (inside binTablesView so it hides with tables) ----
      h += "<div class='graph'>";
      h += "<h2>Definitions <span class='smaller'>.........Note: read definitions below OR hover cursor over the table headings above.</span></h2>";
      h += '<table>';
      h += "<tr><td class='TitlesDEF'>ELV</td><td class='ValuesDEF'>Elevation at specified location (feet)</td></tr>";
      h += "<tr><td class='TitlesDEF'>P</td><td class='ValuesDEF'>Standard pressure corrected for elevation (inHg)</td></tr>";
      h += "<tr><td class='TitlesDEF'>ODB</td><td class='ValuesDEF'>Outside dry bulb (F)</td></tr>";
      h += "<tr><td class='TitlesDEF'>OWB</td><td class='ValuesDEF'>Outside mean coincident wet bulb (F)</td></tr>";
      h += "<tr><td class='TitlesDEF'>OHR</td><td class='ValuesDEF'>Outside humidity ratio</td></tr>";
      h += "<tr><td class='TitlesDEF'>IHR</td><td class='ValuesDEF'>Inside humidity ratio</td></tr>";
      h += "<tr><td class='TitlesDEF'>IRH</td><td class='ValuesDEF'>Inside relative humidity (%)</td></tr>";
      h += "<tr><td class='TitlesDEF'>Hrs</td><td class='ValuesDEF'>Number of hours that this outside condition occurs during the specified schedule in one year at the specified location</td></tr>";
      h += "<tr><td class='TitlesDEF'>NVSLd</td><td class='ValuesDEF'>Non-ventilation sensible load (kBtuh)</td></tr>";
      h += "<tr><td class='TitlesDEF'>VSLd</td><td class='ValuesDEF'>Ventilation sensible load (kBtuh)</td></tr>";
      h += "<tr><td class='TitlesDEF'>SLd</td><td class='ValuesDEF'>Sum of sensible loads, before considering economizer (kBtuh)</td></tr>";
      h += "<tr><td class='TitlesDEF'>ESCp</td><td class='ValuesDEF'>Economizer sensible capacity: derived from fan capacity [excluding ventilation flow] (kBtuh)</td></tr>";
      h += "<tr><td class='TitlesDEF'>SLdE</td><td class='ValuesDEF'>The sensible load after the economizer is considered (kBtuh)</td></tr>";
      h += "<tr><td class='TitlesDEF'>LLdE</td><td class='ValuesDEF'>The latent load after the economizer is considered (kBtuh)</td></tr>";
      h += "<tr><td class='TitlesDEF'>EDB</td><td class='ValuesDEF'>Entering, mixed air dry bulb (F)</td></tr>";
      h += "<tr><td class='TitlesDEF'>EHR</td><td class='ValuesDEF'>Entering, mixed air humidity ratio</td></tr>";
      h += "<tr><td class='TitlesDEF'>ERH</td><td class='ValuesDEF'>Entering, mixed air relative humidity (%)</td></tr>";
      h += "<tr><td class='TitlesDEF'>EWB</td><td class='ValuesDEF'>Entering, mixed air wet bulb (F)</td></tr>";
      h += "<tr><td class='TitlesDEF'>TCF</td><td class='ValuesDEF'>Correction factor applied to ARI-rated gross total capacity; Used in calculating PCF.</td></tr>";
      h += "<tr><td class='TitlesDEF'>BPF</td><td class='ValuesDEF'>ByPass Factor (BPF) adjusted for mass flow at bin conditions.</td></tr>";
      h += "<tr><td class='TitlesDEF'>S/T</td><td class='ValuesDEF'>Sensible to total capacity ratio; Used with TCF in calculating sensible capacity.</td></tr>";
      h += "<tr><td class='TitlesDEF'>1/ECF</td><td class='ValuesDEF'>Inverse (i.e. Out/In, like an EER) of DOE-2 defined Input Efficiency (In/Out); Used in calculating PCF.</td></tr>";
      h += "<tr><td class='TitlesDEF'>SL</td><td class='ValuesDEF'>Stage level: for a staged RTU, 2.9 indicates the unit runs at stage level 3 for 90% of the hour and at level 2 for 10%; for a variable-capacity RTU, 0.73 indicates the unit is running at 73% of capacity.</td></tr>";
      h += "<tr><td class='TitlesDEF'>PCF</td><td class='ValuesDEF'>Correction factor applied to ARI-rated condenser power (= TCF/ECF)</td></tr>";
      h += "<tr><td class='TitlesDEF'>OCF</td><td class='ValuesDEF'>Overall system correction factor, proportional to energy usage ( = PCF / (TCF * S/T_CF) ), where S/T_CF = (S/T_@EnteringConditions) / (S/T_@ARI)</td></tr>";
      h += "<tr><td class='TitlesDEF'>kW</td><td class='ValuesDEF'>Peak wattage (kW)</td></tr>";
      h += "<tr><td class='TitlesDEF'>E_Aux</td><td class='ValuesDEF'>Unit's auxiliary (electronics) energy consumption (kWhrs)</td></tr>";
      h += "<tr><td class='TitlesDEF'>E_Fan</td><td class='ValuesDEF'>Unit's evaporator-fan energy consumption (kWhrs)</td></tr>";
      h += "<tr><td class='TitlesDEF'>E_Cond</td><td class='ValuesDEF'>Unit's condenser unit (fan and compressor) energy consumption (kWhrs)</td></tr>";
      h += '</table>';
      h += '</div>';

      // Close binTablesView wrapper (opened before bin tables loop)
      h += '</div>';

      // (toggle link shown by submitToEngine after innerHTML is set)
   }

   // ---- Chart Grid View (hidden by default, toggled by link) ----
   // Placed after all floated sections; clear:both ensures it starts below Parameter Summary etc.
   if (showDetails && dc.outdoor && bins.candidate) {
      h += "<div style='clear:both;'></div>";
      h += "<div id='binChartsGrid' style='display:none;'>";
      h += "<br><br>";
      var gridRows = [
         { label: 'Loads --- Occupied Hours', occType: 'Occ', chartType: 'LH' },
         { label: 'Performance --- Occupied Hours', occType: 'Occ', chartType: 'EP' },
         { label: 'Loads --- Unoccupied Hours', occType: 'Unocc', chartType: 'LH' },
         { label: 'Performance --- Unoccupied Hours', occType: 'Unocc', chartType: 'EP' }
      ];
      for (var gr = 0; gr < gridRows.length; gr++) {
         var grow = gridRows[gr];
         h += "<div style='display:flex; gap:16px; flex-wrap:nowrap; margin:10px 0;'>";
         h += "<div id='grid" + grow.chartType + "_candidate_" + grow.occType + "' style='min-width:450px; width:450px; height:375px;'></div>";
         h += "<div id='grid" + grow.chartType + "_standard_" + grow.occType + "' style='min-width:450px; width:450px; height:375px;'></div>";
         h += "</div>";
      }
      h += '</div>';
   }

   return h;
}

// ---- Initialize on page load ----
// loadCityData() must complete before recalcVentilation() can run because the
// engine needs design-condition weather data from Stations.json.  ASP's
// Establish_SIV runs server-side after all form values (including city data)
// are available; this is the JS equivalent.
loadCityData().then(function() { recalcVentilation(); });