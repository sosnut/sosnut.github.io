function text_to_excel() {
	var text = document.getElementById('text-to-convert').value;
	if (text != "" && text != null) {
		var splitter;
		let splitterForm = document.getElementById('form-delimiter');
		let splitterType = splitterForm.options[splitterForm.selectedIndex].value;
		switch (splitterType) {
			case 'comma':
				splitter = /,/g;
				text = text.replace(/,/g, ' ,');
				break;
			case 'semicolon':
				splitter = /;/g;
				text = text.replace(/;/g, ' ;');
				break;
			case 'tab':
			default:
				splitter = /\t/g;
				text = text.replace(/\t/g, ' \t');
				break;
		}
		text = text.replace(/\n/g, ' \n');
		text = text.split("\n");
		for (var i = 0; i < text.length; i++) {
			var divider = isHeaderRow && i === 0 ? '||' : '|';
			text[i] = divider + text[i].replace(splitter, divider) + divider;
			var columns = text[i].split(divider);
			for (var j = 0; j < columns.length; j++) {
				var columnTrim = columns[j].trim();
				if (italicizeNull && columnTrim === 'NULL') {
					columnTrim = '_' + columnTrim + '_';
				}
				columns[j] = columnTrim !== '' ? columnTrim : ' ';
			}
			text[i] = columns.join(divider).trim();
			if (isHeaderColumn && ((isHeaderRow && i !== 0) || !isHeaderRow)) {
				text[i] = '|' + text[i];
			}
		}
		text = text.join("\r\n");
	}
	document.getElementById('jira-result').innerHTML = text;
}

function select_result_content() {
    document.getElementById('conversion-result').select();
}

var ALL_PARAMETERS = "Date validation;Numéro rec;Client;Caractéristiques;Cuve;Ex Cuve;Volume;Couleur;AA;SO2T_S;SO2L_S;SO2_ACTIF;TAV_IR;GF_IR;GF_IRVF;GLUC_FRUCT;PH_IR;AT_IR;IC;NUANCE;ACD_MAL;ACD_LACT_G_L;IPT_IR;DENSITE_IR;CO2_IR;FER_S;CU_S;TEST_PROT;TEST_ROSIS;DO420_IR2;DO520_IR2;DO620_IR2;TURBIDITE;IOTA_IR;ALCOOL_PUISSANCE_IRD;ALCOOL_TOTAL_IRD;OOL_TOTAL;IC2;AT_MEQ;AA_MEQ;ACD_MAL_IR;ACD_MAL_REPRISE;ACD_SORBIQUE;ACD_TART_IR;ACIDES_AMINES_IR;ALCOOL_PROB_1500;ALCOOL_PROB_1683;ALCOOL_TOTAL_REF_GF;ALCOOL_TOTAL_VIN;ALUMINIUM;AMMONIUM_IR;AT_MANU;AV_MANU;AZOTE_ASSIMILABLE_IR;b;C_E;C_INF;C_INIT;CARBAMATE_ETHYLE;CF;CF_CONCLUSION;CF_OBSERVATIONS;CL_HL;CMC;CO2_CORNING;COEF_CONV;COLMATAGE;CONDUCTIVITE;CU;CU_TOTAL;CU_VDN;CUIVRE_FERRO;DEGRE_BRIX;DEGRE_POTENTIEL;DO280_MANU;DO420_IR;DO420_MANU;DO520_IR;DO520_MANU;DO620_IR;DO620_MANU;EA_IR;EST;ETHANAL;ETHANAL_LIBRE;EXTRAIT_SEC_REDUIT;FER;FER_FERRO;GF_REPRISE;IC_MANU;IC_MANU_CP;IGA_BACT_10J;IGA_BACT_72;IGA_LEVURES_10J;IGAC_BACTERIES;IGAC_LEVURES;IGAPLUS_BACT_72;IGAPLUS_BACT_72bis;IGAPLUS_BACT_7J;IGAPLUS_BACT_7Jbis;IGAPLUS_LEVURES_7J;IGAPLUS_LEVURES_7Jbis;IGC_BACTERIES;IGC_LEVURES;IPT_GLORIES_IR;K;K_IR;K_IR2;KPA1;KPA23;L;MASSE_VOL_G_L;METHANOL;MP_IR;NUANCE_MANU;OTA_HPLC;OXYGENE;PH_MANU;PLOMB;POIDS_RAISINS;RECHERCH_FERRO;REEDITION_CP;SO2L_REF;SO2T_REF;SORBIQUE;SORBIQUE2;STAB_COLLOIDALE;STABILITE;SUCRES_IR;SUCRES_TOTAUX;SULFATES;SURPRESSION;SURPRESSION_CALCUL;TAUX_CHUTE;TAV_PAAR;TEMPS_INDICATIF;TEST_BDB;TRAITE_FERRO;TURBIDITE_APRES_KPA;TURBIDITE_AVANT_KPA;TURBIDITE_FRIGO;a;A_PH1_IR;A_PH32_IR";


function execute_conversion() {
    var csvData = document.getElementById('text-to-convert').value;
    var conversionResult = '';
    if (csvData != null && csvData != "") {
        var lines = parseIntoLines(csvData);
        if (lines.length > 0) {
            conversionResult = normalize(ALL_PARAMETERS, lines);
        }
    } else {
        document.getElementById('message-area').innerHTML = '';
    }
    document.getElementById('conversion-result').innerHTML = conversionResult;
}

function normalize(allParameters, lines) {
    var paramsArray = allParameters.split(";");
    
    var maps = parseCsvLines(lines, paramsArray);
    var normalizedLines = buildNormalizedLines(paramsArray, maps);

    var conversionResult = paramsArray.join("\t") + "\n" + normalizedLines.join("\n");
    return conversionResult;
}

function parseIntoLines(csvData) {
    var lines = csvData.split("\n").map(line => line.trim());
    return lines.filter(line => line.length > 0);
}

function parseCsvLines(lines, paramsArray) {
    var maps = [];
    var header = lines[0];
    var columnNames = parseCSVHeader(header);
    checkColumnNames(columnNames, paramsArray);

    var i;
    for (i = 1; i < lines.length; i++) {
        var map = new Map();
        var line = lines[i];
        var cells = line.split(";");
        for (j = 0; j < columnNames.length; j++) {
            var columnName = columnNames[j];
            var cell = cells[j];
            if (columnName) {
                map.set(columnName, cell);
            }
        }
        maps.push(map);
    }
    return maps;
}

function parseCSVHeader(header) {
    var columnNames = header.split(";");
    if (columnNames[columnNames.length - 1] === '') {
        columnNames.pop();
    }
    return columnNames;
}

function checkColumnNames(columnNames, paramsArray) {
    var unknownColumns = [];
    var message = "";
    var i;
    for (i = 0; i < columnNames.length; i++) {
        var columnName = columnNames[i];
        if (!(paramsArray.includes(columnName))) {
            unknownColumns.push(columnName);
            console.log("paramètre inconnu en colonne " + (i+1) + " : " + columnName);
        }
    }
    if (unknownColumns.length) {
        var message = "Avertissement : " + (unknownColumns.length == 1 ? "paramètre inconnu" : "paramètres inconnus")+": " + unknownColumns;
        message += "\n" + "(placé" + (unknownColumns.length == 1 ? "" : "s") + " en fin de liste)";
        for (u of unknownColumns.sort()) {
            paramsArray.push(u);
        }
    }
    document.getElementById('message-area').innerHTML = message;
}

function buildNormalizedLines(paramsArray, maps) {
    var paramValuesLines = [];
    for (map of maps) {
        var paramValues = [];
        for (param of paramsArray) {
            paramValues.push(map.get(param));
        }
        var tabSeparatedValues = paramValues.join("\t");
        paramValuesLines.push(tabSeparatedValues);
    }
    return paramValuesLines;
}

function testOnData() {
    var allParameters = "Date validation;Numéro rec;Client;Caractéristiques;Cuve;Volume;Couleur;a;A_PH1_IR;A_PH32_IR;AA;ACD_LACT_G_L;ACD_MAL;ACD_MAL_IR;ACD_MAL_REPRISE;ACD_SORBIQUE;ACD_TART_IR;ACIDES_AMINES_IR;ALCOOL_PROB_1500;ALCOOL_PROB_1683;ALCOOL_TOTAL_IRD;ALCOOL_TOTAL_REF_GF;ALCOOL_TOTAL_VIN;ALUMINIUM;AMMONIUM_IR;AT_IR;AT_MANU;AV_MANU;AZOTE_ASSIMILABLE_IR;b;C_E;C_INF;C_INIT;CARBAMATE_ETHYLE;CF;CF_CONCLUSION;CF_OBSERVATIONS;CL_HL;CMC;CO2_CORNING;CO2_IR;COEF_CONV;COLMATAGE;CONDUCTIVITE;CU;CU_TOTAL;CU_VDN;CUIVRE_FERRO;DEGRE_BRIX;DEGRE_POTENTIEL;DENSITE_IR;DO280_MANU;DO420_IR;DO420_MANU;DO520_IR;DO520_MANU;DO620_IR;DO620_MANU;EA_IR;EST;ETHANAL;ETHANAL_LIBRE;EXTRAIT_SEC_REDUIT;FER;FER_FERRO;GF_IR;GF_IRVF;GF_REPRISE;GLUC_FRUCT;IC;IC_MANU;IC_MANU_CP;IC2;IGA_BACT_10J;IGA_BACT_72;IGA_LEVURES_10J;IGAC_BACTERIES;IGAC_LEVURES;IGAPLUS_BACT_72;IGAPLUS_BACT_72bis;IGAPLUS_BACT_7J;IGAPLUS_BACT_7Jbis;IGAPLUS_LEVURES_7J;IGAPLUS_LEVURES_7Jbis;IGC_BACTERIES;IGC_LEVURES;IOTA_IR;IPT_GLORIES_IR;IPT_IR;K;K_IR;K_IR2;KPA1;KPA23;L;MASSE_VOL_G_L;METHANOL;MP_IR;NUANCE_MANU;OTA_HPLC;OTA_IR;OXYGENE;PH_IR;PH_MANU;PLOMB;POIDS_RAISINS;RECHERCH_FERRO;REEDITION_CP;SO2_ACTIF;SO2L_REF;SO2L_S;SO2T_REF;SO2T_S;SORBIQUE;SORBIQUE2;STAB_COLLOIDALE;SUCRES_IR;SUCRES_TOTAUX;SULFATES;SURPRESSION;SURPRESSION_CALCUL;TAUX_CHUTE;TAV_IR;TAV_PAAR;TEST_BDB;TEST_PROT;TEST_ROSIS;TRAITE_FERRO;TURBIDITE;TURBIDITE_APRES_KPA;TURBIDITE_AVANT_KPA;TURBIDITE_FRIGO"
    const csvData = "Date validation;Numéro rec;Client;Caractéristiques;Cuve;Volume;Couleur;AA;SO2T_S;SO2L_S;SO2_ACTIF;TAV_IR;GF_IRVF;PH_IR;AT_IR;IC;NUANCE;TEST_ROSIS;ACD_MAL;ACD_LACT_G_L;IPT_IR;DENSITE_IR;CO2_IR;FER_S;CU_S;TURBIDITE;TEST_PROT;COEF_CONV;C_INIT;C_INF;C_E;TAUX_CHUTE;STABILITE;TEMPS_INDICATIF;DO420_IR2;DO520_IR2;DO620_IR2;" + "\r\n" + "25/06/2020 18:06:36;VS-0002 ;ICB0006;NATURE DU VIN ;CUVE;100.000000;3;0.50;100;10;0.83;12.00;5.0;3.00;5.00;3.0;1.0;;;;45;;500;5.0;;;;;;;;;;;1.00;1.00;1.00;" + "\r\n" + "25/06/2020 18:06:36;VS-0003 ;ICB0006;NATURE DU VIN ;CUVE;100.000000;3;0.50;100;10;0.83;12.00;< 1;3.00;5.00;3.0;1.0;;;;;;500;5.0;< 0.3;;5.0;;;;;;;;1.00;1.00;1.00;";

    var lines = parseIntoLines(csvData);
    var conversionResult = normalize(allParameters, lines);

    var expected = "Date validation\tNuméro rec\tClient\tCaractéristiques\tCuve\tVolume\tCouleur\ta\tA_PH1_IR\tA_PH32_IR\tAA\tACD_LACT_G_L\tACD_MAL\tACD_MAL_IR\tACD_MAL_REPRISE\tACD_SORBIQUE\tACD_TART_IR\tACIDES_AMINES_IR\tALCOOL_PROB_1500\tALCOOL_PROB_1683\tALCOOL_TOTAL_IRD\tALCOOL_TOTAL_REF_GF\tALCOOL_TOTAL_VIN\tALUMINIUM\tAMMONIUM_IR\tAT_IR\tAT_MANU\tAV_MANU\tAZOTE_ASSIMILABLE_IR\tb\tC_E\tC_INF\tC_INIT\tCARBAMATE_ETHYLE\tCF\tCF_CONCLUSION\tCF_OBSERVATIONS\tCL_HL\tCMC\tCO2_CORNING\tCO2_IR\tCOEF_CONV\tCOLMATAGE\tCONDUCTIVITE\tCU\tCU_TOTAL\tCU_VDN\tCUIVRE_FERRO\tDEGRE_BRIX\tDEGRE_POTENTIEL\tDENSITE_IR\tDO280_MANU\tDO420_IR\tDO420_MANU\tDO520_IR\tDO520_MANU\tDO620_IR\tDO620_MANU\tEA_IR\tEST\tETHANAL\tETHANAL_LIBRE\tEXTRAIT_SEC_REDUIT\tFER\tFER_FERRO\tGF_IR\tGF_IRVF\tGF_REPRISE\tGLUC_FRUCT\tIC\tIC_MANU\tIC_MANU_CP\tIC2\tIGA_BACT_10J\tIGA_BACT_72\tIGA_LEVURES_10J\tIGAC_BACTERIES\tIGAC_LEVURES\tIGAPLUS_BACT_72\tIGAPLUS_BACT_72bis\tIGAPLUS_BACT_7J\tIGAPLUS_BACT_7Jbis\tIGAPLUS_LEVURES_7J\tIGAPLUS_LEVURES_7Jbis\tIGC_BACTERIES\tIGC_LEVURES\tIOTA_IR\tIPT_GLORIES_IR\tIPT_IR\tK\tK_IR\tK_IR2\tKPA1\tKPA23\tL\tMASSE_VOL_G_L\tMETHANOL\tMP_IR\tNUANCE_MANU\tOTA_HPLC\tOTA_IR\tOXYGENE\tPH_IR\tPH_MANU\tPLOMB\tPOIDS_RAISINS\tRECHERCH_FERRO\tREEDITION_CP\tSO2_ACTIF\tSO2L_REF\tSO2L_S\tSO2T_REF\tSO2T_S\tSORBIQUE\tSORBIQUE2\tSTAB_COLLOIDALE\tSUCRES_IR\tSUCRES_TOTAUX\tSULFATES\tSURPRESSION\tSURPRESSION_CALCUL\tTAUX_CHUTE\tTAV_IR\tTAV_PAAR\tTEST_BDB\tTEST_PROT\tTEST_ROSIS\tTRAITE_FERRO\tTURBIDITE\tTURBIDITE_APRES_KPA\tTURBIDITE_AVANT_KPA\tTURBIDITE_FRIGO\tCU_S\tDO420_IR2\tDO520_IR2\tDO620_IR2\tFER_S\tNUANCE\tSTABILITE\tTEMPS_INDICATIF\n25/06/2020 18:06:36\tVS-0002 \tICB0006\tNATURE DU VIN \tCUVE\t100.000000\t3\t\t\t\t0.50\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t5.00\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t500\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t5.0\t\t\t3.0\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t45\t\t\t\t\t\t\t\t\t\t\t\t\t\t3.00\t\t\t\t\t\t0.83\t\t10\t\t100\t\t\t\t\t\t\t\t\t\t12.00\t\t\t\t\t\t\t\t\t\t\t1.00\t1.00\t1.00\t5.0\t1.0\t\t\n25/06/2020 18:06:36\tVS-0003 \tICB0006\tNATURE DU VIN \tCUVE\t100.000000\t3\t\t\t\t0.50\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t5.00\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t500\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t< 1\t\t\t3.0\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t3.00\t\t\t\t\t\t0.83\t\t10\t\t100\t\t\t\t\t\t\t\t\t\t12.00\t\t\t5.0\t\t\t\t\t\t\t< 0.3\t1.00\t1.00\t1.00\t5.0\t1.0\t\t";

    assertEquals("Wrong conversion result", expected, conversionResult);
}

function assertEquals(message, expected, actual) {
    if (!(expected === actual)) {
        document.write("<strong>False assertion</strong><br/>\n");
        document.write(message + "<br/>\n");
        document.write("Expected: " + expected.length + " characters<br/>\n");
        document.write("Actual:   " + actual.length + " characters<br/>\n");
        document.write("");
        throw new Error("Assertion error: " + message);
    }
}
