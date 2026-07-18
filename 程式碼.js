// ======================== 🛡️ 安全與通訊金鑰配置 ========================
var SCRIPT_PROPS = PropertiesService.getScriptProperties();
var WEB_API_KEY = SCRIPT_PROPS.getProperty("WEB_API_KEY");
var LINE_ACCESS_TOKEN = SCRIPT_PROPS.getProperty("LINE_ACCESS_TOKEN");
var LINE_USER_ID = SCRIPT_PROPS.getProperty("LINE_USER_ID");

var PLANT_CODE_MAP = { "茉莉": "J", "桂花": "O", "朝天椒": "P" };

var SENSOR_LAT = Number(SCRIPT_PROPS.getProperty("SENSOR_LAT"));
var SENSOR_LON = Number(SCRIPT_PROPS.getProperty("SENSOR_LON"));
var TIMEZONE_OFFSET = 8;

// BH1750擴散罩衰減校正係數(2026/07/15用7/6~7/8裸機晴天vs7/15加罩晴天跨日比對修正，
// 單點峰值比對與DLI全日積分反推兩種獨立方法收斂在2.89~2.91，故定為2.9)
var LIGHT_CORRECTION_FACTOR = 2.9;

// ---- 仰角動態調整因子（實驗性，2026/07/17新增，僅寫入測試欄位S，尚未取代正式F欄）----
// 依據：陽台原始值/官方站日射量(W/m2)比值，僅取直曬窗內樣本迴歸，對數擬合 R²=0.99
// ratio(elev) = -17.95*ln(elev) + 103.26
// 基準仰角87°（對應2.9係數當初校準時的正午峰值情境），該處adjustFactor=1.0
var ELEV_ADJ_REFERENCE = 87;
var ELEV_ADJ_MIN = 0.7;
var ELEV_ADJ_MAX = 2.0;

function getElevationAdjustmentFactor_(elevation, azimuth) {
    if (!isBalconyLit_(elevation, azimuth)) return 1; // 窗外（遮蔽/散射）暫不修正
    var ratio = -17.95 * Math.log(elevation) + 103.26;
    var refRatio = -17.95 * Math.log(ELEV_ADJ_REFERENCE) + 103.26;
    var factor = ratio / refRatio;
    return Math.max(ELEV_ADJ_MIN, Math.min(ELEV_ADJ_MAX, factor));
}

// ======================== 📡 核心 1：ESP32 打卡大門 (doGet) ========================
function doGet(e) {
    // ---- AI 匯出模式：獨立驗證，不動用主系統的 WEB_API_KEY ----
    if (e && e.parameter && e.parameter.mode === "export") {
        var exportKey = SCRIPT_PROPS.getProperty("EXPORT_API_KEY");
        if (!exportKey || e.parameter.key !== exportKey) {
            return ContentService.createTextOutput("UNAUTHORIZED")
                .setMimeType(ContentService.MimeType.TEXT);
        }
        return handleAIExport_(e);
    }

    // ---- 以下為原本 ESP32 打卡邏輯 ----
    var now = new Date();
    var nowTime = now.getTime();
    var dateStr = Utilities.formatDate(now, "GMT+8", "yyyy/MM/dd");
    var timeStr = Utilities.formatDate(now, "GMT+8", "HH:mm:ss");

    if (!e || !e.parameter || e.parameter.key !== WEB_API_KEY) {
        console.warn("拒絕未授權的存取請求！");
        return ContentService.createTextOutput("UNAUTHORIZED");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var totalDataSheet = ss.getSheetByName("資訊總表");
    var dliSheet = ss.getSheetByName("即時訊息");

    var cleanTemp = sanitizeInput(e.parameter.temp, true);
    var cleanHum = sanitizeInput(e.parameter.hum, true);
    var cleanPres = sanitizeInput(e.parameter.pres, true);
    var cleanLight = sanitizeInput(e.parameter.light, true);
    var cleanSoilJ = sanitizeInput(e.parameter.soilJ, true);
    var cleanSoilO = sanitizeInput(e.parameter.soilO, true);
    var cleanSoilP = sanitizeInput(e.parameter.soilP, true);
    var cleanStatus = sanitizeInput(e.parameter.status || "正常", false);

    var cleanSoilAgeJ = sanitizeInput(e.parameter.soilAgeJ, true);
    var cleanSoilAgeO = sanitizeInput(e.parameter.soilAgeO, true);
    var cleanSoilAgeP = sanitizeInput(e.parameter.soilAgeP, true);

    // ---- 第二顆 BH1750 光感器（ADDR接3.3V，位址0x5C）----
    var cleanLight2 = sanitizeInput(e.parameter.light2, true);
    var rawLight2 = "";

    var calcPPFD = "";
    var rawLightUncorrected = "";
    var correctedLight = "";
    var solarElevation = "";
    var lightConfidence = "";
    var elevAdjustedLight = "";   // 仰角修正測試值(寫入S欄，不影響正式F欄)

    if (cleanPres <= 900 || cleanPres > 1100) { cleanPres = ""; }

    // BH1750休眠中(-2)是韌體主動判斷、非異常，跟「讀值異常」(<0其他情況)區分開，
    // 讓試算表能顯示「夜間休眠中」而不是被歸類成感測器故障
    if (cleanLight === -2) {
        cleanLight = "";
        calcPPFD = "";
        rawLightUncorrected = "夜間休眠中";
        solarElevation = Math.round(calculateSolarElevation_(now, SENSOR_LAT, SENSOR_LON) * 10) / 10;
        lightConfidence = "";
    } else if (cleanLight <= -1) {
        cleanLight = "";
        calcPPFD = "";
    } else {
        if (cleanLight < 1) { cleanLight = 0; }
        rawLightUncorrected = cleanLight;

        var pos = calculateSolarPosition_(now, SENSOR_LAT, SENSOR_LON);
        solarElevation = Math.round(pos.elevation * 10) / 10;

        correctedLight = Math.round(cleanLight * LIGHT_CORRECTION_FACTOR * 10) / 10;
        calcPPFD = (correctedLight * 0.0185).toFixed(1);

        var elevFactor = getElevationAdjustmentFactor_(pos.elevation, pos.azimuth);
        elevAdjustedLight = Math.round(cleanLight * LIGHT_CORRECTION_FACTOR * elevFactor * 10) / 10;

        lightConfidence = calculateLightConfidence_(solarElevation);
    }

    // 第二顆光感器：沿用同樣的休眠(-2)/異常(-1)判斷邏輯，僅寫入原始值(不做校正)
    if (cleanLight2 === -2) {
        rawLight2 = "夜間休眠中";
    } else if (cleanLight2 <= -1) {
        rawLight2 = "";
    } else {
        if (cleanLight2 < 1) { cleanLight2 = 0; }
        rawLight2 = cleanLight2;
    }

    if (cleanSoilJ === 100) cleanSoilJ = "";
    if (cleanSoilO === 100) cleanSoilO = "";
    if (cleanSoilP === 100) cleanSoilP = "";

    var storedJ = SCRIPT_PROPS.getProperty("BASE_J");
    var storedO = SCRIPT_PROPS.getProperty("BASE_O");
    var storedP = SCRIPT_PROPS.getProperty("BASE_P");
    var baseJ = (storedJ !== null) ? Number(storedJ) : (cleanSoilJ === "" ? 0 : cleanSoilJ);
    var baseO = (storedO !== null) ? Number(storedO) : (cleanSoilO === "" ? 0 : cleanSoilO);
    var baseP = (storedP !== null) ? Number(storedP) : (cleanSoilP === "" ? 0 : cleanSoilP);

    var lastResetDay = SCRIPT_PROPS.getProperty("LAST_RESET_DAY") || "";
    var currentHour = now.getHours();
    var currentRows = totalDataSheet.getLastRow();

    if (currentHour >= 6 && lastResetDay !== dateStr) {
        if (cleanSoilJ !== "" && cleanSoilO !== "" && cleanSoilP !== "") {
            baseJ = cleanSoilJ; baseO = cleanSoilO; baseP = cleanSoilP;
            SCRIPT_PROPS.setProperties({
                "BASE_J": baseJ.toString(), "BASE_O": baseO.toString(), "BASE_P": baseP.toString(), "LAST_RESET_DAY": dateStr
            });
        }
    }

    if (currentRows >= 2) {
        var prevSoilJ = totalDataSheet.getRange(2, 8).getValue();
        var prevSoilO = totalDataSheet.getRange(2, 9).getValue();
        var prevSoilP = totalDataSheet.getRange(2, 10).getValue();
        if (cleanSoilJ !== "" && prevSoilJ !== "") { if (cleanSoilJ - Number(prevSoilJ) > 10) { baseJ = cleanSoilJ; SCRIPT_PROPS.setProperty("BASE_J", baseJ.toString()); } }
        if (cleanSoilO !== "" && prevSoilO !== "") { if (cleanSoilO - Number(prevSoilO) > 10) { baseO = cleanSoilO; SCRIPT_PROPS.setProperty("BASE_O", baseO.toString()); } }
        if (cleanSoilP !== "" && prevSoilP !== "") { if (cleanSoilP - Number(prevSoilP) > 10) { baseP = cleanSoilP; SCRIPT_PROPS.setProperty("BASE_P", baseP.toString()); } }
    }

    var cumJ = (cleanSoilJ === "") ? 0 : Math.max(0, baseJ - cleanSoilJ);
    var cumO = (cleanSoilO === "") ? 0 : Math.max(0, baseO - cleanSoilO);
    var cumP = (cleanSoilP === "") ? 0 : Math.max(0, baseP - cleanSoilP);

    totalDataSheet.insertRowBefore(2);

    totalDataSheet.getRange(2, 1, 1, 14).setValues([[
        dateStr, timeStr,
        cleanTemp, cleanHum, cleanPres,
        correctedLight, calcPPFD,
        cleanSoilJ, cleanSoilO, cleanSoilP,
        cleanStatus,
        Math.round(cumJ), Math.round(cumO), Math.round(cumP)
    ]]);

    // P/Q/R/S欄：原始光照 / 太陽仰角 / 光照可信度 / 仰角修正測試值(S欄為新增)
    totalDataSheet.getRange(2, 16, 1, 4).setValues([[
        rawLightUncorrected,
        solarElevation,
        lightConfidence,
        elevAdjustedLight
    ]]);

    // T欄：第二顆光感器(ADDR接3.3V，位址0x5C)原始lux值
    totalDataSheet.getRange(2, 20).setValue(rawLight2);

    // 即時訊息新版面：A2~E2=日期時間溫濕氣壓、F2~H2=三株土壤濕度、
    // C4~D4=光照與PPFD(獨立寫入，不與row2連續)
    dliSheet.getRange(2, 1, 1, 5).setValues([[dateStr, timeStr, cleanTemp, cleanHum, cleanPres]]);
    dliSheet.getRange(2, 6, 1, 3).setValues([[cleanSoilJ, cleanSoilO, cleanSoilP]]);
    dliSheet.getRange("C4:D4").setValues([[correctedLight, calcPPFD]]);

    // ESP32通訊區：I2=設備狀態、J2=重置狀態、I4=馬達指令狀態
    dliSheet.getRange("I2").setValue(cleanStatus);

    var updatedRows = totalDataSheet.getLastRow();
    var maxRowsAllowed = 14401;
    var purgeDays = 3;
    var purgeRowCount = purgeDays * 1440;

    if (updatedRows > maxRowsAllowed) {
        var deleteStartRow = updatedRows - purgeRowCount + 1;
        totalDataSheet.deleteRows(deleteStartRow, purgeRowCount);
    }

    var lastDataTimeStr = SCRIPT_PROPS.getProperty("LAST_DATA_TIME");
    if (lastDataTimeStr) {
        var lastDataTime = parseInt(lastDataTimeStr);
        if (nowTime - lastDataTime > 210000) {
            SCRIPT_PROPS.setProperty("LAST_RESET_TIME", nowTime.toString());
        }
    }
    SCRIPT_PROPS.setProperty("LAST_DATA_TIME", nowTime.toString());

    if (updatedRows >= 15) {
        var v = totalDataSheet.getRange("C2:E14").getValues();
        var temp_now = (Number(v[0][0]) + Number(v[1][0]) + Number(v[2][0])) / 3;
        var hum_now = (Number(v[0][1]) + Number(v[1][1]) + Number(v[2][1])) / 3;
        var pres_now = (Number(v[0][2]) + Number(v[1][2]) + Number(v[2][2])) / 3;
        var temp_old = (Number(v[10][0]) + Number(v[11][0]) + Number(v[12][0])) / 3;
        var pres_old = (Number(v[10][2]) + Number(v[11][2]) + Number(v[12][2])) / 3;

        var lastResetStr = SCRIPT_PROPS.getProperty("LAST_RESET_TIME");
        var lastAlarmStr = SCRIPT_PROPS.getProperty("LAST_ALARM_TIME");
        var isImmune = (lastResetStr && (nowTime - parseInt(lastResetStr) < 15 * 60 * 1000));
        var isCooled = (lastAlarmStr && (nowTime - parseInt(lastAlarmStr) < 30 * 60 * 1000));

        if (!isImmune && !isCooled) {
            if ((temp_old - temp_now >= 2.5) && (hum_now >= 90.0) && ((pres_old - pres_now >= 1.0) || (pres_now <= 1005.0))) {
                var alertMsg = "⛈️【陽台暴風雨特報】\n大腦偵測到氣象劇烈突變，大雨已來襲！\n\n" +
                    "🔹 溫度暴跌：" + (temp_old - temp_now).toFixed(1) + " °C\n" +
                    "🔹 空氣濕度：" + hum_now.toFixed(1) + " %\n" +
                    "🔹 大氣壓力：" + pres_now.toFixed(1) + " hPa";
                sendLinePushMessage(alertMsg);
                SCRIPT_PROPS.setProperty("LAST_ALARM_TIME", nowTime.toString());
            }
        }
    }

    var returnCmd = "OK";
    var resetFlag = dliSheet.getRange("J2").getValue();
    var motorFlag = dliSheet.getRange("I4").getValue();

    if (resetFlag === "RESET_PENDING" || resetFlag === "YES") {
        dliSheet.getRange("J2").setValue("NORMAL");
        SCRIPT_PROPS.setProperty("LAST_RESET_TIME", nowTime.toString());
        returnCmd = "CMD_RESET";
    }
    else if (motorFlag !== "" && motorFlag !== "NORMAL") {
        dliSheet.getRange("I4").setValue("NORMAL");
        returnCmd = motorFlag;
    }

    return ContentService.createTextOutput(returnCmd);
}

// ======================== 太陽仰角計算與可信度評分 ========================
function calculateSolarElevation_(date, lat, lon) {
    var rad = Math.PI / 180;
    var startOfYear = new Date(date.getFullYear(), 0, 1);
    var dayOfYear = Math.floor((date - startOfYear) / 86400000) + 1;
    var hour = date.getHours() + date.getMinutes() / 60 + date.getSeconds() / 3600;
    var gamma = (2 * Math.PI / 365) * (dayOfYear - 1 + (hour - 12) / 24);
    var eqtime = 229.18 * (0.000075 + 0.001868 * Math.cos(gamma) - 0.032077 * Math.sin(gamma)
        - 0.014615 * Math.cos(2 * gamma) - 0.040849 * Math.sin(2 * gamma));
    var decl = 0.006918 - 0.399912 * Math.cos(gamma) + 0.070257 * Math.sin(gamma)
        - 0.006758 * Math.cos(2 * gamma) + 0.000907 * Math.sin(2 * gamma)
        - 0.002697 * Math.cos(3 * gamma) + 0.00148 * Math.sin(3 * gamma);
    var timeOffset = eqtime + 4 * lon - 60 * TIMEZONE_OFFSET;
    var trueSolarTime = hour * 60 + timeOffset;
    var hourAngle = (trueSolarTime / 4) - 180;
    var latRad = lat * rad;
    var haRad = hourAngle * rad;
    var cosZenith = Math.sin(latRad) * Math.sin(decl) + Math.cos(latRad) * Math.cos(decl) * Math.cos(haRad);
    cosZenith = Math.max(-1, Math.min(1, cosZenith));
    var zenithRad = Math.acos(cosZenith);
    return 90 - (zenithRad / rad);
}

function calculateLightConfidence_(elevationDeg) {
    if (elevationDeg <= 0) return 0;
    var sinE = Math.sin(elevationDeg * Math.PI / 180);
    var sinRef = Math.sin(30 * Math.PI / 180);
    var confidence = Math.min(100, (sinE / sinRef) * 100);
    return Math.round(confidence * 10) / 10;
}

function calculateSolarPosition_(date, lat, lon) {
    var rad = Math.PI / 180;
    var startOfYear = new Date(date.getFullYear(), 0, 1);
    var dayOfYear = Math.floor((date - startOfYear) / 86400000) + 1;
    var hour = date.getHours() + date.getMinutes() / 60 + date.getSeconds() / 3600;
    var gamma = (2 * Math.PI / 365) * (dayOfYear - 1 + (hour - 12) / 24);
    var eqtime = 229.18 * (0.000075 + 0.001868 * Math.cos(gamma) - 0.032077 * Math.sin(gamma)
        - 0.014615 * Math.cos(2 * gamma) - 0.040849 * Math.sin(2 * gamma));
    var decl = 0.006918 - 0.399912 * Math.cos(gamma) + 0.070257 * Math.sin(gamma)
        - 0.006758 * Math.cos(2 * gamma) + 0.000907 * Math.sin(2 * gamma)
        - 0.002697 * Math.cos(3 * gamma) + 0.00148 * Math.sin(3 * gamma);
    var timeOffset = eqtime + 4 * lon - 60 * TIMEZONE_OFFSET;
    var trueSolarTime = hour * 60 + timeOffset;
    var hourAngle = (trueSolarTime / 4) - 180;
    var latRad = lat * rad;
    var haRad = hourAngle * rad;
    var cosZenith = Math.sin(latRad) * Math.sin(decl) + Math.cos(latRad) * Math.cos(decl) * Math.cos(haRad);
    cosZenith = Math.max(-1, Math.min(1, cosZenith));
    var zenithRad = Math.acos(cosZenith);
    var elevation = 90 - zenithRad / rad;
    var elevRad = elevation * rad;
    var cosAz = (Math.sin(decl) - Math.sin(elevRad) * Math.sin(latRad)) / (Math.cos(elevRad) * Math.cos(latRad));
    cosAz = Math.max(-1, Math.min(1, cosAz));
    var azRaw = Math.acos(cosAz) / rad;
    var azimuth = (hourAngle > 0) ? (360 - azRaw) : azRaw;
    return { elevation: elevation, azimuth: azimuth };
}

var BALCONY_AZ_MIN = 58;
var BALCONY_AZ_MAX = 233;
var BALCONY_EL_MIN = 20.4;
var BALCONY_EL_MAX = 84;

function isBalconyLit_(elevation, azimuth) {
    return (azimuth >= BALCONY_AZ_MIN && azimuth <= BALCONY_AZ_MAX) &&
           (elevation >= BALCONY_EL_MIN && elevation <= BALCONY_EL_MAX);
}

function estimateShadowTransitionTime_(dateStr, lat, lon) {
    var parts = dateStr.split('/');
    var baseDate = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 5, 0, 0);
    var prevLit = null;
    var transitionTime = '';
    var minDiffFromNoon = Infinity;

    for (var m = 0; m <= 14 * 60; m++) {
        var t = new Date(baseDate.getTime() + m * 60000);
        var pos = calculateSolarPosition_(t, lat, lon);
        var lit = isBalconyLit_(pos.elevation, pos.azimuth);
        if (prevLit === true && lit === false) {
            var diffFromNoon = Math.abs((t.getHours() * 60 + t.getMinutes()) - 12 * 60);
            if (diffFromNoon < minDiffFromNoon) {
                minDiffFromNoon = diffFromNoon;
                transitionTime = Utilities.formatDate(t, 'GMT+8', 'HH:mm');
            }
        }
        prevLit = lit;
    }
    return transitionTime;
}

// ======================== 🛡️ 核心 3 函式：資安過濾 ========================
function sanitizeInput(val, isNumber) {
    if (val === undefined || val === null) return isNumber ? 0 : "";
    var str = val.toString().trim();
    if (isNumber) { var num = Number(str); return isNaN(num) ? 0 : num; }
    if (str.indexOf('=') === 0 || str.indexOf('+') === 0 || str.indexOf('-') === 0 || str.indexOf('@') === 0) { return "'" + str; }
    return str;
}

// ======================== 🛡️ 核心 1 函式：雲端看門狗 ========================
function checkEspStatusAndInsertBlank() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName("資訊總表");
    if (!dataSheet) return;

    var nowTime = new Date().getTime();
    var lastDataTimeStr = SCRIPT_PROPS.getProperty("LAST_DATA_TIME");
    if (!lastDataTimeStr) return;

    var diff = nowTime - parseInt(lastDataTimeStr);
    if (diff > 120000 && diff < 86400000) {
        var now = new Date();
        var dateStr = Utilities.formatDate(now, "GMT+8", "yyyy/MM/dd");
        var timeStr = Utilities.formatDate(now, "GMT+8", "HH:mm:ss");
        dataSheet.insertRowBefore(2);
        dataSheet.getRange(2, 1, 1, 14).setValues([[dateStr, timeStr, "", "", "", "", "", "", "", "", "斷線", "", "", ""]]);
    }
}

// ======================== 🛡️ 核心 4 函式：每日 DLI 歷史紀錄結算與鎖定引擎 ========================
function archiveDailyDLI() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DLI統計");
    if (!sheet) {
        console.error("找不到 'DLI統計' 分頁！");
        return;
    }

    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+8", "yyyy/MM/dd");
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var dateRange = sheet.getRange(1, 1, lastRow, 1).getValues();
    var targetRow = -1;

    for (var i = 0; i < dateRange.length; i++) {
        var cellValue = dateRange[i][0];
        if (cellValue instanceof Date) {
            var cellDateStr = Utilities.formatDate(cellValue, "GMT+8", "yyyy/MM/dd");
            if (cellDateStr === todayStr) {
                targetRow = i + 1;
                break;
            }
        } else if (typeof cellValue === "string" && cellValue.replace(/-/g, "/") === todayStr) {
            targetRow = i + 1;
            break;
        }
    }

    if (targetRow === -1) {
        targetRow = lastRow + 1;
        sheet.getRange(targetRow, 1).setValue(todayStr);
        SpreadsheetApp.flush();
        console.log("成功生成新日期：" + todayStr);
    }

    var range = sheet.getRange(targetRow, 2, 1, 4);
    var liveValues = range.getValues();

    var numericValues = [[
        Number(liveValues[0][0]) || 0,
        Number(liveValues[0][1]) || 0,
        Number(liveValues[0][2]) || 0,
        Number(liveValues[0][3]) || 0
    ]];

    range.setValues(numericValues);
    console.log("已成功將 " + todayStr + " 的 B~E 欄鎖定為「純數字」歷史紀錄！");
}

// ======================== 🆕 每日統計表：系統直接寫入,不用公式 ========================

// 🔧 修正用：把各種可能出現的日期字串/日期物件，統一 normalize 成 yyyy/MM/dd（補零）格式，
// 避免「2026/7/16」與「2026/07/16」被判定為不同日期，導致明明已有列卻又新增一列的問題。
function normalizeDateStr_(v) {
    if (v instanceof Date) return Utilities.formatDate(v, "GMT+8", "yyyy/MM/dd");
    var s = (v || "").toString().trim().replace(/-/g, "/");
    var parts = s.split("/");
    if (parts.length !== 3) return s;
    var y = parts[0];
    var m = ("0" + parts[1]).slice(-2);
    var d = ("0" + parts[2]).slice(-2);
    return y + "/" + m + "/" + d;
}

function archiveDailyStatsTable() {
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+8", "yyyy/MM/dd");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var source = ss.getSheetByName("資訊總表");
    if (!source) { console.error("找不到資訊總表分頁！"); return; }
    var lastRow = source.getLastRow();
    if (lastRow < 2) return;

    var scanRows = Math.min(1500, lastRow - 1);
    var data = source.getRange(2, 1, scanRows, 14).getValues();

    var todayRows = [];
    for (var i = 0; i < data.length; i++) {
        var d = data[i][0];
        var dStr = (d instanceof Date) ? Utilities.formatDate(d, "GMT+8", "yyyy/MM/dd") : d;
        if (dStr === todayStr) todayRows.push(data[i]);
    }
    if (todayRows.length === 0) { console.log("找不到 " + todayStr + " 的資料，略過每日統計表"); return; }

    computeAndWriteDailyStats_(todayStr, todayRows);
}

function archiveDailyStatsTableForDate_(todayStr) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var source = ss.getSheetByName("資訊總表");
    if (!source) { console.error("找不到資訊總表分頁！"); return; }
    var lastRow = source.getLastRow();
    if (lastRow < 2) return;

    var allData = source.getRange(2, 1, lastRow - 1, 14).getValues();

    var todayRows = [];
    for (var i = 0; i < allData.length; i++) {
        var d = allData[i][0];
        var dStr = (d instanceof Date) ? Utilities.formatDate(d, "GMT+8", "yyyy/MM/dd") : d;
        if (dStr === todayStr) todayRows.push(allData[i]);
    }
    if (todayRows.length === 0) { console.log("找不到 " + todayStr + " 的資料，略過每日統計表"); return; }

    computeAndWriteDailyStats_(todayStr, todayRows);
}

function computeAndWriteDailyStats_(todayStr, todayRowsInput) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("每日統計表");
    if (!sheet) { console.error("找不到「每日統計表」分頁！"); return; }

    // 傳進來的 todayStr 也一併 normalize，確保跟表內比對用的是同一種格式
    todayStr = normalizeDateStr_(todayStr);

    var todayRows = todayRowsInput.slice().reverse();

    function findExtreme(colIndex, wantMax) {
        var best = null, bestTime = "";
        for (var i = 0; i < todayRows.length; i++) {
            var v = todayRows[i][colIndex];
            if (v === "" || v === null || isNaN(v)) continue;
            v = Number(v);
            if (best === null || (wantMax && v > best) || (!wantMax && v < best)) {
                best = v;
                var tv = todayRows[i][1];
                bestTime = (tv instanceof Date) ? Utilities.formatDate(tv, "GMT+8", "HH:mm:ss") : tv;
            }
        }
        return [best === null ? "" : best, bestTime];
    }

    var maxTemp  = findExtreme(2, true);
    var minTemp  = findExtreme(2, false);
    var maxHum   = findExtreme(3, true);
    var minHum   = findExtreme(3, false);
    var maxPres  = findExtreme(4, true);
    var minPres  = findExtreme(4, false);
    var maxLight = findExtreme(5, true);
    var maxPPFD  = findExtreme(6, true);

    var empiricalTransition = "";
    if (maxLight[0] !== "") {
        var peakValue = Number(maxLight[0]);
        var peakIndex = -1;
        for (var i = 0; i < todayRows.length; i++) {
            var v = todayRows[i][5];
            if (v !== "" && v !== null && !isNaN(v) && Number(v) === peakValue) { peakIndex = i; break; }
        }
        if (peakIndex !== -1) {
            var threshold = peakValue * 0.2;
            for (var j = peakIndex + 1; j < todayRows.length; j++) {
                var vv = todayRows[j][5];
                if (vv !== "" && vv !== null && !isNaN(vv) && Number(vv) < threshold) {
                    var tv2 = todayRows[j][1];
                    empiricalTransition = (tv2 instanceof Date) ? Utilities.formatDate(tv2, "GMT+8", "HH:mm") : tv2;
                    break;
                }
            }
        }
    }

    var estimatedTransition = estimateShadowTransitionTime_(todayStr, SENSOR_LAT, SENSOR_LON);

    var dliSum = 0;
    for (var i = 0; i < todayRows.length; i++) {
        var ppfd = todayRows[i][6];
        if (ppfd !== "" && ppfd !== null && !isNaN(ppfd)) {
            dliSum += Number(ppfd) * 60;
        }
    }
    var cumulativeDLI = Math.round((dliSum / 1000000) * 1000) / 1000;

    var normalCount = 0, disconnectCount = 0, sensorErrorCount = 0;
    for (var i = 0; i < todayRows.length; i++) {
        var st = todayRows[i][10];
        if (st === "正常") normalCount++;
        else if (st === "斷線") disconnectCount++;
        else if (st !== "" && st !== null) sensorErrorCount++;
    }
    var sampleRate      = Math.min(100, Math.round((normalCount / 1440) * 1000) / 10) + "%";
    var disconnectRate  = Math.min(100, Math.round((disconnectCount / 1440) * 1000) / 10) + "%";
    var sensorErrorRate = Math.min(100, Math.round((sensorErrorCount / 1440) * 1000) / 10) + "%";

    // 🔧 修正：比對前先把表內既有日期 normalize，避免補零/不補零格式不一致
    //    導致同一天被判定成不同日期，進而重複新增列而不是覆蓋既有列。
    var lastRowInSheet = sheet.getLastRow();
    var targetRow = -1;
    if (lastRowInSheet >= 2) {
        var dateCol = sheet.getRange(2, 1, lastRowInSheet - 1, 1).getValues();
        for (var i = 0; i < dateCol.length; i++) {
            var vStr = normalizeDateStr_(dateCol[i][0]);
            if (vStr === todayStr) { targetRow = i + 2; break; }
        }
    }
    if (targetRow === -1) {
        sheet.insertRowBefore(2);
        targetRow = 2;
    }

    sheet.getRange(targetRow, 1, 1, 17).setValues([[
        todayStr,
        maxTemp[0], maxTemp[1],
        minTemp[0], minTemp[1],
        maxHum[0], minHum[0],
        maxPres[0], minPres[0],
        maxLight[0],
        empiricalTransition,
        estimatedTransition,
        maxPPFD[0],
        cumulativeDLI,
        sampleRate, disconnectRate, sensorErrorRate
    ]]);

    console.log("已寫入 " + todayStr + " 每日統計表（實測轉換" + empiricalTransition + "／預估轉換" + estimatedTransition + "／正常" + sampleRate + "）");
}
// ======================== 正規 LINE 推播引擎 ========================
function sendLinePushMessage(textMessage) {
    var url = "https://api.line.me/v2/bot/message/push";
    var payload = { "to": LINE_USER_ID, "messages": [{ "type": "text", "text": textMessage }] };
    var options = {
        "method": "post", "contentType": "application/json",
        "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
        "payload": JSON.stringify(payload), "muteHttpExceptions": true
    };
    try {
        UrlFetchApp.fetch(url, options);
    } catch (e) {
        console.error("LINE 推播例外: " + e.message);
    }
}

// ======================== 🤖 LINE Bot 接收大門 ========================
function doPost(e) {
    if (typeof e === 'undefined') return ContentService.createTextOutput("OK");
    var msg = JSON.parse(e.postData.contents);
    var events = msg.events;

    for (var i = 0; i < events.length; i++) {
        var event = events[i];
        if (event.type === 'message' && event.message.type === 'text') {
            var userMessage = event.message.text.trim();
            var replyToken = event.replyToken;

            if (event.source.userId !== LINE_USER_ID) {
                replyText(replyToken, "⚠️ 您無權限操作此系統。");
                continue;
            }

            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var dliSheet = ss.getSheetByName("即時訊息");
            var dataSheet = ss.getSheetByName("資訊總表");
            var nowTime = new Date().getTime();

            if (userMessage === "看狀態") {
                var latestData = dataSheet.getRange(2, 1, 1, 11).getValues()[0];
                replyText(replyToken, "📊 【陽台即時狀態】\n🕒 時間：" + latestData[1] + "\n🌡️ 溫度：" + latestData[2] + " °C\n💧 濕度：" + latestData[3] + " %\n☀️ 光照：" + latestData[5] + " lx\n🌱 盆栽狀態：" + latestData[10]);
            }
            else if (userMessage === "看圖表") {
                var chart1 = buildQuickChartUrl("temp_moisture", dataSheet);
                var chart2 = buildQuickChartUrl("light_pressure", dataSheet);
                replyImage(replyToken, chart1, chart2);
            }
            else if (userMessage === "澆水功能") {
                replyButtonMenu(replyToken, "澆水", "請選擇澆水目標", "💦 請點擊下方按鈕，選擇要【澆水】的盆栽：");
            }
            else if (userMessage === "洗盆功能") {
                replyButtonMenu(replyToken, "洗盆", "請選擇洗盆目標", "🚿 請點擊下方按鈕，選擇要【洗盆】的盆栽：");
            }
            else if (userMessage.indexOf("澆水 ") === 0 || userMessage.indexOf("洗盆 ") === 0) {
                var action = userMessage.split(" ")[0];
                var plant = userMessage.split(" ")[1];
                var commandString = "";

                var code = PLANT_CODE_MAP[plant];
                if (code) {
                    commandString = (action === "澆水" ? "PUMP_" : "WASH_") + code;
                }

                if (commandString !== "") {
                    dliSheet.getRange("I4").setValue(commandString);
                    replyText(replyToken, "✅ 已下達「" + userMessage + "」指令！\nESP32 將於下次打卡時啟動對應繼電器。");
                } else {
                    replyText(replyToken, "❌ 找不到該盆栽，請確認名稱是否正確。");
                }
            }
            else {
                replyText(replyToken, "🤔 系統聽不懂。請點選圖文選單或輸入：看狀態、看圖表、澆水功能、洗盆功能。");
            }
        }
    }
    return ContentService.createTextOutput("OK");
}

// ======================== 📱 LINE 通訊與按鈕引擎區 ========================
function replyText(replyToken, text) {
    fetchLineApi({ "replyToken": replyToken, "messages": [{ "type": "text", "text": text }] });
}
function replyImage(replyToken, url1, url2) {
    fetchLineApi({
        "replyToken": replyToken,
        "messages": [
            { "type": "image", "originalContentUrl": url1, "previewImageUrl": url1 },
            { "type": "image", "originalContentUrl": url2, "previewImageUrl": url2 }
        ]
    });
}
function replyButtonMenu(replyToken, actionPrefix, altText, menuText) {
    fetchLineApi({
        "replyToken": replyToken,
        "messages": [{
            "type": "template",
            "altText": altText,
            "template": {
                "type": "buttons",
                "text": menuText,
                "actions": [
                    { "type": "message", "label": "🌸 茉莉", "text": actionPrefix + " 茉莉" },
                    { "type": "message", "label": "🌿 桂花", "text": actionPrefix + " 桂花" },
                    { "type": "message", "label": "🌶️ 朝天椒", "text": actionPrefix + " 朝天椒" }
                ]
            }
        }]
    });
}
function fetchLineApi(payload) {
    var response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
        "method": "post",
        "contentType": "application/json",
        "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
    });
    var code = response.getResponseCode();
    if (code !== 200) {
        console.error("LINE reply 失敗 (HTTP " + code + "): " + response.getContentText());
    }
}

// ======================== 📊 QuickChart 圖表生成 ========================
function buildQuickChartUrl(chartType, sheet) {
    var lastRow = sheet.getLastRow();
    var numRows = Math.min(15, lastRow - 1);
    if (numRows < 1) return "";
    var dataRange = sheet.getRange(2, 2, numRows, 9).getValues();
    dataRange.reverse();
    var labels = [], d_temp = [], d_soilJ = [], d_soilO = [], d_soilP = [], d_light = [], d_pres = [];

    for (var i = 0; i < dataRange.length; i++) {
        var row = dataRange[i];
        var timeStr = (row[0] instanceof Date) ? Utilities.formatDate(row[0], "GMT+8", "HH:mm") : row[0].toString().substring(0, 5);
        labels.push(timeStr);
        d_temp.push(row[1]); d_pres.push(row[3]); d_light.push(row[4]);
        d_soilJ.push(row[6]); d_soilO.push(row[7]); d_soilP.push(row[8]);
    }

    var chartConfig = {};
    if (chartType === "temp_moisture") {
        chartConfig = {
            type: 'line', data: {
                labels: labels, datasets: [
                    { label: '溫度', data: d_temp, borderColor: 'red', fill: false, yAxisID: 'y1' },
                    { label: '茉莉', data: d_soilJ, borderColor: 'blue', fill: false, yAxisID: 'y2' },
                    { label: '桂花', data: d_soilO, borderColor: 'green', fill: false, yAxisID: 'y2' },
                    { label: '朝天椒', data: d_soilP, borderColor: 'orange', fill: false, yAxisID: 'y2' }
                ]
            }, options: {
                title: { display: true, text: '🌡️ 最近溫度與盆栽濕度 (15筆)' },
                scales: { yAxes: [{ id: 'y1', position: 'left', scaleLabel: { display: true, labelString: '溫度(°C)' } }, { id: 'y2', position: 'right', scaleLabel: { display: true, labelString: '土壤濕度' } }] }
            }
        };
    } else {
        chartConfig = {
            type: 'line', data: {
                labels: labels, datasets: [
                    { label: '光照', data: d_light, borderColor: 'orange', fill: false, yAxisID: 'y1' },
                    { label: '氣壓', data: d_pres, borderColor: 'purple', fill: false, yAxisID: 'y2' }
                ]
            }, options: {
                title: { display: true, text: '☀️ 最近光照與大氣壓力 (15筆)' },
                scales: { yAxes: [{ id: 'y1', position: 'left', scaleLabel: { display: true, labelString: '光照(lx)' } }, { id: 'y2', position: 'right', scaleLabel: { display: true, labelString: '氣壓(hPa)' } }] }
            }
        };
    }
    return getQuickChartShortUrl(chartConfig);
}

function getQuickChartShortUrl(chartConfig) {
    var payload = {
        "chart": chartConfig,
        "backgroundColor": "white",
        "width": 500,
        "height": 300,
        "format": "png"
    };
    var options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
    };
    try {
        var response = UrlFetchApp.fetch("https://quickchart.io/chart/create", options);
        var result = JSON.parse(response.getContentText());
        if (result.success && result.url) {
            return result.url;
        } else {
            console.error("QuickChart 短網址建立失敗，改用原始長網址備援: " + response.getContentText());
            return "https://quickchart.io/chart?bkg=white&c=" + encodeURIComponent(JSON.stringify(chartConfig));
        }
    } catch (e) {
        console.error("QuickChart 請求例外，改用原始長網址備援: " + e.message);
        return "https://quickchart.io/chart?bkg=white&c=" + encodeURIComponent(JSON.stringify(chartConfig));
    }
}

// ======================== 🗄️ 封存機制：每小時統計表 ========================
const ARCHIVE_CONFIG = {
  SOURCE_SHEET_NAME: '資訊總表',
  ARCHIVE_SHEET_NAME: '每小時統計表',
  DATE_COL: 1,
  TIME_COL: 2
};

function parseRowDateTime(dateVal, timeVal) {
  if (dateVal instanceof Date && timeVal instanceof Date) {
    return new Date(
      dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate(),
      timeVal.getHours(), timeVal.getMinutes(), timeVal.getSeconds()
    );
  }
  if (dateVal instanceof Date && typeof timeVal === 'string') {
    var t = timeVal.split(':');
    return new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate(),
                     Number(t[0]), Number(t[1]), Number(t[2] || 0));
  }
  var d = dateVal.toString().split('/');
  var t = timeVal.toString().split(':');
  return new Date(Number(d[0]), Number(d[1]) - 1, Number(d[2]),
                   Number(t[0]), Number(t[1]), Number(t[2] || 0));
}

function isSameHourBucket(a, b) {
  return a.getFullYear() === b.getFullYear() &&
         a.getMonth() === b.getMonth() &&
         a.getDate() === b.getDate() &&
         a.getHours() === b.getHours();
}

function getOrCreateArchiveSheet(ss) {
  var archive = ss.getSheetByName(ARCHIVE_CONFIG.ARCHIVE_SHEET_NAME);
  if (!archive) {
    archive = ss.insertSheet(ARCHIVE_CONFIG.ARCHIVE_SHEET_NAME);
    archive.getRange(1, 1, 1, 16).setValues([[
      '日期', '時間', '溫度均', '濕度均', '氣壓均', '光照均', 'PPFD均',
      '土壤J均', '土壤O均', '土壤P均', '累積J', '累積O', '累積P',
      '取樣率(正常)', '斷線比率', '感測器異常率'
    ]]);
  }
  return archive;
}

function archiveHourlyAverage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName(ARCHIVE_CONFIG.SOURCE_SHEET_NAME);
  var archive = getOrCreateArchiveSheet(ss);

  var now = new Date();
  var targetBucketTime = new Date(now.getTime() - 60 * 60 * 1000);

  var archiveLastRow = archive.getLastRow();
  if (archiveLastRow >= 2) {
    var lastArchived = archive.getRange(2, 1, 1, 2).getValues()[0];
    var lastArchivedTime = parseRowDateTime(lastArchived[0], lastArchived[1]);
    if (isSameHourBucket(lastArchivedTime, targetBucketTime)) {
      console.log('該時段已封存過，跳過：' + targetBucketTime);
      return;
    }
  }
  var lastRow = source.getLastRow();
  if (lastRow < 2) return;

  var scanRows = Math.min(150, lastRow - 1);
  var allData = source.getRange(2, 1, scanRows, source.getLastColumn()).getValues();

  var matchedRows = [];
  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    var dateVal = row[ARCHIVE_CONFIG.DATE_COL - 1];
    var timeVal = row[ARCHIVE_CONFIG.TIME_COL - 1];
    if (!dateVal || !timeVal) continue;

    var rowTime = parseRowDateTime(dateVal, timeVal);
    if (isNaN(rowTime.getTime())) continue;

    if (isSameHourBucket(rowTime, targetBucketTime)) {
      matchedRows.push(row);
    }
  }

  if (matchedRows.length === 0) {
    console.log('目標時段沒有資料，略過本次封存：' + targetBucketTime);
    return;
  }

  matchedRows.reverse();

  function avgCol(colIndex) {
    var sum = 0, count = 0;
    for (var i = 0; i < matchedRows.length; i++) {
      var v = matchedRows[i][colIndex];
      if (v !== '' && v !== null && !isNaN(v)) { sum += Number(v); count++; }
    }
    return count > 0 ? Math.round((sum / count) * 10) / 10 : '';
  }

  var lastRowOfHour = matchedRows[matchedRows.length - 1];

  var avgTemp   = avgCol(2);
  var avgHum    = avgCol(3);
  var avgPres   = avgCol(4);
  var avgLight  = avgCol(5);
  var avgPPFD   = avgCol(6);
  var avgSoilJ  = avgCol(7);
  var avgSoilO  = avgCol(8);
  var avgSoilP  = avgCol(9);
  var lastCumJ  = lastRowOfHour[11];
  var lastCumO  = lastRowOfHour[12];
  var lastCumP  = lastRowOfHour[13];

  var normalCount = 0, disconnectCount = 0, sensorErrorCount = 0;
  for (var k = 0; k < matchedRows.length; k++) {
    var st = matchedRows[k][10];
    if (st === '正常') { normalCount++; }
    else if (st === '斷線') { disconnectCount++; }
    else if (st !== '' && st !== null) { sensorErrorCount++; }
  }
  var sampleRate      = Math.min(100, Math.round((normalCount / 60) * 1000) / 10) + '%';
  var disconnectRate  = Math.min(100, Math.round((disconnectCount / 60) * 1000) / 10) + '%';
  var sensorErrorRate = Math.min(100, Math.round((sensorErrorCount / 60) * 1000) / 10) + '%';

  var bucketDateStr = Utilities.formatDate(targetBucketTime, 'GMT+8', 'yyyy/MM/dd');
  var bucketHourStr = Utilities.formatDate(targetBucketTime, 'GMT+8', 'HH') + ':00:00';

  archive.insertRowBefore(2);
  archive.getRange(2, 1, 1, 16).setValues([[
    bucketDateStr, bucketHourStr,
    avgTemp, avgHum, avgPres, avgLight, avgPPFD,
    avgSoilJ, avgSoilO, avgSoilP,
    lastCumJ, lastCumO, lastCumP,
    sampleRate, disconnectRate, sensorErrorRate
  ]]);
  archive.getRange('B:B').setNumberFormat('HH:mm');
  console.log('已封存 ' + bucketDateStr + ' ' + bucketHourStr + ' 均值，正常：' + sampleRate + '，斷線：' + disconnectRate + '，感測器異常：' + sensorErrorRate);
}

function archiveHourlyIfNeeded() {
  archiveHourlyAverage();
}


// ============================================================
// 🌤️ CWA 天氣預報快照模組
// ============================================================
const CWA_LOCATION = '恆春鎮';
const CWA_SNAPSHOT_START_COL = 1;

function fetchCWA_(dataId, elementCodes) {
  var apiKey = SCRIPT_PROPS.getProperty('CWA_API_KEY');
  if (!apiKey) {
    throw new Error('尚未設定 CWA_API_KEY,請至指令碼屬性新增');
  }

  var url = 'https://opendata.cwa.gov.tw/api/v1/rest/datastore/' + dataId
    + '?Authorization=' + apiKey
    + '&locationName=' + encodeURIComponent(CWA_LOCATION)
    + '&elementName=' + elementCodes.join(',')
    + '&format=JSON';

  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('CWA API 呼叫失敗 (HTTP ' + code + '): ' + response.getContentText());
  }

  var json = JSON.parse(response.getContentText());
  if (json.success !== 'true' && json.success !== true) {
    throw new Error('CWA API 回傳失敗: ' + JSON.stringify(json));
  }
  return json;
}

function parseISO_(iso) {
  if (!iso) return null;
  var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})([+-])(\d{2}):(\d{2})$/);
  if (!m) {
    Logger.log('警告:無法解析時間字串 → ' + iso);
    return new Date(iso);
  }
  var y = +m[1], mo = +m[2], d = +m[3], h = +m[4], mi = +m[5], s = +m[6];
  var sign = m[7], oh = +m[8], om = +m[9];
  var utcMs = Date.UTC(y, mo - 1, d, h, mi, s);
  var offsetMs = (sign === '+' ? -1 : 1) * (oh * 60 + om) * 60000;
  return new Date(utcMs + offsetMs);
}

function getWeatherElement_(location, name) {
  var el = location.WeatherElement.find(function (e) {
    return e.ElementName === name;
  });
  return el ? el.Time : [];
}

function toNum_(v) {
  if (v === undefined || v === null || v === '-') return '';
  var n = parseFloat(v);
  return isNaN(n) ? v : n;
}

function pickCurrentRow_(rows) {
  var now = new Date();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][1] && rows[i][1].getTime() > now.getTime()) {
      return rows[i];
    }
  }
  return rows.length > 0 ? rows[rows.length - 1] : null;
}

function parseCWA3Day_(json) {
  var loc = json.records.Locations[0].Location[0];

  var tMap = {};
  getWeatherElement_(loc, '溫度').forEach(function (t) {
    tMap[String(t.DataTime).trim()] = t.ElementValue[0].Temperature;
  });

  var rhMap = {};
  getWeatherElement_(loc, '相對濕度').forEach(function (t) {
    rhMap[String(t.DataTime).trim()] = t.ElementValue[0].RelativeHumidity;
  });

  var dirMap = {};
  getWeatherElement_(loc, '風向').forEach(function (t) {
    dirMap[String(t.DataTime).trim()] = t.ElementValue[0].WindDirection;
  });

  var popMap = {};
  getWeatherElement_(loc, '3小時降雨機率').forEach(function (t) {
    popMap[String(t.StartTime).trim()] = t.ElementValue[0].ProbabilityOfPrecipitation;
  });

  var wxMap = {};
  getWeatherElement_(loc, '天氣現象').forEach(function (t) {
    wxMap[String(t.StartTime).trim()] = t.ElementValue[0];
  });

  var windSpeedArr = getWeatherElement_(loc, '風速');

  return windSpeedArr.map(function (w) {
    var key = String(w.DataTime).trim();
    var startDate = parseISO_(key);
    var endDate = startDate ? new Date(startDate.getTime() + 3 * 60 * 60 * 1000) : null;

    var dir = dirMap[key] || '';
    var pop = popMap[key];
    var wx = wxMap[key] || {};

    return [
      startDate, endDate,
      toNum_(tMap[key]), toNum_(rhMap[key]),
      toNum_(w.ElementValue[0].WindSpeed), toNum_(w.ElementValue[0].BeaufortScale),
      dir, toNum_(pop),
      wx.Weather || ''
    ];
  });
}

function parseCWA1Week_(json) {
  var loc = json.records.Locations[0].Location[0];

  var maxTMap = {};
  getWeatherElement_(loc, '最高溫度').forEach(function (t) {
    maxTMap[String(t.StartTime).trim()] = t.ElementValue[0].MaxTemperature;
  });
  var minTMap = {};
  getWeatherElement_(loc, '最低溫度').forEach(function (t) {
    minTMap[String(t.StartTime).trim()] = t.ElementValue[0].MinTemperature;
  });
  var rhMap = {};
  getWeatherElement_(loc, '平均相對濕度').forEach(function (t) {
    rhMap[String(t.StartTime).trim()] = t.ElementValue[0].RelativeHumidity;
  });
  var windMap = {};
  getWeatherElement_(loc, '風速').forEach(function (t) {
    windMap[String(t.StartTime).trim()] = t.ElementValue[0];
  });
  var dirMap = {};
  getWeatherElement_(loc, '風向').forEach(function (t) {
    dirMap[String(t.StartTime).trim()] = t.ElementValue[0].WindDirection;
  });
  var popMap = {};
  getWeatherElement_(loc, '12小時降雨機率').forEach(function (t) {
    popMap[String(t.StartTime).trim()] = t.ElementValue[0].ProbabilityOfPrecipitation;
  });
  var wxMap = {};
  getWeatherElement_(loc, '天氣現象').forEach(function (t) {
    wxMap[String(t.StartTime).trim()] = t.ElementValue[0];
  });
  var uviMap = {};
  getWeatherElement_(loc, '紫外線指數').forEach(function (t) {
    uviMap[String(t.StartTime).trim()] = t.ElementValue[0];
  });

  var avgTArr = getWeatherElement_(loc, '平均溫度');

  return avgTArr.map(function (t) {
    var key = String(t.StartTime).trim();
    var startDate = parseISO_(key);
    var endDate = parseISO_(String(t.EndTime).trim());

    var wind = windMap[key] || {};
    var wx = wxMap[key] || {};
    var uvi = uviMap[key] || {};

    return [
      startDate, endDate,
      toNum_(t.ElementValue[0].Temperature),
      toNum_(maxTMap[key]), toNum_(minTMap[key]), toNum_(rhMap[key]),
      toNum_(wind.WindSpeed), toNum_(wind.BeaufortScale),
      dirMap[key] || '', toNum_(popMap[key]),
      wx.Weather || '',
      toNum_(uvi.UVIndex), uvi.UVExposureLevel || ''
    ];
  });
}

function getCurrentUVI_() {
  var json = fetchCWA_('F-D0047-035', ['UVI']);
  var loc = json.records.Locations[0].Location[0];
  var uviTimes = getWeatherElement_(loc, '紫外線指數');
  var now = new Date();
  for (var i = 0; i < uviTimes.length; i++) {
    var end = parseISO_(String(uviTimes[i].EndTime).trim());
    if (end && end.getTime() > now.getTime()) {
      return uviTimes[i].ElementValue[0];
    }
  }
  return {};
}

function writeSnapshotRow_(sheet, headerRow, dataRow, headers, values) {
  sheet.getRange(headerRow, CWA_SNAPSHOT_START_COL, 1, headers.length).setValues([headers]);
  sheet.getRange(dataRow, CWA_SNAPSHOT_START_COL, 1, values.length).setValues([values]);
}

function updateWeatherSnapshot_3Day() {
  var json = fetchCWA_('F-D0047-033', ['T', 'RH', 'Wind', 'PoP', 'Wx']);
  var rows = parseCWA3Day_(json);
  var current = pickCurrentRow_(rows);
  if (!current) {
    Logger.log('3小時預報:找不到目前時段資料');
    return;
  }

  var uvi = getCurrentUVI_();

  var reordered = [
    current[0], current[1], current[2],
    toNum_(uvi.UVIndex), uvi.UVExposureLevel || '',
    current[3], current[4], current[5], current[6], current[7], current[8]
  ];

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('即時訊息');
  writeSnapshotRow_(sheet, 9, 10,
    ['時段起', '時段迄', '溫度(°C)', '紫外線指數', 'UV等級', '相對濕度(%)', '風速(m/s)', '蒲福風級', '風向', '降雨機率(%)', '天氣現象'],
    reordered
  );
  Logger.log('3小時預報快照更新完成:' + current[0]);
}

function updateWeatherSnapshot_1Week() {
  var json = fetchCWA_('F-D0047-035', ['T', 'RH', 'Wind', 'PoP', 'Wx']);
  var rows = parseCWA1Week_(json);
  var current = pickCurrentRow_(rows);
  if (!current) {
    Logger.log('1週預報:找不到目前時段資料');
    return;
  }

  var trimmed = current.slice(0, 11);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('即時訊息');
  writeSnapshotRow_(sheet, 11, 12,
    ['時段起', '時段迄', '平均溫度(°C)', '最高溫度(°C)', '最低溫度(°C)', '平均相對濕度(%)',
     '風速(m/s)', '蒲福風級', '風向', '降雨機率(%)', '天氣現象'],
    trimmed
  );
  Logger.log('1週預報快照更新完成:' + current[0]);
}


// ============================================================
// 🧰 手動 / 一次性工具函式（不會被觸發器自動呼叫，需要時才手動執行一次）
// ============================================================

// 【觸發ESP32遠端重置】
// 設定方式：不用設定，隨時要重置ESP32時，在此編輯器函式選單選 triggerESP32Reset，點「執行」即可。
// 執行後會把「即時訊息」分頁 J2 設為 RESET_PENDING，等 ESP32 下次打卡(doGet)時會被清成 NORMAL、
// 並回傳 CMD_RESET 指令給硬體，讓 ESP32 收到後自行重啟。
function triggerESP32Reset() {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("即時訊息").getRange("J2").setValue("RESET_PENDING");
    SpreadsheetApp.getUi().alert("🚀【指令已下達】\n遠端重置訊號已就緒！");
}

// 【建立/重建 CWA 天氣預報快照觸發器】
// 設定方式：只有在「觸發器」頁面被手動刪除、或第一次要啟用天氣快照功能時，才需要手動執行一次
// setupWeatherTriggers。執行前記得先在「專案設定→指令碼屬性」設好 CWA_API_KEY，否則抓取會失敗。
// 執行後會自動建立：3小時預報每3小時觸發一次(00/03/06.../21)、1週預報每12小時觸發一次(06/18)，
// 執行前會先清掉舊的同名觸發器，避免重複建立、跑好幾份。
function setupWeatherTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var fn = trigger.getHandlerFunction();
    if (fn === 'updateWeatherSnapshot_3Day' || fn === 'updateWeatherSnapshot_1Week') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  var startHour3Day = 0;
  var interval3Day = 3;
  for (var h1 = startHour3Day; h1 < 24; h1 += interval3Day) {
    ScriptApp.newTrigger('updateWeatherSnapshot_3Day')
      .timeBased().atHour(h1).nearMinute(5).everyDays(1).create();
  }

  var startHour1Week = 6;
  var interval1Week = 12;
  for (var h2 = startHour1Week; h2 < startHour1Week + 24; h2 += interval1Week) {
    ScriptApp.newTrigger('updateWeatherSnapshot_1Week')
      .timeBased().atHour(h2 % 24).nearMinute(5).everyDays(1).create();
  }

  Logger.log('氣象預報觸發器建立完成:3天預報對齊00/03/06.../21,1週預報對齊06/18(各延遲5分鐘緩衝)');
}

// 【回填每日統計表全部歷史】
// 設定方式：不用改任何參數。當「資訊總表」裡有任何日期在「每日統計表」還沒有對應列時
// (例如新裝了感測器補進舊資料、或每日統計表分頁重建過)，執行一次 backfillDailyStatsTableHistory，
// 會自動掃描資訊總表現存的「所有」日期，逐日呼叫 archiveDailyStatsTableForDate_ 補算，
// 不用像 archiveDailyStatsTableForDate_ 那樣要自己一天一天輸入日期字串下參數。
// 執行完可以在「執行記錄」看到掃到幾個日期、逐一補算完成的訊息。
function backfillDailyStatsTableHistory() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var source = ss.getSheetByName("資訊總表");
    if (!source) { console.error("找不到資訊總表分頁！"); return; }
    var lastRow = source.getLastRow();
    if (lastRow < 2) { console.log("資訊總表沒有資料"); return; }

    var allData = source.getRange(2, 1, lastRow - 1, 14).getValues();

    var rowsByDate = {};
    for (var i = 0; i < allData.length; i++) {
        var d = allData[i][0];
        var dStr = (d instanceof Date) ? Utilities.formatDate(d, "GMT+8", "yyyy/MM/dd") : d;
        if (!dStr) continue;
        if (!rowsByDate[dStr]) rowsByDate[dStr] = [];
        rowsByDate[dStr].push(allData[i]);
    }

    var dateList = Object.keys(rowsByDate);
    console.log("找到 " + dateList.length + " 個不同日期，開始逐日補算：" + dateList.join(", "));

    for (var j = 0; j < dateList.length; j++) {
        var dateStr = dateList[j];
        computeAndWriteDailyStats_(dateStr, rowsByDate[dateStr]);
    }
    console.log("歷史資料補算完成，共 " + dateList.length + " 天");
}
// ======================== 🤖 AI 匯出模式：跨分頁掃描 / 公式+值輸出 ========================

function handleAIExport_(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = e.parameter.sheet;
    var rangeStr = e.parameter.range;
    var scope = e.parameter.scope;

    var result;

    if (scope === "all") {
        // 一次掃描所有分頁，公式+值，大表自動限制列數避免逾時/資料爆量
        var sheets = ss.getSheets();
        var allData = {};
        var MAX_ROWS_PER_SHEET = 30; // 大表只抓最新30列，小表本來就不到30列會整份給

        for (var i = 0; i < sheets.length; i++) {
            var sh = sheets[i];
            var lastRow = sh.getLastRow();
            var lastCol = sh.getLastColumn();
            if (lastRow < 1 || lastCol < 1) {
                allData[sh.getName()] = { note: "空分頁" };
                continue;
            }
            var numRows = Math.min(MAX_ROWS_PER_SHEET, lastRow);
            var range = sh.getRange(1, 1, numRows, lastCol);
            allData[sh.getName()] = {
                totalRows: lastRow,
                totalCols: lastCol,
                rowsReturned: numRows,
                note: lastRow > MAX_ROWS_PER_SHEET ? "僅回傳前 " + MAX_ROWS_PER_SHEET + " 列，完整共 " + lastRow + " 列" : "",
                values: range.getValues(),
                formulas: range.getFormulas()
            };
        }
        result = { spreadsheetName: ss.getName(), sheets: allData };

    } else if (!sheetName) {
        // 原本的分頁分布總覽邏輯 (維持不變)
        var sheets = ss.getSheets();
        var sheetsInfo = [];
        for (var i = 0; i < sheets.length; i++) {
            var sh = sheets[i];
            sheetsInfo.push({
                name: sh.getName(), index: i,
                lastRow: sh.getLastRow(), lastColumn: sh.getLastColumn(),
                hidden: sh.isSheetHidden()
            });
        }
        result = { spreadsheetName: ss.getName(), sheetCount: sheets.length, sheets: sheetsInfo };

    } else {
        // 原本的單一分頁邏輯 (維持不變)
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            result = { error: "找不到分頁：" + sheetName };
        } else {
            var targetRange = rangeStr ? sheet.getRange(rangeStr) : sheet.getDataRange();
            result = {
                sheetName: sheetName, range: rangeStr || "(全部資料範圍)",
                numRows: targetRange.getNumRows(), numCols: targetRange.getNumColumns(),
                values: targetRange.getValues(), formulas: targetRange.getFormulas()
            };
        }
    }

    return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}
// ============================================================
// 🌤️ 仰角修正曲線統計模組（2026/07新增）
// ============================================================
// 目的：累積「資訊總表」裡裸版(T欄)/罩版(P欄)比值，依仰角(Q欄)分桶，
//       只採用可信度(R欄)≥90的樣本，持續收斂出穩定的仰角修正曲線。
// 設計：累積式更新，只處理「上次統計過的最後一列」之後的新資料，
//       不用每次全表重算，避免「資訊總表」13000+列造成逾時。

const ELEV_CURVE_SHEET_NAME = '仰角修正曲線';
const ELEV_CURVE_BIN_SIZE = 2; // 每桶涵蓋的仰角度數
const ELEV_CURVE_MIN_CONFIDENCE = 90;

function getOrCreateElevCurveSheet_(ss) {
  var sheet = ss.getSheetByName(ELEV_CURVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ELEV_CURVE_SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([[
      '仰角區間下限', '仰角區間上限', '樣本數', '比值總和', '平均比值', '標準差'
    ]]);
  }
  return sheet;
}

// 讀取既有桶次統計（樣本數、比值總和、平方和），回傳 map: binIndex -> {count, sum, sumSq}
function loadExistingBins_(sheet) {
  var lastRow = sheet.getLastRow();
  var bins = {};
  if (lastRow < 2) return bins;

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    var binLow = data[i][0];
    var count = Number(data[i][2]) || 0;
    var sum = Number(data[i][3]) || 0;
    // 反推平方和：用既有標準差還原(近似值即可，因為每次都是增量疊加，累積誤差可忽略)
    var std = Number(data[i][5]) || 0;
    var mean = count > 0 ? sum / count : 0;
    var sumSq = count > 0 ? count * (std * std + mean * mean) : 0;
    var binIndex = Math.floor(binLow / ELEV_CURVE_BIN_SIZE);
    bins[binIndex] = { count: count, sum: sum, sumSq: sumSq };
  }
  return bins;
}

function updateElevationCorrectionCurve() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName('資訊總表');
  if (!source) { console.error('找不到資訊總表分頁！'); return; }

  var curveSheet = getOrCreateElevCurveSheet_(ss);

  // 記錄上次統計到哪一列(用Script Properties存，避免每次全表重算)
  var lastProcessedRowStr = SCRIPT_PROPS.getProperty('ELEV_CURVE_LAST_ROW');
  var lastProcessedRow = lastProcessedRowStr ? Number(lastProcessedRowStr) : 0;

  var lastRow = source.getLastRow();
  if (lastRow < 2) return;

  // 資訊總表是「新資料在最上方」(insertRowBefore(2))，所以「尚未處理的新資料」
  // 是從第2列開始、往下數到「上次處理時的總列數」對應的那一列為止。
  // 用總列數差值判斷這次要往下掃多少列新資料。
  var totalRows = lastRow - 1; // 扣掉表頭
  var newRowCount = lastProcessedRow === 0 ? totalRows : (totalRows - lastProcessedRow);
  if (newRowCount <= 0) {
    console.log('沒有新資料需要統計，略過本次執行');
    return;
  }
  // 安全上限，避免單次掃描過量資料逾時(6分鐘執行限制)
  var scanRows = Math.min(newRowCount, 5000);

  // P欄(16)=罩版原始值, Q欄(17)=仰角, R欄(18)=可信度, T欄(20)=裸版原始值
  var data = source.getRange(2, 16, scanRows, 5).getValues();

  var bins = loadExistingBins_(curveSheet);

  for (var i = 0; i < data.length; i++) {
    var covered = data[i][0];      // P欄 罩版原始值
    var elevation = data[i][1];    // Q欄 仰角
    var confidence = data[i][2];   // R欄 可信度
    var bare = data[i][4];         // T欄 裸版原始值

    if (covered === '' || covered === null || isNaN(covered) || Number(covered) <= 0) continue;
    if (bare === '' || bare === null || isNaN(bare)) continue;
    if (elevation === '' || elevation === null || isNaN(elevation)) continue;
    if (confidence === '' || confidence === null || isNaN(confidence)) continue;
    if (Number(confidence) < ELEV_CURVE_MIN_CONFIDENCE) continue;

    var ratio = Number(bare) / Number(covered);
    var binIndex = Math.floor(Number(elevation) / ELEV_CURVE_BIN_SIZE);

    if (!bins[binIndex]) bins[binIndex] = { count: 0, sum: 0, sumSq: 0 };
    bins[binIndex].count += 1;
    bins[binIndex].sum += ratio;
    bins[binIndex].sumSq += ratio * ratio;
  }

  // 寫回：依binIndex排序後整批覆蓋寫入(桶數不多，全部重寫比逐筆定位更新更簡單可靠)
  var binIndexList = Object.keys(bins).map(Number).sort(function(a, b) { return a - b; });
  var outputRows = binIndexList.map(function(idx) {
    var b = bins[idx];
    var mean = b.sum / b.count;
    var variance = (b.sumSq / b.count) - (mean * mean);
    var std = variance > 0 ? Math.sqrt(variance) : 0;
    return [
      idx * ELEV_CURVE_BIN_SIZE,
      (idx + 1) * ELEV_CURVE_BIN_SIZE,
      b.count,
      Math.round(b.sum * 10000) / 10000,
      Math.round(mean * 10000) / 10000,
      Math.round(std * 10000) / 10000
    ];
  });

  if (outputRows.length > 0) {
    // 先清掉舊資料列(保留表頭)，再整批寫入，避免桶次增減造成殘留列
    var oldLastRow = curveSheet.getLastRow();
    if (oldLastRow > 1) {
      curveSheet.getRange(2, 1, oldLastRow - 1, 6).clearContent();
    }
    curveSheet.getRange(2, 1, outputRows.length, 6).setValues(outputRows);
  }

  SCRIPT_PROPS.setProperty('ELEV_CURVE_LAST_ROW', totalRows.toString());
  console.log('仰角修正曲線更新完成，本次新增處理 ' + scanRows + ' 列，累積 ' + outputRows.length + ' 個仰角區間');
}

// 【手動工具】重置累積統計，重新從頭掃描全部歷史資料
// 使用時機：改了ELEV_CURVE_BIN_SIZE或ELEV_CURVE_MIN_CONFIDENCE等參數後，
// 舊桶次統計已經不適用，需要執行這個清空重來。
function resetElevationCorrectionCurve() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var curveSheet = ss.getSheetByName(ELEV_CURVE_SHEET_NAME);
  if (curveSheet) {
    var lastRow = curveSheet.getLastRow();
    if (lastRow > 1) curveSheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }
  SCRIPT_PROPS.deleteProperty('ELEV_CURVE_LAST_ROW');
  console.log('仰角修正曲線統計已重置，下次執行updateElevationCorrectionCurve將從頭掃描全部歷史資料');
}