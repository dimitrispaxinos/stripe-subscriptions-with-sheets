const YES = 'Yes';
const NO = 'No';

class Settings {
    constructor(sheetName = "Settings") {
        this.sheetName = sheetName;
        this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        this.settingsSheet = this.spreadsheet.getSheetByName(sheetName);
        this.settingsMap = this.initSettingsMap();
    }


    init() {
        if (!this.settingsSheet) {
            this.settingsSheet = this.spreadsheet.insertSheet(this.sheetName);
            this.settingsSheet.appendRow(['Setting', 'Value']);
            this.settingsSheet.getRange('1:1').setFontWeight('bold');
        }
    }

    initSettingsMap() {
        if (!this.settingsSheet) return new Map();

        const data = this.settingsSheet.getDataRange().getValues();
        const map = new Map();

        for (const [key, value] of data) {
            map.set(key, value);
        }

        return map;
    }

    setSetting(settingName, settingValue) {
        const rowIndex = [...this.settingsMap.keys()].indexOf(settingName) + 1;

        if (rowIndex > 0) {
            this.settingsSheet.getRange(rowIndex, 2).setValue(settingValue);
        } else {
            this.settingsSheet.appendRow([settingName, settingValue]);
        }

        this.settingsMap.set(settingName, settingValue);
    }

    getSetting(settingName) {
        return this.settingsMap.get(settingName) || null;
    }

    getBooleanSetting(settingName) {
        const settingValue = this.getSetting(settingName);

        if (settingValue === YES) return true;
        if (settingValue === NO) return false;

        Logger.log(`Setting value is not ${YES} or ${NO}: ${settingName}`);
        return null;
    }

    setSettingInScriptProperties(settingName, settingValue) {
        PropertiesService.getScriptProperties().setProperty(settingName, settingValue);
    }

    getSettingFromScriptProperties(settingName) {
        return PropertiesService.getScriptProperties().getProperty(settingName);
    }
}
