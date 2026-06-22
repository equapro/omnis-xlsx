const omnis_calls = require('omnis_calls');
const XLSX = require("./vendor/sheetjs/xlsx");

let autoSendResponse = true; // Set to false in methods which should not send a response to Omnis when they exit. (e.g. async methods)

const PRECISION_CORRECTION = (function() {
    function getTimezoneOffsetMS(date) {
        const time = date.getTime();
        const utcTime = Date.UTC(
            date.getFullYear(),
            date.getMonth(),
            date.getDate(),
            date.getHours(),
            date.getMinutes(),
            date.getSeconds(),
            date.getMilliseconds()
        );
        return time - utcTime;
    }

    const basedate = new Date(1899, 11, 30, 0, 0, 0);
    const dnthreshAsIs = (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
    const dnthreshToBe = getTimezoneOffsetMS(new Date()) - getTimezoneOffsetMS(basedate);
    return dnthreshAsIs - dnthreshToBe;
})();

function fixSheetJSDate(date) {
    const timezoneOffset = date.getTimezoneOffset() * 60 * 1000;
    return new Date(date.getTime() + timezoneOffset + PRECISION_CORRECTION);
}

const methodMap = {
    /* =================================
     *  Writing Workbooks
     * ================================= */
    write: function (param) {
        // parameters
        const filename = param.filename;
        const sheetName = param.sheetName || 'Feuil1';
        const dateIndexes = Array.isArray(param.dateIndexes) ? param.dateIndexes : [];
        let rowHeader;
        if (Array.isArray(param.rowHeader)) {
            // tableau, pas de cast
            rowHeader = param.rowHeader;
        } else if (param.rowHeader && typeof param.rowHeader === 'object') {
            // objet, cast en tableau
            rowHeader = Object.values(param.rowHeader);
        } else {
            // absent (undefined/null)
            rowHeader = [];
        }
        if (rowHeader.length && !Array.isArray(rowHeader[0])) {
            rowHeader = [rowHeader];
        }

        const data = dateIndexes.length ?
            param.data.map(row => row.map((value, index) => {
                if (!dateIndexes.includes(index)) return value;
                if (!value) return null;
                const date = new Date(value);
                if (isNaN(date.getTime())) return value;
                return fixSheetJSDate(date);
            })) :
            param.data;

        // new workbook
        const wb = XLSX.utils.book_new();

        // new worksheet
        const ws = XLSX.utils.aoa_to_sheet([]);
        let origin = "A1";

        // Header
        if (rowHeader.length) {
            XLSX.utils.sheet_add_aoa(ws, rowHeader, { origin: origin });
            origin = 'A' + (rowHeader.length + 1);
        }

        // Data
        XLSX.utils.sheet_add_aoa(ws, data, {
            cellDates: true,
            origin: origin
        });

        // add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, sheetName);

        // write file
        XLSX.writeFile(wb, filename);

        return {
            'status': true
        };
    },
    /* =================================
     *  Reading Workbooks
     * ================================= */
    read: function (param) {
        // parameters
        const filename = param.filename;

        const workbook = XLSX.readFile(filename, {type: 'binary', cellDates: true});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const row = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: false,
            defval: null
        });

        return {
            'status': true,
            'data': row
        };
    }
};

module.exports = {
    call: function (method, param, response) { // The only requirement of an Omnis module is that it implement this function.
        autoSendResponse = true;

        if (methodMap[method]) {
            const result = methodMap[method](param, response);
            if (autoSendResponse) {
                omnis_calls.sendResponse(result, response);
            }

            return true;
        } else {
            throw Error("Method '" + method + "' does not exist");
        }
    }
};
