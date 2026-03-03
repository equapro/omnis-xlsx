const omnis_calls = require('omnis_calls');
var XLSX = require("./vendor/sheetjs/xlsx");

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
        var filename = param.filename;
        var sheetName = param.sheetName || 'Feuil1';
        var dateIndexes = param.dateIndexes;
        var rowHeader = param.rowHeader;

        var data;
        if (dateIndexes.length) {
            // dates parsing
            data = param.data.map(function (row) {
                // line
                return row.map((value, index) => {
                    // cell
                    if (!dateIndexes.includes(index)) {
                        return value;
                    }

                    // Valeur vide
                    if (!value) {
                        return null;
                    }

                    // Transformation et validation de la date
                    let date = new Date(value);
                    if (!(date instanceof Date)) {
                        return value;
                    }

                    // Correction de la Timezone et de l'erreur de précision de la librairie
                    return fixSheetJSDate(date);
                });
            });
        } else {
            data = param.data;
        }

        // new workbook
        var wb = XLSX.utils.book_new();
        
        // new worksheet
        const ws = XLSX.utils.aoa_to_sheet([]);
        var origin = "A1";
        
        // Header
        if (rowHeader && rowHeader.length > 0) {
        	XLSX.utils.sheet_add_aoa(ws, rowHeader, { origin: origin });
        	origin = "A2";
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
        var filename = param.filename;

        var workbook = XLSX.readFile(filename, {type: 'binary', cellDates: true});
        var sheet = workbook.Sheets[workbook.SheetNames[0]];

        var row = XLSX.utils.sheet_to_json(sheet, {
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
