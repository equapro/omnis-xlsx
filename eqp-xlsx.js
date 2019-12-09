const omnis_calls = require('omnis_calls');
const XLSX = require('xlsx');

let autoSendResponse = true; // Set to false in methods which should not send a response to Omnis when they exit. (e.g. async methods)

const methodMap = {
    /* =================================
     *  Writing Workbooks
     * ================================= */
    write: function (param) {
        // parameters
        var filename = param.filename;
        var sheetName = param.sheetName || 'Feuil1';
        var dateIndexes = param.dateIndexes;

        var data;
        if (dateIndexes.length) {
            // dates parsing
            data = param.data.map(function (row) {
                // line
                return row.map((value, index) => {
                    // cell
                    if (-1 === dateIndexes.indexOf(index)) {
                        return value;
                    }

                    return new Date(value);
                });
            });
        } else {
            data = param.data;
        }

        // new workbook
        var wb = XLSX.utils.book_new();
        // new worksheet
        var ws = XLSX.utils.aoa_to_sheet(data, {cellDates: true});
        // add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        // write file
        XLSX.writeFile(wb, filename);

        return {
            'status': true
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
