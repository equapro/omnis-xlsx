const omnis_calls = require('omnis_calls');
const fs = require('node:fs');
const XLSX = require("./vendor/sheetjs/xlsx");

/* ============================================================================
 * Dates
 * Omnis envoie les dates en heure MURALE suffixée d'un Z via omnistoiso8601.
 * Exemple: "2026-05-28T09:48:23.0Z". Le Z n'est PAS de l'UTC réel.
 * Les valeurs seront lues telles quelles et converties en serial Excel.
 * ========================================================================== */

// Origine du calendrier Excel : 30/12/1899 (gère le décalage "1900 bissextile" pour les dates modernes).
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

// Format d'affichage par défaut pour les dates.
const DEFAULT_DATE_FMT = 'dd/mm/yyyy hh:mm:ss';

/**
 * Convertit une date ISO Omnis (heure murale) en numéro de série Excel.
 * Interprétation UTC forcée
 *   - indépendant du fuseau du worker ET des bascules été/hiver (DST)
 *   - sans l'artefact d'arrondi LMT de SheetJS.
 *
 * @returns {number|null} le serial, ou null si la valeur n'est pas une date valide.
 */
function toExcelSerial(value) {
    const hasTz = /[zZ]|[+-]\d\d:?\d\d$/.test(value);
    const date = new Date(hasTz ? value : value + 'Z'); // sans offset, JS parserait en LOCAL -> on force UTC
    if (isNaN(date.getTime())) return null;
    return (date.getTime() - EXCEL_EPOCH) / 86400000;
}

/**
 * Normalise param.dateIndexes en tableau d'objets { col, format }.
 * Gère deux formats (période de transition) :
 *   - ancien : tableau d'int       [1, 3, 5]
 *   - nouveau : tableau d'objets   [{ col:1, format:"dd/mm/yyyy" }]
 */
function normalizeDateColumns(dateIndexes) {
    if (!Array.isArray(dateIndexes)) return [];

    return dateIndexes
        .map(entry => (entry && typeof entry === 'object')
            ? { col: entry.col, format: entry.format || DEFAULT_DATE_FMT }   // nouveau format
            : { col: entry, format: DEFAULT_DATE_FMT })                      // ancien format (int)
        .filter(e => Number.isInteger(e.col));
}

/**
 * Applique un format d'affichage aux cellules numériques d'une colonne (format par cellule en SheetJS CE).
 */
function setColumnFormat(ws, col, startRow, count, fmt) {
    for (let r = startRow; r < startRow + count; r++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c: col })];
        if (cell && cell.t === 'n') cell.z = fmt; // cellules nulles/absentes ignorées
    }
}

/**
 * Déduit le bookType SheetJS à partir de l'extension du fichier.
 */
function bookTypeFromName(name) {
    const ext = (name.split('.').pop() || '').toLowerCase();

    return ({ xlsx: 'xlsx', xlsm: 'xlsm', xlsb: 'xlsb', xls: 'biff8', ods: 'ods' })[ext] || 'xlsx';
}

/** Valide et retourne le nom de fichier. */
function requireFilename(param) {
    if (!param.filename || typeof param.filename !== 'string') {
        throw new Error("Paramètre 'filename' manquant ou invalide");
    }

    return param.filename;
}

/* ----------------------------------------------------------------------------
 *  Logique commune (partagée sync/async)
 * -------------------------------------------------------------------------- */

/** Construit le classeur (validation, conversion des dates, en-têtes, formats). */
function buildWorkbook(param) {
    if (!Array.isArray(param.data)) {
        throw new Error("Paramètre 'data' manquant ou n'est pas une liste");
    }

    const sheetName = param.sheetName || 'Feuil1';
    const dateCols = normalizeDateColumns(param.dateIndexes);
    const dateColSet = new Set(dateCols.map(d => d.col)); // lookup rapide pour la conversion

    // Entêtes : accepte un tableau, un objet (cast en tableau) ou rien.
    let rowHeader;
    if (Array.isArray(param.rowHeader)) {
        rowHeader = param.rowHeader;
    } else if (param.rowHeader && typeof param.rowHeader === 'object') {
        rowHeader = Object.values(param.rowHeader);
    } else {
        rowHeader = [];
    }
    if (rowHeader.length && !Array.isArray(rowHeader[0])) {
        rowHeader = [rowHeader];
    }

    // Conversion des colonnes date en serial Excel.
    const data = dateColSet.size ?
        param.data.map(row => row.map((value, index) => {
            if (!dateColSet.has(index)) return value;
            if (!value) return null;
            const serial = toExcelSerial(value);
            return serial === null ? value : serial; // valeur non-date laissée telle quelle
        })) :
        param.data;

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([]);
    let origin = "A1";

    if (rowHeader.length) {
        XLSX.utils.sheet_add_aoa(ws, rowHeader, { origin: origin });
        origin = 'A' + (rowHeader.length + 1);
    }

    XLSX.utils.sheet_add_aoa(ws, data, { origin: origin });

    // Format d'affichage : chaque colonne reçoit son propre format.
    if (dateCols.length) {
        const startRow = rowHeader.length;
        dateCols.forEach(d => setColumnFormat(ws, d.col, startRow, data.length, d.format));
    }

    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    return wb;
}

/**
 * Extrait les données d'un classeur en tableau de lignes.
 */
function readWorkbook(workbook, param) {
    const sheetName = (param.sheetName && workbook.Sheets[param.sheetName])
        ? param.sheetName
        : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const row = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: null
    });

    return { 'status': true, 'data': row };
}

/* ----------------------------------------------------------------------------
 *  Méthodes exposées
 * -------------------------------------------------------------------------- */

/**
 * Description d'une colonne de dates (nouveau format de dateIndexes).
 * @typedef  {Object} DateColumn
 * @property {number} col                     Index (base 0) de la colonne contenant des dates.
 * @property {string} [format='dd/mm/yyyy']   Format d'affichage Excel de la colonne.
 */

/**
 * Paramètres des méthodes d'écriture (write / writeAsync).
 * @typedef  {Object} WriteParam
 * @property {string} filename                                    Chemin complet du fichier à écrire.
 * @property {Array.<Array.<(string|number|boolean|null)>>} data
 * @property {string} [sheetName='Feuil1']                        Nom de l'onglet.
 * @property {Array.<number>|Array.<DateColumn>} [dateIndexes]    Colonnes à convertir en dates.
 * @property {Array.<Array.<*>>|Array.<*>|Object} [rowHeader]     Ligne(s) d'en-tête
 */

/**
 * Paramètres des méthodes de lecture (read / readAsync).
 * @typedef  {Object} ReadParam
 * @property {string} filename     Chemin complet du fichier à lire.
 * @property {string} [sheetName]  Onglet à lire. Si absent ou introuvable, le premier onglet est utilisé.
 */

const methodMap = {
    /**
     * Écriture synchrone (bloquante).
     * @param   {WriteParam} param
     * @returns {{status: boolean}}
     */
    write: function (param) {
        const filename = requireFilename(param);
        XLSX.writeFile(buildWorkbook(param), filename);
        return { 'status': true };
    },

    /**
     * Lecture synchrone (bloquante).
     * @param   {ReadParam} param
     * @returns {{status: boolean, data: Array.<Array.<*>>}}
     */
    read: function (param) {
        const filename = requireFilename(param);
        const workbook = XLSX.readFile(filename, { cellDates: true });
        return readWorkbook(workbook, param);
    },

    /**
     * Écriture asynchrone (non bloquante) : sérialisation en buffer + fs.promises.
     * @param   {WriteParam} param
     * @returns {{status: boolean}}
     */
    writeAsync: async function (param) {
        const filename = requireFilename(param);
        const buffer = XLSX.write(buildWorkbook(param), { type: 'buffer', bookType: bookTypeFromName(filename) });
        await fs.promises.writeFile(filename, buffer);
        return { 'status': true };
    },

    /**
     * Lecture asynchrone (non bloquante) : fs.promises + parsing buffer.
     * @param   {ReadParam} param
     * @returns {{status: boolean, data: Array.<Array.<*>>}}
     */
    readAsync: async function (param) {
        const filename = requireFilename(param);
        const buffer = await fs.promises.readFile(filename);
        const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
        return readWorkbook(workbook, param);
    }
};

module.exports = {
    // call() hybride : gère les méthodes synchrones (try/catch) ET asynchrones (Promise + .catch).
    // Une réponse est toujours envoyée, le worker n'est jamais bloqué.
    call: function (method, param, response) {
        try {
            if (!methodMap[method]) {
                // noinspection ExceptionCaughtLocallyJS
                throw new Error("Method '" + method + "' does not exist");
            }

            const result = methodMap[method](param, response);

            if (result instanceof Promise) {
                result
                    .then(r => { if (!response._writableState?.ended) omnis_calls.sendResponse(r, response); })
                    .catch(err => { if (!response._writableState?.ended) omnis_calls.sendError(response, 400, err?.message || String(err)); });
            } else {
                if (!response._writableState?.ended) omnis_calls.sendResponse(result, response);
            }
        } catch (err) {
            if (response && !response._writableState?.ended) {
                omnis_calls.sendError(response, 400, err?.message || String(err));
            }
        }

        return true;
    }
};
