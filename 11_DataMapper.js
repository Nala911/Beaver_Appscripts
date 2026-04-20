/**
 * DataMapper — Casts raw sheet row data into strongly typed objects
 * based on the tool's COL_SCHEMA.
 */
var DataMapper = (function() {

    /**
     * Casts a raw row object properties based on tool schema.
     * @param {string} toolKey - Tool registry key.
     * @param {Object} rawRow - Row object with raw string values.
     * @returns {Object} A new object with typed properties.
     */
    function castRow(toolKey, rawRow) {
        var cfg = SyncEngine.getTool(toolKey);
        var schema = (cfg.FORMAT_CONFIG && cfg.FORMAT_CONFIG.COL_SCHEMA) ? cfg.FORMAT_CONFIG.COL_SCHEMA : [];
        var casted = {};

        // Copy metadata (internal fields)
        Object.keys(rawRow).forEach(function(key) {
            if (key.indexOf('_') === 0) {
                casted[key] = rawRow[key];
            }
        });

        schema.forEach(function(col) {
            var val = rawRow[col.header];
            var isMissing = (val === undefined || val === null || (typeof val === 'string' && val.trim() === ''));

            if (isMissing) {
                casted[col.header] = null;
                return;
            }

            switch (col.type) {
                case 'DATE':
                case 'DATETIME':
                    var d = new Date(val);
                    casted[col.header] = isNaN(d.getTime()) ? val : d;
                    break;

                case 'NUMBER':
                    var n = parseFloat(val);
                    casted[col.header] = isNaN(n) ? val : n;
                    break;

                case 'EMAIL_LIST':
                case 'LIST':
                    casted[col.header] = typeof val === 'string' 
                        ? val.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; })
                        : [val];
                    break;

                case 'CHECKBOX':
                    casted[col.header] = (val === true || val === 'true' || val === 'CHECKED' || val === 'Yes');
                    break;

                default:
                    casted[col.header] = val;
            }
        });

        return casted;
    }

    return {
        castRow: castRow
    };

})();
