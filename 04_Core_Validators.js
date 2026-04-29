// ==========================================
// Centralized Data Validators
// ==========================================
var SYSTEM_VALIDATORS = {
    EMAIL: function(val) { return typeof val === 'string' && val.indexOf('@') !== -1; },
    DATE: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    DATETIME: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); }
};

