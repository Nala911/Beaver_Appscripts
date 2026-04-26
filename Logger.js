/**
 * Developer Logging System
 * Version: 7.0 (Silent Architecture)
 */

var Logger = (function () {
    return {
        setLoggingState: function (enabled) {},
        setRunId: function (id) { return null; },
        info: function (src, ref, msg, ctx) {},
        success: function (src, ref, msg, ctx) {},
        warn: function (src, ref, msg, ctx) {},
        debug: function (src, ref, msg, ctx) {},
        error: function (src, ref, err, ctx) {},
        step: function (src, ref, name) {},
        flushLogs: function () {},
        clearLogs: function () {},
        isEnabled: function() { return false; },

        run: function (toolKey, reference, callback, forceLog) {
            try {
                return callback();
            } catch (e) {
                throw e; 
            }
        },

        wrap: function (source, reference, func) {
            return function() {
                try {
                    return func.apply(this, arguments);
                } catch(e) {
                    throw e;
                }
            };
        }
    };
})();

