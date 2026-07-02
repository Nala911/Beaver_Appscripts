/**
 * Developer Logging System
 * Version: 7.0 (Silent Architecture)
 */

var Logger = (function () {
    var isLoggingEnabled = false;

    return {
        setLoggingState: function (enabled) {
            isLoggingEnabled = !!enabled;
        },
        setRunId: function (id) { return id; },
        info: function (src, ref, msg, ctx) {
            if (isLoggingEnabled) console.info("[" + src + "::" + ref + "] " + msg, ctx || "");
        },
        success: function (src, ref, msg, ctx) {
            if (isLoggingEnabled) console.log("✅ [" + src + "::" + ref + "] " + msg, ctx || "");
        },
        warn: function (src, ref, msg, ctx) {
            if (isLoggingEnabled) console.warn("⚠️ [" + src + "::" + ref + "] " + msg, ctx || "");
        },
        debug: function (src, ref, msg, ctx) {
            if (isLoggingEnabled) console.log("🐞 [" + src + "::" + ref + "] " + msg, ctx || "");
        },
        error: function (src, ref, err, ctx) {
            var errMsg = err ? (err.message || String(err)) : "Unknown Error";
            console.error("❌ [" + src + "::" + ref + "] " + errMsg, ctx || "", err && err.stack ? err.stack : "");
        },
        step: function (src, ref, name) {
            if (isLoggingEnabled) console.log("➔ [" + src + "::" + ref + "] Step: " + name);
        },
        flushLogs: function () {},
        clearLogs: function () {},
        isEnabled: function() { return isLoggingEnabled; },

        run: function (toolKey, reference, callback, forceLog) {
            var oldState = isLoggingEnabled;
            if (forceLog) isLoggingEnabled = true;
            try {
                if (isLoggingEnabled) console.log("➔ [" + toolKey + "] Starting execution: " + reference);
                var res = callback();
                if (isLoggingEnabled) console.log("➔ [" + toolKey + "] Finished execution: " + reference);
                return res;
            } catch (e) {
                console.error("❌ [" + toolKey + "] Failed execution: " + reference, e);
                throw e; 
            } finally {
                isLoggingEnabled = oldState;
            }
        },

        wrap: function (source, reference, func) {
            return function() {
                var oldState = isLoggingEnabled;
                try {
                    return func.apply(this, arguments);
                } catch(e) {
                    console.error("❌ [" + source + "] Function error: " + reference, e);
                    throw e;
                } finally {
                    isLoggingEnabled = oldState;
                }
            };
        }
    };
})();

