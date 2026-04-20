/**
 * STORAGE LAYER (Properties & Cache)
 * ==========================================
 * Unified interface for PropertiesService and CacheService.
 */

Object.assign(App.Storage, (function() {
    var _memoryCache = {};

    function _getStore(type) {
        switch (type) {
            case App.Config.STORE_TYPES.DOCUMENT: return PropertiesService.getDocumentProperties();
            case App.Config.STORE_TYPES.USER: return PropertiesService.getUserProperties();
            case App.Config.STORE_TYPES.SCRIPT: return PropertiesService.getScriptProperties();
            default: throw new Error("Invalid store type: " + type);
        }
    }

    function _getCache(type) {
        switch (type) {
            case App.Config.STORE_TYPES.DOCUMENT: return CacheService.getDocumentCache();
            case App.Config.STORE_TYPES.USER: return CacheService.getUserCache();
            case App.Config.STORE_TYPES.SCRIPT: return CacheService.getScriptCache();
            default: return null;
        }
    }

    return {
        get: function(prop) {
            var key = prop.key;
            if (_memoryCache.hasOwnProperty(key)) return _memoryCache[key];

            var cache = _getCache(prop.store);
            var val = cache ? cache.get(key) : null;

            if (val === null) {
                val = _getStore(prop.store).getProperty(key);
                if (val && cache) cache.put(key, val, 21600);
            }

            if (!val) return null;
            var res = prop.isJson ? JSON.parse(val) : val;
            _memoryCache[key] = res;
            return res;
        },

        set: function(prop, value) {
            var str = prop.isJson ? JSON.stringify(value) : String(value);
            _getStore(prop.store).setProperty(prop.key, str);
            _memoryCache[prop.key] = value;
            var cache = _getCache(prop.store);
            if (cache) cache.put(prop.key, str, 21600);
        },

        delete: function(prop) {
            _getStore(prop.store).deleteProperty(prop.key);
            delete _memoryCache[prop.key];
            var cache = _getCache(prop.store);
            if (cache) cache.remove(prop.key);
        }
    };
})());

// Backward Compatibility Aliases
function _App_getProperty(p) { return App.Storage.get(p); }
function _App_setProperty(p, v) { return App.Storage.set(p, v); }
function _App_deleteProperty(p) { return App.Storage.delete(p); }
