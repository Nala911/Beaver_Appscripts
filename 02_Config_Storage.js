var __propsCache = {};

/**
 * Helper to get the appropriate properties store.
 */
function _App_getStore_(storeType) {
    switch (storeType) {
        case STORE_TYPES.DOCUMENT: return PropertiesService.getDocumentProperties();
        case STORE_TYPES.USER: return PropertiesService.getUserProperties();
        case STORE_TYPES.SCRIPT: return PropertiesService.getScriptProperties();
        default: throw new Error("Invalid store type: " + storeType);
    }
}

/**
 * Helper to get the appropriate cache store.
 */
function _App_getCacheStore_(storeType) {
    switch (storeType) {
        case STORE_TYPES.DOCUMENT: return CacheService.getDocumentCache() || CacheService.getScriptCache();
        case STORE_TYPES.USER: return CacheService.getUserCache();
        case STORE_TYPES.SCRIPT: return CacheService.getScriptCache();
        default: return null;
    }
}

/**
 * Retrieves a property from the registry. Automatically parses JSON if configured.
 * @param {Object} propConfig An entry from APP_PROPS
 * @returns {*} The value or null if not found
 */
function _App_getProperty(propConfig) {
    var cacheKey = propConfig.key;
    
    // 1. Fast Memory Cache
    if (__propsCache.hasOwnProperty(cacheKey)) {
        return __propsCache[cacheKey];
    }

    // 2. CacheService Layer
    var cacheStore = _App_getCacheStore_(propConfig.store);
    var valStr = cacheStore ? cacheStore.get(cacheKey) : null;
    
    // 3. PropertiesService Fallback
    if (valStr === null) {
        var store = _App_getStore_(propConfig.store);
        valStr = store.getProperty(cacheKey);
        if (valStr && cacheStore) {
            cacheStore.put(cacheKey, valStr, 21600); // Max 6 hours
        }
    }

    if (!valStr) return null;

    var result = valStr;
    if (propConfig.isJson) {
        try {
            result = JSON.parse(valStr);
        } catch (e) {
            return null;
        }
    }

    // Save to memory cache for subsequent calls
    __propsCache[cacheKey] = result;
    return result;
}

function _App_getRawProperty(propConfig) {
    return _App_getStore_(propConfig.store).getProperty(propConfig.key);
}

/**
 * Sets a property in the registry. Automatically stringifies JSON if configured.
 * @param {Object} propConfig An entry from APP_PROPS
 * @param {*} value The value to set (can be an object or primitive)
 */
function _App_setProperty(propConfig, value) {
    var valToStore = propConfig.isJson ? JSON.stringify(value) : String(value);
    
    // Save to DB
    var store = _App_getStore_(propConfig.store);
    store.setProperty(propConfig.key, valToStore);
    
    // Update Caches
    __propsCache[propConfig.key] = value;
    var cacheStore = _App_getCacheStore_(propConfig.store);
    if (cacheStore) {
        cacheStore.put(propConfig.key, valToStore, 21600);
    }
}

/**
 * Deletes a property from the registry.
 * @param {Object} propConfig An entry from APP_PROPS
 */
function _App_deleteProperty(propConfig) {
    // Delete from DB
    var store = _App_getStore_(propConfig.store);
    store.deleteProperty(propConfig.key);
    
    // Clear Caches
    delete __propsCache[propConfig.key];
    var cacheStore = _App_getCacheStore_(propConfig.store);
    if (cacheStore) {
        cacheStore.remove(propConfig.key);
    }
}
