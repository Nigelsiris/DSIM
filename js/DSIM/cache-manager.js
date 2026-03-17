// Cache Manager for DSIM
class CacheManager {
    constructor() {
        this.VERSION = '1.0';
        this.CACHE_PREFIX = 'dsim_';
    }

    // Set item with expiration
    setItem(key, value, expirationMinutes = 30) {
        const item = {
            value: value,
            timestamp: new Date().getTime(),
            expiresIn: expirationMinutes * 60 * 1000
        };
        localStorage.setItem(this.CACHE_PREFIX + key, JSON.stringify(item));
    }

    // Get item if not expired
    getItem(key) {
        const item = localStorage.getItem(this.CACHE_PREFIX + key);
        if (!item) return null;

        const parsedItem = JSON.parse(item);
        const now = new Date().getTime();
        
        if (now - parsedItem.timestamp > parsedItem.expiresIn) {
            localStorage.removeItem(this.CACHE_PREFIX + key);
            return null;
        }

        return parsedItem.value;
    }

    // Clear all cached items
    clearCache() {
        Object.keys(localStorage)
            .filter(key => key.startsWith(this.CACHE_PREFIX))
            .forEach(key => localStorage.removeItem(key));
    }

    // Save form data temporarily
    saveFormData(formId, data) {
        this.setItem(`form_${formId}`, data, 60); // Cache for 1 hour
    }

    // Restore form data
    restoreFormData(formId) {
        return this.getItem(`form_${formId}`);
    }
}

// Initialize cache manager
window.cacheManager = new CacheManager();