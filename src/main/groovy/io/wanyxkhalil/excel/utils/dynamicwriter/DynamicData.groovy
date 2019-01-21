package io.wanyxkhalil.excel.utils.dynamicwriter

class DynamicData {
    def map = new LinkedHashMap<String, Object>()

    Object get(String key) { map[key] }

    void set(String key, Object value) { map[key] = value }

    void propertySet(String field, value) { this."$field" = value }
}
