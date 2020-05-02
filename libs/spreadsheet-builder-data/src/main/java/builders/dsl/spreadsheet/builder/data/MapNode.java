package builders.dsl.spreadsheet.builder.data;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

class MapNode implements Node {

    private final Map<String, Object> map = new LinkedHashMap<>();

    public void add(String key, Node newItem) {
        add(key, newItem.getContent());
    }

    @SuppressWarnings("unchecked")
    public void add(String key, Object newItem) {
        List<Object> list = (List<Object>) map.computeIfAbsent(key, k -> new ArrayList<>());
        list.add(newItem);
    }

    public void set(String key, Node value) {
        set(key, value.getContent());
    }

    public void set(String key, Object value) {
        map.put(key, value);
    }

    @Override
    public Object getContent() {
        return map;
    }
}
