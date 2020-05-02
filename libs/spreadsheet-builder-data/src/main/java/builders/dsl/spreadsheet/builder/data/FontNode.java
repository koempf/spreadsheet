package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.api.Color;
import builders.dsl.spreadsheet.api.FontStyle;
import builders.dsl.spreadsheet.builder.api.FontDefinition;

class FontNode extends AbstractNode implements FontDefinition {
    
    @Override
    public FontDefinition color(String hexColor) {
        node.set("color", hexColor);
        return this;
    }

    @Override
    public FontDefinition color(Color color) {
        node.set("color", color.getHex());
        return this;
    }

    @Override
    public FontDefinition size(int size) {
        node.set("size", size);
        return this;
    }

    @Override
    public FontDefinition name(String name) {
        node.set("name", name);
        return this;
    }

    @Override
    public FontDefinition style(FontStyle first, FontStyle... other) {
        node.add("styles", first.name());
        for(FontStyle style : other) {
            node.add("styles", style.name());
        }
        return this;
    }
    
}
