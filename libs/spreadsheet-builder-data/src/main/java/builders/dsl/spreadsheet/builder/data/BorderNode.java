package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.api.BorderStyle;
import builders.dsl.spreadsheet.api.Color;
import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.builder.api.BorderDefinition;

import java.util.List;

class BorderNode extends AbstractNode implements BorderDefinition {

    @Override
    public BorderDefinition style(BorderStyle style) {
        node.set("style", style);
        return this;
    }

    @Override
    public BorderDefinition color(String hexColor) {
        node.set("color", hexColor);
        return this;
    }

    @Override
    public BorderDefinition color(Color color) {
        node.set("style", color.getHex());
        return this;
    }


    public void setSide(List<Keywords.BorderSide> sides) {
        node.set("side", sides);
    }

}
