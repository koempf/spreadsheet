package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.builder.api.CellStyleDefinition;
import builders.dsl.spreadsheet.builder.api.SheetDefinition;
import builders.dsl.spreadsheet.builder.api.WorkbookDefinition;

import java.util.function.Consumer;

class WorkbookNode extends AbstractNode implements WorkbookDefinition {

    @Override
    public WorkbookDefinition sheet(String name, Consumer<SheetDefinition> sheetDefinition) {
        SheetNode sheet = new SheetNode(name);
        sheetDefinition.accept(sheet);
        node.add("sheets", sheet);
        return this;
    }

    @Override
    public WorkbookDefinition style(String name, Consumer<CellStyleDefinition> styleDefinition) {
        CellStyleNode style = new CellStyleNode(name);
        styleDefinition.accept(style);
        node.add("styles", style);
        return this;
    }
}
