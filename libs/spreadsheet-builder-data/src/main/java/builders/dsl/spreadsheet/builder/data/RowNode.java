package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.builder.api.CellDefinition;
import builders.dsl.spreadsheet.builder.api.CellStyleDefinition;
import builders.dsl.spreadsheet.builder.api.RowDefinition;

import java.util.Collections;
import java.util.function.Consumer;

class RowNode extends AbstractNode implements RowDefinition {

    @Override
    public RowDefinition cell(Consumer<CellDefinition> cellDefinition) {
        CellNode cellNode =  new CellNode();
        cellDefinition.accept(cellNode);
        node.add("cells", cellNode);
        return this;
    }

    @Override
    public RowDefinition cell(int column, Consumer<CellDefinition> cellDefinition) {
        CellNode cellNode =  new CellNode();
        cellNode.setColumn(column);
        cellDefinition.accept(cellNode);
        node.add("cells", cellNode);
        return this;
    }

    @Override
    public RowDefinition group(Consumer<RowDefinition> insideGroupDefinition) {
        RowNode rowNode = new RowNode();
        insideGroupDefinition.accept(rowNode);
        node.add("cells", Collections.singletonMap("group", rowNode.getContent()));
        return this;
    }

    @Override
    public RowDefinition collapse(Consumer<RowDefinition> insideGroupDefinition) {
        RowNode rowNode = new RowNode();
        insideGroupDefinition.accept(rowNode);
        node.add("cells", Collections.singletonMap("collapse", rowNode.getContent()));
        return this;
    }

    @Override
    public RowDefinition styles(Iterable<String> styles, Iterable<Consumer<CellStyleDefinition>> styleDefinitions) {
        for (String style : styles) {
            node.add("styles", style);
        }
        for (Consumer<CellStyleDefinition> styleDefinition : styleDefinitions) {
            CellStyleNode cellStyleNode = new CellStyleNode();
            styleDefinition.accept(cellStyleNode);
            node.add("styles", cellStyleNode);
        }
        return this;
    }

    public void setNumber(int number) {
        node.set("number", number);
    }
}
