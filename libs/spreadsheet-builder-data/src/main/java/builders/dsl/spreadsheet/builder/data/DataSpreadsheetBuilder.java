package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder;
import builders.dsl.spreadsheet.builder.api.WorkbookDefinition;

import java.util.Map;
import java.util.function.Consumer;

public class DataSpreadsheetBuilder implements SpreadsheetBuilder {

    private WorkbookNode node;

    @Override
    public void build(Consumer<WorkbookDefinition> workbookDefinition) {
        this.node = new WorkbookNode();
        workbookDefinition.accept(node);
    }

    @SuppressWarnings("unchecked")
    public Map<String, Object> getData() {
        return (Map<String, Object>) node.getContent();
    }

}
