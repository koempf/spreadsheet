package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.builder.api.FitDimension;
import builders.dsl.spreadsheet.builder.api.PageDefinition;

class FitDimensionHelper implements FitDimension {

    private final PageNode parent;
    private final String widthOrHeight;

    FitDimensionHelper(PageNode parent, String widthOrHeight) {
        this.parent = parent;
        this.widthOrHeight = widthOrHeight;
    }

    @Override
    public PageDefinition to(int numberOfPages) {
        MapNode value = new MapNode();
        value.set(widthOrHeight, numberOfPages);
        parent.node.set("fit", value);
        return parent;
    }

}
