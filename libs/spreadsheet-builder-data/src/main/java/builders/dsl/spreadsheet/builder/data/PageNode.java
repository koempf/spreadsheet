package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.builder.api.FitDimension;
import builders.dsl.spreadsheet.builder.api.PageDefinition;

class PageNode extends AbstractNode implements PageDefinition {

    @Override
    public PageDefinition orientation(Keywords.Orientation orientation) {
        node.set("orientation", orientation);
        return this;
    }

    @Override
    public PageDefinition paper(Keywords.Paper paper) {
        node.set("paper", paper);
        return this;
    }

    @Override
    public FitDimension fit(Keywords.Fit widthOrHeight) {
        return new FitDimensionHelper(this, widthOrHeight == Keywords.Fit.WIDTH ? "width" : "height");
    }

}
