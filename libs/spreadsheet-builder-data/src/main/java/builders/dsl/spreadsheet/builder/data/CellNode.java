package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.builder.api.*;
import builders.dsl.spreadsheet.impl.HeightModifier;
import builders.dsl.spreadsheet.impl.WidthModifier;

import java.util.function.Consumer;

class CellNode extends AbstractNode implements CellDefinition {

    private static final double WIDTH_POINTS_PER_CM = 4.6666666666666666666667;
    private static final double WIDTH_POINTS_PER_INCH = 12;
    private static final double HEIGHT_POINTS_PER_CM = 28;
    private static final double HEIGHT_POINTS_PER_INCH = 72;

    @Override
    public CellDefinition value(Object value) {
        node.set("value", value);
        return this;
    }

    @Override
    public CellDefinition name(String name) {
        node.set("name", name);
        return this;
    }

    @Override
    public CellDefinition formula(String formula) {
        node.set("formula", formula);
        return this;
    }

    @Override
    public CellDefinition comment(Consumer<CommentDefinition> commentDefinition) {
        CommentNode commentNode = new CommentNode();
        commentDefinition.accept(commentNode);
        node.set("comment", commentNode);
        return this;
    }

    @Override
    public LinkDefinition link(Keywords.To to) {
        return new LinkHelper(this);
    }

    @Override
    public CellDefinition colspan(int span) {
        node.set("colspan", span);
        return this;
    }

    @Override
    public CellDefinition rowspan(int span) {
        node.set("rowspan", span);
        return this;
    }

    @Override
    public DimensionModifier width(double width) {
        node.set("width", width);
        return new WidthModifier(this, width, WIDTH_POINTS_PER_CM, WIDTH_POINTS_PER_INCH);
    }

    @Override
    public DimensionModifier height(double height) {
        node.set("height", height);
        return new HeightModifier(this, height, HEIGHT_POINTS_PER_CM, HEIGHT_POINTS_PER_INCH);
    }

    @Override
    public CellDefinition width(Keywords.Auto auto) {
        node.set("width", "auto");
        return this;
    }

    @Override
    public CellDefinition text(String text) {
        node.add("text", text);
        return this;
    }

    @Override
    public CellDefinition text(String text, Consumer<FontDefinition> fontConfiguration) {
        FontNode fontNode = new FontNode();
        fontConfiguration.accept(fontNode);

        MapNode mapNode = new MapNode();
        mapNode.set("content", text);
        mapNode.set("font", fontNode);

        node.add("text", mapNode);

        return this;
    }

    @Override
    public ImageCreator png(Keywords.Image image) {
        return new ImageHelper(this, "png");
    }

    @Override
    public ImageCreator jpeg(Keywords.Image image) {
        return new ImageHelper(this, "jpeg");
    }

    @Override
    public ImageCreator pict(Keywords.Image image) {
        return new ImageHelper(this, "pict");
    }

    @Override
    public ImageCreator emf(Keywords.Image image) {
        return new ImageHelper(this, "emf");
    }

    @Override
    public ImageCreator wmf(Keywords.Image image) {
        return new ImageHelper(this, "wmf");
    }

    @Override
    public ImageCreator dib(Keywords.Image image) {
        return new ImageHelper(this, "dib");
    }

    @Override
    public CellDefinition styles(Iterable<String> styles, Iterable<Consumer<CellStyleDefinition>> styleDefinitions) {
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

    public void setColumn(int column) {
        node.set("column", column);
    }
}
