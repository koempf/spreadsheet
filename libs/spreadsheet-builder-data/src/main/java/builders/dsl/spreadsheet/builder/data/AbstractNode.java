package builders.dsl.spreadsheet.builder.data;

abstract class AbstractNode implements Node {

    protected final MapNode node = new MapNode();

    @Override
    public final Object getContent() {
        return node.getContent();
    }

}
