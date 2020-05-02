package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.builder.api.CommentDefinition;

class CommentNode extends AbstractNode implements CommentDefinition {

    @Override
    public CommentDefinition author(String author) {
        node.set("author", author);
        return this;
    }

    @Override
    public CommentDefinition text(String text) {
        node.set("text", text);
        return this;
    }

}
