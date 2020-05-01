package builders.dsl.spreadsheet.api;

public interface Cell extends Spannable {

    int getColumn();
    Object getValue();
    String getColumnAsString();
    <T> T read(Class<T> type);
    Row getRow();

    String getName();
    Comment getComment();
    CellStyle getStyle();

    Cell getAbove();
    Cell getBelow();
    Cell getLeft();
    Cell getRight();
    Cell getAboveLeft();
    Cell getAboveRight();
    Cell getBelowLeft();
    Cell getBelowRight();

}
