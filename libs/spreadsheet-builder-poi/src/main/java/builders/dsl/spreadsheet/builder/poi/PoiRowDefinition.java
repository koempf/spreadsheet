package builders.dsl.spreadsheet.builder.poi;

import builders.dsl.spreadsheet.builder.api.RowDefinition;
import builders.dsl.spreadsheet.impl.AbstractCellDefinition;
import builders.dsl.spreadsheet.impl.AbstractRowDefinition;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

class PoiRowDefinition extends AbstractRowDefinition implements RowDefinition {

    private final Row xssfRow;

    PoiRowDefinition(PoiSheetDefinition sheet, Row xssfRow) {
        super(sheet);
        this.xssfRow = xssfRow;
    }

    @Override
    protected AbstractCellDefinition createCell(int zeroBasedCellNumber) {
        Cell cell = xssfRow.getCell(zeroBasedCellNumber);

        if (cell == null) {
            cell = xssfRow.createCell(zeroBasedCellNumber);
        }

        return new PoiCellDefinition(this, cell);
    }

    @Override
    protected void handleSpans(AbstractCellDefinition cell) {
        if (cell instanceof PoiCellDefinition) {
            PoiCellDefinition poiCell = (PoiCellDefinition) cell;
            if (poiCell.getColspan() > 1 || poiCell.getRowspan() > 1) {
                xssfRow.getSheet().addMergedRegion(poiCell.getCellRangeAddress());
            }
        } else {
            throw new IllegalArgumentException("Unsupported cell: " + cell);
        }
    }

    @Override
    public int getNumber() {
        return xssfRow.getRowNum() + 1;
    }

    public PoiSheetDefinition getSheet() {
        return (PoiSheetDefinition) sheet;
    }

    protected Row getRow() {
        return xssfRow;
    }


    protected void doCreateGroup(int startPosition, int endPosition, boolean collapsed) {
        getSheet().getSheet().groupColumn(startPosition, endPosition);
        if (collapsed) {
            getSheet().getSheet().setColumnGroupCollapsed(endPosition, true);
        }
    }

}
