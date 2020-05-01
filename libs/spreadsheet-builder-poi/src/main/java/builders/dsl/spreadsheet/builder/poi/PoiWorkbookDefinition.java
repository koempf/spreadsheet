package builders.dsl.spreadsheet.builder.poi;

import builders.dsl.spreadsheet.builder.api.WorkbookDefinition;
import builders.dsl.spreadsheet.impl.AbstractWorkbookDefinition;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class PoiWorkbookDefinition extends AbstractWorkbookDefinition implements WorkbookDefinition {

    private final Workbook workbook;

    PoiWorkbookDefinition(Workbook workbook) {
        if (!(workbook instanceof XSSFWorkbook) && !(workbook instanceof SXSSFWorkbook)) {
            throw new IllegalArgumentException("Only XSSF and SXSSF workbooks are supported");
        }
        this.workbook = workbook;
    }

    @Override
    protected PoiSheetDefinition createSheet(String name) {
        Sheet sheet = workbook.getSheet(WorkbookUtil.createSafeSheetName(name));
        return new PoiSheetDefinition(this, sheet != null ? sheet : workbook.createSheet(WorkbookUtil.createSafeSheetName(name)));
    }

    @Override
    protected PoiCellStyleDefinition createCellStyle() {
        return new PoiCellStyleDefinition(this);
    }

    Workbook getWorkbook() {
        return workbook;
    }

    void addPendingLink(String ref, PoiCellDefinition cell) {
        addPendingLink(new PoiPendingLink(cell, ref));
    }

    short parseColor(String hexColor) {
        // TODO: implement
        throw new UnsupportedOperationException();
    }

    protected void doLock(Sheet sheet) {
        if (sheet instanceof XSSFSheet) {
            ((XSSFSheet) sheet).enableLocking();
        }
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).enableLocking();
        }
    }

    String getReference(Cell cell) {
        if (cell instanceof XSSFCell) {
            XSSFCell xssfCell = (XSSFCell) cell;
            return xssfCell.getReference();
        }
        return cell.getAddress().formatAsString();
    }
}
