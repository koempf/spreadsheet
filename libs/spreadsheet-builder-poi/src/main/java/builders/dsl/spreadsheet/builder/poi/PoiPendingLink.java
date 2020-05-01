package builders.dsl.spreadsheet.builder.poi;

import builders.dsl.spreadsheet.impl.AbstractPendingLink;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Pending link is a link which needs to be resolved at the end of the build when all the named references are known.
 */
class PoiPendingLink extends AbstractPendingLink {

    PoiPendingLink(PoiCellDefinition cell, final String name) {
        super(cell, name);
    }

    public void resolve() {
        Workbook workbook = getPoiCell().getCell().getRow().getSheet().getWorkbook();
        Name xssfName = workbook.getName(getName());

        if (xssfName == null) {
            throw new IllegalArgumentException("Name " + getName() + " does not exist!");
        }


        Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
        link.setAddress(xssfName.getRefersToFormula());

        getPoiCell().getCell().setHyperlink(link);
    }

    private PoiCellDefinition getPoiCell() {
        return (PoiCellDefinition) getCell();
    }

}
