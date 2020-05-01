package builders.dsl.spreadsheet.api;

import java.util.Collection;

public interface Row {

    int getNumber();
    Sheet getSheet();

    Collection<? extends Cell> getCells();

    Row getAbove();
    Row getAbove(int howMany);
    Row getBelow();
    Row getBelow(int howMany);

}
