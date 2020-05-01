package builders.dsl.spreadsheet.query.google;

import builders.dsl.spreadsheet.google.GoogleSpreadsheets;
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria;
import builders.dsl.spreadsheet.query.poi.PoiSpreadsheetCriteria;
import com.google.api.client.http.HttpRequestInitializer;

public class GoogleSpreadsheetCriteriaFactory {

    public static GoogleSpreadsheetCriteriaFactory create(String spreadsheetId, GoogleSpreadsheets spreadsheets) {
        return create(spreadsheetId, spreadsheets);
    }

    public static GoogleSpreadsheetCriteriaFactory create(String spreadsheetId, HttpRequestInitializer credentials) {
        return new GoogleSpreadsheetCriteriaFactory(spreadsheetId, GoogleSpreadsheets.create(credentials));
    }

    private final String id;
    private final GoogleSpreadsheets spreadsheets;

    private GoogleSpreadsheetCriteriaFactory(String id, GoogleSpreadsheets spreadsheets) {
        this.id = id;
        this.spreadsheets = spreadsheets;
    }

    public SpreadsheetCriteria criteria() {
        return PoiSpreadsheetCriteria.FACTORY.forStream(spreadsheets.convertAndDownload(id));
    }

}
