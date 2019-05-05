package builders.dsl.spreadsheet.query.google;

import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria;
import builders.dsl.spreadsheet.query.poi.PoiSpreadsheetCriteria;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;

import java.io.IOException;
import java.security.GeneralSecurityException;

public class GoogleSpreadsheetCriteriaFactory {

    public static GoogleSpreadsheetCriteriaFactory create(String spreadsheetId, HttpRequestInitializer credentials) {
        return new GoogleSpreadsheetCriteriaFactory(spreadsheetId, credentials);
    }

    private static final String APPLICATION_NAME = "Google Spreadsheet Criteria";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    private final String id;
    private final HttpRequestInitializer credentials;

    private GoogleSpreadsheetCriteriaFactory(String id, HttpRequestInitializer credentials) {
        this.id = id;
        this.credentials = credentials;
    }

    public SpreadsheetCriteria criteria() throws GeneralSecurityException, IOException {
        Drive service = new Drive.Builder(GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, credentials)
                .setApplicationName(APPLICATION_NAME)
                .build();

        Drive.Files.Export export = service.files().export(id, EXCEL_MIME_TYPE);
        return PoiSpreadsheetCriteria.FACTORY.forStream(export.executeMediaAsInputStream());
    }

}
