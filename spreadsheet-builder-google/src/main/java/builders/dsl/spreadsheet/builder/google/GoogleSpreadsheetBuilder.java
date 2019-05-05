package builders.dsl.spreadsheet.builder.google;

import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder;
import builders.dsl.spreadsheet.builder.api.WorkbookDefinition;
import builders.dsl.spreadsheet.builder.poi.PoiSpreadsheetBuilder;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.ByteArrayContent;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.function.Consumer;

public class GoogleSpreadsheetBuilder implements SpreadsheetBuilder {

    public static GoogleSpreadsheetBuilder create(String name, HttpRequestInitializer credentials) {
        return new GoogleSpreadsheetBuilder(name, credentials);
    }

    private static final String APPLICATION_NAME = "Google Spreadsheet Builder";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private static final String GOOGLE_SHEET_MIME_TYPE = "application/vnd.google-apps.spreadsheet";

    private final String name;
    private final HttpRequestInitializer credentials;

    private String webLink;
    private String id;

    private GoogleSpreadsheetBuilder(String name, HttpRequestInitializer credentials) {
        this.name = name;
        this.credentials = credentials;
    }

    @Override
    public void build(Consumer<WorkbookDefinition> workbookDefinition) {
        try {
            upload(workbookDefinition);
        } catch (GeneralSecurityException | IOException e) {
            throw new IllegalStateException("Exception uploading spreadsheet", e);
        }
    }

    /**
     * Creates an Excel file and uploads it to Google Drive.
     *
     * @param workbookDefinition the definition of the spreadsheet
     *
     * @return web view link of the newly created file
     */
    public String upload(Consumer<WorkbookDefinition> workbookDefinition) throws GeneralSecurityException, IOException {
        final NetHttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

        Drive service = new Drive.Builder(httpTransport, JSON_FACTORY, credentials)
                .setApplicationName(APPLICATION_NAME)
                .build();

        File googleFile = new File();
        googleFile.setName(name);
        googleFile.setMimeType(GOOGLE_SHEET_MIME_TYPE);

        ByteArrayOutputStream out = new ByteArrayOutputStream();

        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.create(out);

        builder.build(workbookDefinition);

        Drive.Files.Create create = service.files().create(googleFile, new ByteArrayContent(EXCEL_MIME_TYPE, out.toByteArray()));

        File created = create.setFields("webViewLink,id").execute();

        id = created.getId();
        webLink = created.getWebViewLink();

        return webLink;
    }

    /**
     * @return the name of the newly crated Spreadsheet as shown in Google Drive
     */
    public String getName() {
        return name;
    }

    /**
     * @return the web link to the last generated spreadsheet or <code>null</code> if no file was generated yet
     */
    public String getWebLink() {
        return webLink;
    }

    /**
     * @return the id of the last generated spreadsheet or <code>null</code> if no file was generated yet
     */
    public String getId() {
        return id;
    }
}
