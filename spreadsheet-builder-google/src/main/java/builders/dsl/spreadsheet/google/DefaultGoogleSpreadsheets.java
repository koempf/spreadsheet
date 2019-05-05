package builders.dsl.spreadsheet.google;

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
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;
import java.util.function.Consumer;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

class DefaultGoogleSpreadsheets implements GoogleSpreadsheets {

    private static final String APPLICATION_NAME = "Google Spreadsheet Builder";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private static final String GOOGLE_SHEET_MIME_TYPE = "application/vnd.google-apps.spreadsheet";

    private final HttpRequestInitializer credentials;

    DefaultGoogleSpreadsheets(HttpRequestInitializer credentials) {
        this.credentials = credentials;
    }

    @Override
    public File updateAndConvert(final String id, final String name, final Iterable<String> fields, final Consumer<OutputStream> withOutputStream) {
        try {
            final ByteArrayOutputStream out = new ByteArrayOutputStream();

            withOutputStream.accept(out);

            final NetHttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();
            final Drive service = new Drive.Builder(httpTransport, JSON_FACTORY, credentials)
                    .setApplicationName(APPLICATION_NAME)
                    .build();

            final File googleFile = new File();
            googleFile.setName(name);
            googleFile.setMimeType(GOOGLE_SHEET_MIME_TYPE);

            String fieldsString = StreamSupport.stream(fields.spliterator(), false).collect(Collectors.joining(","));

            if (id == null) {
                final Drive.Files.Create create = service.files().create(googleFile, new ByteArrayContent(EXCEL_MIME_TYPE, out.toByteArray()));
                return create.setFields(fieldsString).execute();
            }

            final Drive.Files.Update create = service.files().update(id, googleFile, new ByteArrayContent(EXCEL_MIME_TYPE, out.toByteArray()));
            return create.setFields(fieldsString).execute();
        } catch (GeneralSecurityException | IOException e) {
            throw new IllegalStateException("Exception while uploading spreadsheet " + name, e);
        }
    }

    @Override
    public InputStream convertAndDownload(final String id) {
        try {
            Drive service = new Drive.Builder(GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, credentials)
                    .setApplicationName(APPLICATION_NAME)
                    .build();

            Drive.Files.Export export = service.files().export(id, EXCEL_MIME_TYPE);
            return export.executeMediaAsInputStream();
        } catch (GeneralSecurityException | IOException e) {
            throw new IllegalStateException("Exception while downloading spreadsheet " + id, e);
        }
    }
}
