package builders.dsl.spreadsheet.google;

import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.services.drive.model.File;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.function.Consumer;

public interface GoogleSpreadsheets {

    Iterable<String> DEFAULT_FIELDS = Arrays.asList("webViewLink", "id");

    static GoogleSpreadsheets create(HttpRequestInitializer credentials) {
        return new DefaultGoogleSpreadsheets(credentials);
    }

    default File uploadAndConvert(final String name, final Consumer<OutputStream> withOutputStream) {
        return uploadAndConvert(name, DEFAULT_FIELDS, withOutputStream);
    }

    default File uploadAndConvert(final String name, final Iterable<String> fields, final Consumer<OutputStream> withOutputStream) {
        return updateAndConvert(null, name, fields, withOutputStream);
    }

    default File updateAndConvert(String id, final String name, final Consumer<OutputStream> withOutputStream) {
        return updateAndConvert(id, name, DEFAULT_FIELDS, withOutputStream);
    }

    File updateAndConvert(String id, final String name, final Iterable<String> fields, final Consumer<OutputStream> withOutputStream);

    InputStream convertAndDownload(final String id);

}