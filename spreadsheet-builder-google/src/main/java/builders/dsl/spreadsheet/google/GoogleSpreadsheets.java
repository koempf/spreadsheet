package builders.dsl.spreadsheet.google;

import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.services.drive.model.File;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.function.Consumer;

public interface GoogleSpreadsheets {

    String DEFAULT_FIELDS = "webViewLink,id";

    static GoogleSpreadsheets create(HttpRequestInitializer credentials) {
        return new DefaultGoogleSpreadsheets(credentials);
    }

    default File uploadAndConvert(final String name, final Consumer<OutputStream> withOutputStream) {
        return uploadAndConvert(name, DEFAULT_FIELDS, withOutputStream);
    }

    File uploadAndConvert(final String name, final String fields, final Consumer<OutputStream> withOutputStream);

    InputStream convertAndDownload(final String id);

}