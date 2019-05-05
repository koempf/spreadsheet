package builders.dsl.spreadsheet.builder.google;

import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder;
import builders.dsl.spreadsheet.builder.api.WorkbookDefinition;
import builders.dsl.spreadsheet.builder.poi.PoiSpreadsheetBuilder;
import builders.dsl.spreadsheet.google.GoogleSpreadsheets;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.services.drive.model.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.function.Consumer;

public class GoogleSpreadsheetBuilder implements SpreadsheetBuilder {

    public static GoogleSpreadsheetBuilder create(String name, HttpRequestInitializer credentials) {
        return create(name, (InputStream) null, GoogleSpreadsheets.create(credentials));
    }

    public static GoogleSpreadsheetBuilder create(String name, GoogleSpreadsheets spreadsheets) {
        return create(name, (InputStream) null, spreadsheets);
    }

    public static GoogleSpreadsheetBuilder create(String name, String spreadsheetTemplateId, HttpRequestInitializer credentials) {
        return create(name, spreadsheetTemplateId, GoogleSpreadsheets.create(credentials));
    }

    public static GoogleSpreadsheetBuilder create(String name, String spreadsheetTemplateId, GoogleSpreadsheets spreadsheets) {
        return create(name, spreadsheets.convertAndDownload(spreadsheetTemplateId), spreadsheets);
    }

    public static GoogleSpreadsheetBuilder create(String name, java.io.File templateFile, HttpRequestInitializer credentials) throws FileNotFoundException {
        return create(name, templateFile, GoogleSpreadsheets.create(credentials));
    }

    public static GoogleSpreadsheetBuilder create(String name, java.io.File templateFile, GoogleSpreadsheets spreadsheets) throws FileNotFoundException {
        return create(name, new FileInputStream(templateFile), spreadsheets);
    }

    public static GoogleSpreadsheetBuilder create(String name, InputStream template, HttpRequestInitializer credentials) {
        return create(name, template, GoogleSpreadsheets.create(credentials));
    }

    public static GoogleSpreadsheetBuilder create(String name, InputStream template, GoogleSpreadsheets spreadsheets) {
        return new GoogleSpreadsheetBuilder(name, template, spreadsheets);
    }

    private final String name;
    private final GoogleSpreadsheets spreadsheets;
    private final InputStream template;

    private String webLink;
    private String id;

    private GoogleSpreadsheetBuilder(String name, InputStream template, GoogleSpreadsheets spreadsheets) {
        this.name = name;
        this.template = template;
        this.spreadsheets = spreadsheets;
    }

    @Override
    public void build(Consumer<WorkbookDefinition> workbookDefinition) {
        File file = buildInternal(workbookDefinition);

        this.webLink = file.getWebViewLink();
        this.id = file.getId();
    }

    private File buildInternal(Consumer<WorkbookDefinition> workbookDefinition) {
        if (template == null) {
            return spreadsheets.uploadAndConvert(name, out -> PoiSpreadsheetBuilder.create(out).build(workbookDefinition));
        }

        return spreadsheets.uploadAndConvert(name, out -> {
            try {
                PoiSpreadsheetBuilder.create(out, template).build(workbookDefinition);
            } catch (IOException e) {
                throw new IllegalArgumentException("Problem reading file's template from " + template, e);
            }
        });
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
