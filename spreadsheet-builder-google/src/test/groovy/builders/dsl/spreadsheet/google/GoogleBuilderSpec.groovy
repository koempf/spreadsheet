package builders.dsl.spreadsheet.google

import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder
import builders.dsl.spreadsheet.builder.google.GoogleSpreadsheetBuilder
import builders.dsl.spreadsheet.builder.tck.AbstractBuilderSpec
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria
import builders.dsl.spreadsheet.query.google.GoogleSpreadsheetCriteriaFactory
import com.google.api.client.auth.oauth2.AuthorizationCodeTokenRequest
import com.google.api.client.auth.oauth2.ClientParametersAuthentication
import com.google.api.client.auth.oauth2.TokenResponse
import com.google.api.client.auth.oauth2.TokenResponseException
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.http.GenericUrl
import com.google.api.client.http.HttpRequestInitializer
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.JsonFactory
import com.google.api.client.json.jackson2.JacksonFactory
import com.google.api.client.util.store.FileDataStoreFactory
import com.google.api.services.drive.Drive
import com.google.api.services.drive.DriveScopes
import com.google.api.services.drive.model.File
import com.google.api.services.drive.model.FileList
import spock.lang.Requires

import java.awt.Desktop

@Requires({ System.getenv('DSL_SPREADSHEET_BUILDER_GOOGLE_CLIENT_SECRET')})
class GoogleBuilderSpec extends AbstractBuilderSpec {

    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance()
    private static final String TOKENS_DIRECTORY_PATH = "tokens"
    private static final List<String> SCOPES = Collections.singletonList(DriveScopes.DRIVE)
    private static final String CREDENTIALS_FILE_PATH = "credentials.json"
    private static final GoogleClientSecrets GOOGLE_SECRETS = buildGoogleSecrets()
    private static final String FILENAME = "DELETE ME"

    HttpRequestInitializer credentials = buildCredentials(GOOGLE_SECRETS)
    GoogleSpreadsheetBuilder builder = GoogleSpreadsheetBuilder.create(FILENAME, credentials)

    void cleanup() {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

        Drive service = new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, credentials)
                .setApplicationName('Google Builder Spec')
                .build();

        FileList files = null

        while (!files || files.getIncompleteSearch()) {
            files = service.files().list().execute()
            for (File file in files.getFiles()) {
                if (file.getName() == FILENAME) {
                    service.files().delete(file.getId())
                }
            }
        }
    }

    @Override
    protected SpreadsheetCriteria createCriteria() {
        assert builder.id
        return GoogleSpreadsheetCriteriaFactory.create(builder.id, credentials).criteria()
    }

    @Override
    protected SpreadsheetBuilder createSpreadsheetBuilder() {
        return builder
    }

    @Override
    protected void openSpreadsheet() {
        println "Open browser at $builder.webLink"
    }

    @Override
    protected boolean isVeryHiddenSupported() { false }

    @Override
    protected boolean isFillSupported() { false }

    // google exports empty rows as well
    @Override
    protected int getExpectedAllRowsSize() { 39663 }

    @Override
    protected int getExpectedAllCellSize() { 80127 }

    /**
     * Tries to open the file in the browser. Only works locally on Mac at the moment. Ignored otherwise.
     * Main purpose of this method is to quickly open the generated file for manual review.
     * @param uri uri to be opened
     */
    private static void open(String uri) {
        try {
            if (uri && Desktop.desktopSupported && Desktop.desktop.isSupported(Desktop.Action.BROWSE)) {
                Desktop.desktop.browse(new URI(uri))
                Thread.sleep(10000)
            }
        } catch(ignored) {
            // CI
        }
    }

    static void requestAccessToken() throws IOException {
        try {
            TokenResponse response = new AuthorizationCodeTokenRequest(new NetHttpTransport(),
                    new JacksonFactory(), new GenericUrl("https://server.example.com/token"),
                    "SplxlOBeZQQYbYS6WxSbIA").setRedirectUri("https://client.example.com/rd")
                                             .setClientAuthentication(
                    new ClientParametersAuthentication("s6BhdRkqt3", "7Fjfp0ZBr1KtDRbnfVdmIw")).execute()
            System.out.println("Access token: " + response.getAccessToken())
        } catch (TokenResponseException e) {
            if (e.getDetails() != null) {
                System.err.println("Error: " + e.getDetails().getError())
                if (e.getDetails().getErrorDescription() != null) {
                    System.err.println(e.getDetails().getErrorDescription())
                }
                if (e.getDetails().getErrorUri() != null) {
                    System.err.println(e.getDetails().getErrorUri())
                }
            } else {
                System.err.println(e.getMessage())
            }
        }
    }

    private static GoogleClientSecrets buildGoogleSecrets() {
        InputStream is = GoogleBuilderSpec.getResourceAsStream(CREDENTIALS_FILE_PATH)
        GoogleClientSecrets secrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(is))

        secrets.getInstalled().setClientSecret(System.getenv('DSL_SPREADSHEET_BUILDER_GOOGLE_CLIENT_SECRET'))

        return secrets
    }

    private static HttpRequestInitializer buildCredentials(GoogleClientSecrets secrets) {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport()
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, secrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build()
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build()
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user")
    }
}
