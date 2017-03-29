import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.slides.v1.SlidesScopes;
import com.google.api.services.slides.v1.model.*;
import com.google.api.services.slides.v1.Slides;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import com.google.api.services.sheets.v4.Sheets;


import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.*;
import com.google.api.services.drive.Drive;
import com.google.api.services.slides.v1.model.Request;


import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.Authenticator;
import java.net.InetSocketAddress;
import java.net.PasswordAuthentication;
import java.net.Proxy;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Quickstart {
    private static final String APPLICATION_NAME = "Google Slides API Java Quickstart";

    /**
     * Directory to store user credentials for this application.
     */
    private static final java.io.File DATA_STORE_DIR = new java.io.File(
            System.getProperty("user.home"), ".credentials/slides.googleapis.com-java-quickstart");

    /**
     * Global instance of the {@link FileDataStoreFactory}.
     */
    private static FileDataStoreFactory DATA_STORE_FACTORY;

    private static final JsonFactory JSON_FACTORY =
            JacksonFactory.getDefaultInstance();

    private static HttpTransport HTTP_TRANSPORT;

    /**
     * Global instance of the scopes required by this quickstart.
     * <p>
     * If modifying these scopes, delete your previously saved credentials
     * at ~/.credentials/slides.googleapis.com-java-quickstart
     */
    private static final List<String> SCOPES =
            Arrays.asList(SlidesScopes.PRESENTATIONS_READONLY);

    static {
        try {
            //private void initializeProxyAuthenticator() {
                final String proxyUser = System.getProperty("http.proxyUser");
                final String proxyPassword = System.getProperty("http.proxyPassword");

                if (proxyUser != null && proxyPassword != null) {
                    Authenticator.setDefault(
                            new Authenticator() {
                                public PasswordAuthentication getPasswordAuthentication() {
                                    return new PasswordAuthentication(
                                            proxyUser, proxyPassword.toCharArray()
                                    );
                                }
                            }
                    );
                }
            //}


            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(1);
        }
    }


    /**
     * Creates an authorized Credential object.
     *
     * @return an authorized Credential object.
     * @throws IOException
     */
    public static Credential authorize() throws IOException {
        // Load client secrets.
        InputStream in =
                Quickstart.class.getResourceAsStream("/client_cred.json");
        GoogleClientSecrets clientSecrets =
                GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                        .setDataStoreFactory(DATA_STORE_FACTORY)
                        .setAccessType("offline")
                        .build();
        Credential credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize("user");
        System.out.println(
                "Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
        return credential;
    }

    /**
     * Build and return an authorized Slides API client service.
     *
     * @return an authorized Slides API client service
     * @throws IOException
     */
    public static Slides getSlidesService() throws IOException {
        Credential credential = authorize();
        return new Slides.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    /**
     * Build and return an authorized Sheets API client service.
     *
     * @return an authorized Sheets API client service
     * @throws IOException
     */
    public static Sheets getSheetsService() throws IOException {
        Credential credential = authorize();
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    /**
     * Build and return an authorized Drive client service.
     *
     * @return an authorized Drive client service
     * @throws IOException
     */
    public static Drive getDriveService() throws IOException {
        Credential credential = authorize();
        return new Drive.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    ////////////////////////////////////////////////////////////////

    public static void body(String templatePresentationId) throws IOException {
        Sheets sheetsService = getSheetsService();
        Drive driveService = getDriveService();
        Slides slidesService = getSlidesService();

        // https://docs.google.com/spreadsheets/d/17VZpeJqviEQVGRNZ0pw8tsjMtWcvEEqfNLb3cRZpwY8/edit
        String dataSpreadsheetId="17VZpeJqviEQVGRNZ0pw8tsjMtWcvEEqfNLb3cRZpwY8";
        String dataRangeNotation = "Sheet1!A2:H10";

        ValueRange sheetsResponse = sheetsService.spreadsheets().values()
                .get(dataSpreadsheetId, dataRangeNotation).execute();
        List<List<Object>> values = sheetsResponse.getValues();

        // For each record, create a new merged presentation.
        for (List<Object> row : values) {
            String customerName = row.get(2).toString();     // name in column 3
            String caseDescription = row.get(5).toString();  // case description in column 6
            String totalPortfolio = row.get(11).toString();  // total portfolio in column 12

            // Duplicate the template presentation using the Drive API.
            String copyTitle = customerName + " presentation";
            File content = new File().setName(copyTitle);
            File presentationFile =
                    driveService.files().copy(templatePresentationId, content).execute();
            String presentationId = presentationFile.getId();

            // Create the text merge (replaceAllText) requests for this presentation.
            // CAREFUL:  used Slides here !!!! for Request
            List<Request> requests = new ArrayList<>();
            requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest()
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText("{{customer-name}}")
                                    .setMatchCase(true))
                            .setReplaceText(customerName)));
            requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest()
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText("{{case-description}}")
                                    .setMatchCase(true))
                            .setReplaceText(caseDescription)));
            requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest()
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText("{{total-portfolio}}")
                                    .setMatchCase(true))
                            .setReplaceText(totalPortfolio)));

            // Execute the requests for this presentation.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest().setRequests(requests);
            BatchUpdatePresentationResponse response =
                    slidesService.presentations().batchUpdate(presentationId, body).execute();
        }
    }
//////////////////////////////////////////////////////////////////////////////////////


    public static void main(String[] args) throws IOException {
        // Build a new authorized API client service.
        Slides service = getSlidesService();

        // Example https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
        // https://docs.google.com/presentation/d/19eF9Oo57gzyJmwSm36BrrkGpocCSrtusBt_ckrE3hI4/edit
        String presentationId = "19eF9Oo57gzyJmwSm36BrrkGpocCSrtusBt_ckrE3hI4";
        Presentation response = service.presentations().get(presentationId).execute();
        List<Page> slides = response.getSlides();

        System.out.printf("The presentation contains %s slides:\n", slides.size());
        for (int i = 0; i < slides.size(); i++) {
            Page slide = slides.get(i);
            System.out.printf("- Slide #%s contains %s elements.\n", i + 1,
                    slide.getPageElements().size());
        }

        body(presentationId);
    }


}
