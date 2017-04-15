import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.api.services.slides.v1.Slides;
import com.google.api.services.slides.v1.SlidesScopes;
import com.google.api.services.slides.v1.model.*;
import com.google.api.services.slides.v1.model.DuplicateObjectRequest;
import gds.brnvrn.com.Services;

import java.io.IOException;
import java.net.Authenticator;
import java.net.PasswordAuthentication;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static java.lang.System.exit;


/*
Not that easy to create one slide per row in a spreadsheet.
The example gives one presentation (=file) per row !!!
- Problem is that the replaceText act on everything (even on the layouts in the Master slide)
  There is no way to reduce the scope (or the context) of the "replace all text" action!
- I did not find a way to get the whole page to recreate it later.
- I did not find a way to get the ObjectId from the Web (GUI) interface
So I am trying to get the objectId of the textElements to replace
Loop:
  Duplicate the page, fixing the new Objects Id
  Changing the text in the first page
 */
public class Quickstart {
    private static final String APPLICATION_NAME = "Google Slides API Java Quickstart";

    /**
     * Global instance of the scopes required by this quickstart.
     * <p>
     * If modifying these scopes, delete your previously saved credentials
     * at ~/.credentials/slides.googleapis.com-java-quickstart
     */
    private static final List<String> SCOPES = Arrays.asList(
                    SlidesScopes.PRESENTATIONS,
                    SheetsScopes.SPREADSHEETS_READONLY,
                    DriveScopes.DRIVE_FILE);

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

        } catch (Throwable t) {
            t.printStackTrace();
            exit(1);
        }
    }


    public static void main(String[] args) throws IOException, GeneralSecurityException {
        Services services = new Services(APPLICATION_NAME, SCOPES);
        // Build a new authorized API client service.
        Slides slides = services.getSlidesService();
        Sheets sheets = services.getSheetsService();
        Drive drive = services.getDriveService();

        // Example https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
        // https://docs.google.com/presentation/d/19eF9Oo57gzyJmwSm36BrrkGpocCSrtusBt_ckrE3hI4/edit
        String presentationId = "19eF9Oo57gzyJmwSm36BrrkGpocCSrtusBt_ckrE3hI4";
        Presentation presentation = slides.presentations().get(presentationId).execute();
        List<Page> pages = presentation.getSlides();
        System.out.printf("The presentation contains %s slides:\n", pages.size());

        // https://docs.google.com/spreadsheets/d/17VZpeJqviEQVGRNZ0pw8tsjMtWcvEEqfNLb3cRZpwY8/edit
        String dataSpreadsheetId="17VZpeJqviEQVGRNZ0pw8tsjMtWcvEEqfNLb3cRZpwY8";
        String dataRangeNotation = "Sheet1!A2:H10";

        ValueRange sheetsResponse = sheets.spreadsheets().values().get(dataSpreadsheetId, dataRangeNotation).execute();
        List<List<Object>> values = sheetsResponse.getValues();

        //String slideLastId = pages.get(0).getObjectId();
        Page lastPage = pages.get(0);
        for (PageElement pae : lastPage.getPageElements()) {
            String objectId = pae.getObjectId();
            System.out.println(" Oid: "+objectId);
            System.out.println(" type: "+pae.getShape().getShapeType());
            TextContent textContent = pae.getShape().getText();
            if (textContent==null) continue;
            for (TextElement te : textContent.getTextElements() ) {
                if (te.getTextRun() != null)
                    System.out.println("tr: "+te.getTextRun().getContent());
            }
        }

        // Create the map of ObjectId -> name

        exit(0);

        // For each record, create a new merged presentation.
        for (List<Object> row : values) {

            String startupName = row.get(1).toString();
            String startupDescription = row.get(4).toString();
            String startupURL = row.get(2).toString();

            // i = 0
            // Duplicate the last slide, mapping old ObjectId to ObjectId-BVE-i
            // Change Old ObjectIds text in the original slide
            // i++

            List<Request> requests = new ArrayList<>();
            requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest()
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText("{{customer-name}}")
                                    .setMatchCase(true))
                            .setReplaceText(startupName)));
            requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest()
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText("{{case-description}}")
                                    .setMatchCase(true))
                            .setReplaceText(startupDescription)));
            requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest()
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText("{{total-portfolio}}")
                                    .setMatchCase(true))
                            .setReplaceText(startupURL)));
            requests.add(new Request()
                    .setCreateSlide(new CreateSlideRequest()));     // Could select a layout

            // Execute the requests for this presentation.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest().setRequests(requests);
            BatchUpdatePresentationResponse response2 =
                    slides.presentations().batchUpdate(presentationId, body).execute();
            //slideLastId = response2.getReplies().get(3).getDuplicateObject().getObjectId();
        }
    }


}
