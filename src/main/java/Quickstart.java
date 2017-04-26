import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.api.services.slides.v1.Slides;
import com.google.api.services.slides.v1.SlidesScopes;
import com.google.api.services.slides.v1.model.*;
import gds.brnvrn.com.Services;
import link.brnvrn.com.Linko;

import java.io.IOException;
import java.net.SocketTimeoutException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;


/**
 *
 *
Not that easy to create one slide per row in a spreadsheet.
The example gives one presentation (=file) per row !!! https://developers.google.com/slides/how-tos/merge
- Problem is that the replaceText act on everything (even on the layouts in the Master slide)
  There is no way to reduce the scope (or the context) of the "replace all text" action!
- I did not find a way to get the whole page to recreate it later.
- I did not find a way to get the ObjectId from the Web (GUI) interface

So I get the objectId of the textElements to replace
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
            SheetsScopes.SPREADSHEETS_READONLY);


    private static void testArgs(String[] args) {
        String msg = Quickstart.class.getSimpleName()+ " PresentationId SpreadsheetId DataRange \n";
        msg += "For example: " + Quickstart.class.getSimpleName()+ " 19eF9Oo57gzyJmwSm36BrrkGpocCSrtusBt_ckrE3hI4 17VZpeJqviEQVGRNZ0pw8tsjMtWcvEEqfNLb3cRZpwY8 \"Sheet1!A1:H30\" \n";
        msg += " Use Ids as they appear in the urls: https://docs.google.com/presentation/d/19eF9Oo57gzyJmwSm36BrrkGpocCSrtusBt_ckrE3hI4/edit\n";
        if (args.length != 3) {
            System.out.println("Error: "+args[0]);
            System.err.println("Usage: " + msg);
            System.exit(1);
        }
    }

    public static void main(String[] args) throws IOException, GeneralSecurityException {

        testArgs(args);

        Services services = new Services(APPLICATION_NAME, SCOPES);
        // Build a new authorized API client service.
        Slides slides = services.getSlidesService();
        Sheets sheets = services.getSheetsService();

        // Example https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
        // https://docs.google.com/presentation/d/19eF9Oo57gzyJmwSm36BrrkGpocCSrtusBt_ckrE3hI4/edit
        String presentationId = args[0]; // "19eF9Oo57gzyJmwSm36BrrkGpocCSrtusBt_ckrE3hI4";
        Presentation presentation = slides.presentations().get(presentationId).execute();
        List<Page> pages = presentation.getSlides();

        // https://docs.google.com/spreadsheets/d/17VZpeJqviEQVGRNZ0pw8tsjMtWcvEEqfNLb3cRZpwY8/edit
        String dataSpreadsheetId = args[1]; // "17VZpeJqviEQVGRNZ0pw8tsjMtWcvEEqfNLb3cRZpwY8";
        String dataRangeNotation = args[2]; // "Sheet1!A1:H30";

        ValueRange sheetsResponse = sheets.spreadsheets().values().get(dataSpreadsheetId, dataRangeNotation).execute();
        List<List<Object>> values = sheetsResponse.getValues();

        Linko linko = new Linko();
        linko.createBridge(values.get(0));

        String slideLastId = pages.get(0).getObjectId();
        Page lastPage = pages.get(0);
        for (PageElement pae : lastPage.getPageElements()) {
            String objectId = pae.getObjectId();
            System.out.println(" Oid: " + objectId);
            TextContent textContent = pae.getShape().getText();
            if (textContent == null) continue;
            for (TextElement te : textContent.getTextElements()) {
                if (te.getTextRun() != null) {
                    String content = te.getTextRun().getContent();
                    if (linko.setMatch(content, objectId))
                        System.out.println("Matching oi:" + objectId + " with content: " + content + ".");
                    else
                        System.out.println("No Match for oi:" + objectId + " with content: " + content + " !!");
                }
            }
        }

        int i = -1;
        for (List<Object> row : values) {

            if (i==-1) {             // We skip the headers
                i=0;
                continue;
            }

            List<Request> requests = new ArrayList<>();
            requests.add(new Request()
                .setDuplicateObject(new DuplicateObjectRequest()
                    .setObjectId(slideLastId)
                    .setObjectIds(linko.getMapId(i))));
            if (i==0) {
                for (String field: linko.getFields()) {
                    requests.add(new Request()
                            .setReplaceAllText(new ReplaceAllTextRequest()
                                    .setContainsText(new SubstringMatchCriteria()
                                            .setText(field)
                                            .setMatchCase(true))
                                    .setReplaceText("")));
                }
            }
            for (Map.Entry<String, Integer> map: linko.getMapObj2Col(i)) {
                requests.add(new Request()
                        .setInsertText(new InsertTextRequest()
                                .setObjectId(map.getKey())
                                .setText(row.get(map.getValue()).toString())));
            }

            // Execute the requests for this presentation.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest().setRequests(requests);
            System.out.println("Body #"+ i +" : "+body.toPrettyString() + "\n");
            BatchUpdatePresentationResponse response = null;
            int retry=2;
            while (response == null && ((retry--) > 0))
                try {
                    response = slides.presentations().batchUpdate(presentationId, body).execute();
                } catch (SocketTimeoutException e) {
                    System.out.println("SocketTimeoutException #" + retry + " !!");
                }
            slideLastId = response.getReplies().get(0).getDuplicateObject().getObjectId();
            i++;
        }
    }

}
