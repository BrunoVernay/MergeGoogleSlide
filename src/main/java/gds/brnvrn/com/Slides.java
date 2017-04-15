package gds.brnvrn.com;

import com.google.api.client.auth.oauth2.Credential;

import java.io.IOException;

/**
 * Created by bruno on 4/12/17.
 */
public class Slides {

    private com.google.api.services.slides.v1.Slides slidesService;

    public Slides(Services services) {
        if (slidesService == null) {
            try {
                slidesService = services.getSlidesService();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }



}
