import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

public class GradesSheetsTunts {
    private static final String APPLICATION_NAME = "Grades Sheets Tunts";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.DRIVE);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";
    final static String SHEET_ID = "1cDuVhJY8_5IgfoOlTk9TwMG5K91ZQ2WomMRXrZNmk74";

    public static void main(String... args) throws IOException, GeneralSecurityException {
        final List<List<Object>> values = getRangeAsList("!A4:F27");

        if (values == null || values.isEmpty()) {
            System.out.println("No data found.");
            return;
        }

        final List<List<Object>> validatedValues = validadeGrades(values);
        updateSheet(validatedValues);
        for (List row : validatedValues) {
            System.out.printf("Registration: %s | Student: %s | Situation: %s\n", row.get(0) ,row.get(1), row.get(6));
        }
    }

    public static void updateSheet(final List<List<Object>> values) throws IOException, GeneralSecurityException {
        final Sheets service = getService();
        ValueRange valueRange = new ValueRange()
                .setValues(values);

        UpdateValuesResponse response = service.spreadsheets().values()
                .update(SHEET_ID, "A4:H27", valueRange)
                .setValueInputOption("RAW")
                .execute();
        System.out.println(response);
    }

    public static void getFinalGrade(final String situation, final List row, final Double average) {
        if (situation != "Final exam") {
            row.add(7, "0");
            return;
        }

        // I think I don't understand how to use the formula
        // I hope it's correct. If it's not I hope I can learn if it's possible :)
        double result = (100 - average);
        row.add(7, Math.ceil(result));
    }

    public static List<List<Object>> validadeGrades(final List<List<Object>> values) {
        for (List row : values) {
            final Double p1 = Double.valueOf(row.get(3).toString());
            final Double p2 = Double.valueOf(row.get(4).toString());
            final Double p3 = Double.valueOf(row.get(5).toString());
            final Double average = (p1 + p2 + p3) / 3;

            final Integer absences = Integer.parseInt(row.get(2).toString());
            final Integer absencesToReprove = (60 * 25) / 100;

            if (absences > absencesToReprove){
                row.add(6, "Disapproved by absences");
                getFinalGrade("Disapproved by absences", row, average);
                continue;
            }

            if (average < 50)
                row.add(6, "Disapproved by grade");
            else if (average == 50 || average < 70)
                row.add(6, "Final exam");
            else
                row.add(6, "Approved");

            final String situation = row.get(6).toString();
            getFinalGrade(situation, row, average);
        }

        return values;
    }

    public static List<List<Object>> getRangeAsList(final String range) throws IOException, GeneralSecurityException {
        final Sheets service = getService();
        ValueRange response = service.spreadsheets().values()
                .get(SHEET_ID, range)
                .execute();
        return response.getValues();
    }

    public static Sheets getService() throws IOException, GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        InputStream in = GradesSheetsTunts.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }
}