package google.app;

import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;
import java.util.Locale;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.DateTime;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

public class SheetsQuickstart
{
    private static String SHEET_NAME = "peopleData";
    private static String SPREADSHEET_ID = "1ppwFPUbfGrvDW5eKy56WLaDC1KjRRwoFBP8-Pzagu58";
    private static JsonFactory getJsonFactory()
    {
        return JacksonFactory.getDefaultInstance();
    }

    private static HttpTransport getHttpTransport()
            throws GeneralSecurityException, IOException
    {
        return GoogleNetHttpTransport.newTrustedTransport();
    }
    public static Credential getCredentials()
            throws GeneralSecurityException, IOException
    {
        File p12 = new File("src/main/resources/testtracker-374220-10e9e623c920.p12");
        System.out.println(p12.getAbsoluteFile());
        List<String> SCOPES_ARRAY =
                Arrays.asList(SheetsScopes.SPREADSHEETS);

        Credential credential = new GoogleCredential.Builder()
                .setTransport(getHttpTransport())
                .setJsonFactory(getJsonFactory())
                .setServiceAccountId("testtracker@testtracker-374220.iam.gserviceaccount.com")
                .setServiceAccountScopes(SCOPES_ARRAY)
                .setServiceAccountPrivateKeyFromP12File(p12)
                .build();

        return credential;
    }
    public static Sheets getSheetsService() throws IOException, GeneralSecurityException{
        Credential credential = getCredentials();
        return new Sheets.Builder(getHttpTransport(),
                getJsonFactory(),
                credential)
                .setApplicationName("Test Tracking")
                .build();
    }

    public static List<List<Object>> getValues(String sheetId, String sheetName, String range)
            throws GeneralSecurityException, IOException
    {
        Sheets sheets = getSheetsService();
        String fullRange = sheetName + range;

        ValueRange response = sheets.spreadsheets()
                .values()
                .get(sheetId, fullRange)
                .execute();
        return response.getValues();
    }

    public static void main(String[] args)
            throws GeneralSecurityException, IOException
    {
        LocalDate dateToday = LocalDate.now();
        DateTime dateTime = new DateTime(dateToday.atStartOfDay(ZoneOffset.UTC).toInstant().toEpochMilli());
        SimpleDateFormat dateFormatter = new SimpleDateFormat("M/d/yyyy");
        String formattedDate = dateFormatter.format(dateTime.getValue());

        if (!(args.length == 0)){
            List<List<Object>> sheetValues = getValues(SPREADSHEET_ID, SHEET_NAME, "!1:157");
            int numberOfRowToUpdate = 1;
            boolean rowIsNotEmpty;

            for (List<Object> row : sheetValues){
                try {
                    row.get(0);
                    System.out.println(row);
                    rowIsNotEmpty = true;
                }catch (Exception e){
                    rowIsNotEmpty = false;
                }
                if (rowIsNotEmpty){
                    if (row.get(0).equals(formattedDate)){
                        break;
                    }
                }
                numberOfRowToUpdate++;
            }

            String rangeOfInserting = "!" + args[0] + numberOfRowToUpdate + ":" + args[1] + numberOfRowToUpdate;
            System.out.println(rangeOfInserting);
            if (getValues(SPREADSHEET_ID, SHEET_NAME, "!A" + numberOfRowToUpdate) == null){
                ValueRange appendBody = new ValueRange().setValues(Arrays.asList(Arrays.asList(formattedDate)));
                AppendValuesResponse appendResult = getSheetsService().spreadsheets().values()
                        .append(SPREADSHEET_ID,"A" + numberOfRowToUpdate , appendBody)
                        .setValueInputOption("USER_ENTERED")
//                        .setInsertDataOption("INSERT_ROWS")
                        .setIncludeValuesInResponse(true)
                        .execute();
            }
            if (getValues(SPREADSHEET_ID, SHEET_NAME, rangeOfInserting) == null){
                String[] lastValuesToInsert = Arrays.copyOfRange(args, 3, args.length);
                ValueRange appendBody = new ValueRange().setValues(Arrays.asList(Arrays.asList(args[2], Arrays.toString(lastValuesToInsert))));
                UpdateValuesResponse appendResult = getSheetsService().spreadsheets().values()
                        .update(SPREADSHEET_ID, rangeOfInserting, appendBody)
                        .setValueInputOption("RAW")
                        .execute();
            }
        }
        else {
            for (List<Object> row : getValues(SPREADSHEET_ID, SHEET_NAME, "!1:157")) {
                boolean valueIsThere = true;

                try {
                    row.get(1);
                } catch (Exception e) {
                    valueIsThere = false;
                }
                if (valueIsThere) {
                    System.out.println("Date: " + row.get(0));
                    System.out.println("Pass/Fail: " + row.get(1));
                    if (row.get(1).toString().equalsIgnoreCase("failed")) {
                        if (row.size() > 2) {
                            System.out.println("Failure Message: " + row.get(2));
                        }
                    }
                }

                System.out.println("------------------------------------");
            }


        }
        System.out.println(getValues(SPREADSHEET_ID, SHEET_NAME, "!A65:C65"));
    }
}
