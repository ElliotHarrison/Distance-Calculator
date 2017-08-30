import com.google.maps.DirectionsApi;
import com.google.maps.DistanceMatrixApi;
import com.google.maps.GeoApiContext;
import com.google.maps.model.DistanceMatrix;
import com.google.maps.model.TravelMode;
import com.google.maps.model.Unit;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.joda.time.DateTime;
import java.io.*;
import org.json.*;
import java.util.*;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

/**
 * Class Distances implements and operates on an instance of the GeoApiContext class,
 * DistanceMatrix and Directions google Apis, by requesting the Google server using
 * a unique Api key to authenticate. The class reads in excel data into a Distance Matrix, transmits and computes
 * via the DistanceMatrixApi and returns a JSON Object through a GSON object to be parsed and formatted into a String.
 */

public class Distances 
{

  private static GeoApiContext context = new GeoApiContext.Builder().apiKey("Api key").build();
  private String[] excelData;

  /**
   * Method returnTownsCities creates an File stream input of the excel data file
   * stores the data in a NPOI file system and HSSF workbook/sheet, the HSSF class row
   * is used in iterating through the rows of the table, the first column is iterated through retrieved
   * and returned as @return String[] excelData.
   */
  private String[] returnTownsCities() {
        
    try {

        FileInputStream fileIn = new FileInputStream("PATH\TO\FILE\test_three.xls");
        NPOIFSFileSystem fs = new NPOIFSFileSystem(fileIn);
        HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row;

        int rows;
        rows = sheet.getPhysicalNumberOfRows();
        int cols = 0;
        int tmp = 0;

        for (int i = 0; i < 10 || i < rows; i++) {
          row = sheet.getRow(i);
          if (row != null) {
            tmp = sheet.getRow(i).getPhysicalNumberOfCells();
            if (tmp > cols) cols = tmp; {
            }
          }
        }
            
        excelData = new String[sheet.getLastRowNum()];
        for (int j = 0; j < sheet.getLastRowNum(); j++) {
          row = sheet.getRow(j);
          Cell celRow = row.getCell(0);
          String celVal = cel.getStringCellValue();
          excelData[j] = celVal;

        }

        } catch (IOException e) {
          System.out.println(e.getMessage());
      } return excelData;
    }

    /**
     * Main Method initializes an object instance of the Distances class.
     * Stores the returned String[] value from the returnTownsCities method,
     * to select a town/city at random. a Distance Matrix input request is made to the Api destination.
     * an origin and a destination are read into the Distance Matrix along with the DistanceMatrixApi parameter definitions.
     * Receives a JSON Array formatted through Gson. the time taken from origin and destination is extracted through iteration of
     * the JSON Array to be printed.
     * @param args
     */
    public static void main(String[] args) {
          
      try {

          Distances distObject = new Distances();
          String[] extractString = distObject.returnTownsCities();

          String originA = extractString[new Random().nextInt(extractString.length)];
          String destinationA = extractString[new Random().nextInt(extractString.length)];

          DistanceMatrix matrix =
                  DistanceMatrixApi.newRequest(context)
                          .origins(originA)
                          .destinations(destinationA)
                          .mode(TravelMode.WALKING)
                          .language("en-AU")
                          .avoid(DirectionsApi.RouteRestriction.TOLLS)
                          .units(Unit.METRIC)
                          .departureTime(
                                  new DateTime().plusMinutes(2))
                          .await();

          Gson gson = new GsonBuilder().setPrettyPrinting().create();
          String jsonInString = gson.toJson(matrix);
          System.out.println(jsonInString);

          JSONObject jsonObject = new JSONObject(jsonInString = gson.toJson(matrix));
          JSONArray rows = jsonObject.getJSONArray("rows");

          for(int i = 0; i < rows.length(); i++) {
            JSONObject one = rows.getJSONObject(i);
            JSONArray element = one.getJSONArray("elements");
            for(int j = 0; j < element.length(); j++) {
              JSONObject duration = element.getJSONObject(j).getJSONObject("duration");
              JSONObject distance = element.getJSONObject(j).getJSONObject("distance");
              System.out.println("It will take " + duration.getString("humanReadable") + " to walk from " + originA + " to " + destinationA);
          }
        }

      } catch (Exception e) {
          System.out.println(e.getMessage());
    }
  }   
}   
