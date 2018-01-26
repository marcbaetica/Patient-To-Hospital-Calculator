package dataCompute;

//dependencies: poi-ooxml-3.15.jar and others (see lib)
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class distanceCompute {

	public static void main(String[] args) {
		
		//TODO: obviously need to make this from the root folder
		String path = "C:\\Users\\marcb\\Desktop\\Work Tasks\\Patient-To-Hospital-Calculator\\Data\\DistancesTime_after_processing.xlsx";
		XSSFWorkbook excelWBook;
		XSSFSheet excelWSheet;
		String sheetName = "Sheet1";
		String source, destination;
		
		String URL_start = "https://maps.googleapis.com/maps/api/distancematrix/json?origins=";
		String URL_middle = "&destinations=";
		String URL_end = "&key=AIzaSyCycHp1eVA-qsW80P7KV5MRhBmt1gffEBQ";
		
		String distance;
		String time;
		
		int hours, minutes;
		
		try {
			
			FileInputStream excelFile = new FileInputStream(path);
			excelWBook = new XSSFWorkbook(excelFile);
			excelWSheet = excelWBook.getSheet(sheetName);

			//for cc1=2, for cc2=5, for cc3=8, for cc4=11, for cc5=14, for cc6=17
			int columnSelected = 5; //TODO: automate more. good for now flushing this out as such: row max=number left (-1) and here column=theOneWIthHospital
			
			//for (int rowSelected=1270; rowSelected<=1590; rowSelected++) { - for bottom row
			for (int rowSelected=452; rowSelected<=455; rowSelected++) {
			
				//to output data from tables
				//excelCell = excelWSheet.getRow(2).getCell(1);
				//cellData = excelCell.getStringCellValue();
				source = returnZIPDataFromCell (rowSelected, 1, excelWSheet);
				//System.out.println("Excell cell 1,1 (patient) has cell data: " + source);
				destination = returnZIPDataFromCell (rowSelected, columnSelected, excelWSheet);
				//System.out.println("Excell cell 1,2 (hospital) has cell data: " + destination);
				//String cellData = returnZIPDataFromCell (70, 1, excelWSheet);
				//System.out.println("Excell cell 1,2 (random cell data) has cell data: " + cellData);
				
				if (!source.equals("") && !destination.equals("")) { //work only if the two cells are not empty
					
					//assume that at this point we get source and destination above
					//make URL happen
					String URL_full = URL_start + source + URL_middle + destination + URL_end;
					//System.out.println("Full URL for REST call: " + URL_full);
					
					//make REST call
					// Connect to the URL using java's native library
				    URL url = new URL(URL_full);
				    HttpURLConnection request = (HttpURLConnection) url.openConnection();
				    request.connect();
		
				    // Convert to a JSON object to print data
				    JsonParser jp = new JsonParser(); //from gson
				    JsonElement root = jp.parse(new InputStreamReader((InputStream) request.getContent())); //Convert the input stream to a json element
				    JsonObject rootobj = root.getAsJsonObject(); //May be an array, may be an object. 
					
				    //if postal codes missing it returns an empty array (by design so we avoid extra if statements) but this means we need to parse the string manually later
				    String responseJSON = rootobj.get("rows").getAsJsonArray().toString();
				    //System.out.println(responseJSON);
				    
				    //parse for distance and time
				    //if returned status is either [{elements":[{"status":"ZERO_RESULTS"}]}] or [{"elements":[{"status":"NOT_FOUND"}]}] or [] (last one if input is empty)
				    if (responseJSON.contains("[{\"elements\":[{\"status\":\"") || responseJSON.contains("[]")) { 
				    	System.out.println("");
				    } else {
	
					    distance = responseJSON.substring(responseJSON.indexOf("distance\":{\"text\":\"") + 19, responseJSON.indexOf(" km\",\"value\":"));
					    if (responseJSON.contains("mins")) {
					    	time = responseJSON.substring(responseJSON.indexOf("duration\":{\"text\":\"") + 19, responseJSON.indexOf(" mins\",\"value\":"));
					    } else {
					    	time = responseJSON.substring(responseJSON.indexOf("duration\":{\"text\":\"") + 19, responseJSON.indexOf(" min\",\"value\":"));
					    }
					    //see if it is in km to add them at end
					    //if (responseJSON.contains("km")) distance+=" km";
					    if (responseJSON.contains("min")) time+=" min";
					    //System.out.println(distance + " " + time);
					    
					    //custom data formatting (insert meters if not km distance, make time conversion to minutes)
					    if (!responseJSON.contains("km")) distance+=" meters";
					    System.out.println(distance);
					    
					    //Print Time
					    //System.out.println(time + ".");
					    //check if hours exists
				    	if (time.contains("hours")) {//break basted on time
				    		String hoursString = time.substring(0, time.indexOf(" hours"));
				    		//System.out.println("hoursString: " + hoursString);//it ends with this?
				    		String minutesString = time.substring(time.indexOf(" hours ") +7, time.indexOf(" min"));
					    	
				    		//System.out.print("For strings we have: " + hoursString + ". ." + minutesString);
				    		hours = Integer.parseInt(hoursString);
				    		minutes = Integer.parseInt(minutesString);
				    		//System.out.print("For strings we have: " + hoursString + ". ." + minutesString);
				    		
				    		minutes = hours*60+minutes;
				    	} else if (time.contains("hour")) {//break basted on time
				    		String hoursString = time.substring(0, time.indexOf(" hour"));
				    		//System.out.println("hoursString: " + hoursString);//it ends with this?
				    		String minutesString = time.substring(time.indexOf(" hour ") +6, time.indexOf(" min"));
					    	
				    		//System.out.print("For strings we have: " + hoursString + ". ." + minutesString);
				    		hours = Integer.parseInt(hoursString);
				    		minutes = Integer.parseInt(minutesString);
				    		//System.out.print("For strings we have: " + hoursString + ". ." + minutesString);
				    		
				    		minutes = hours*60+minutes;
				    	} else {//if only minutes exist
				    		String minutesString = time.substring(0, time.indexOf(" min"));
				    		minutes = Integer.parseInt(minutesString);
				    	}
			    		System.out.println(minutes);
				    }
				    
				    
				    //add details in table
				    //addRestDataToTable (rowSelected, columnSelected, excelWSheet, distance, time); //TODO: automate more
				    
				    //double check data
				    //String distanceFromTable = returnZIPDataFromCell (rowSelected, columnSelected+1, excelWSheet);
					//System.out.println("Excell cell has cell data: " + distanceFromTable);
					//String timeFromTable = returnZIPDataFromCell (rowSelected, columnSelected+2, excelWSheet);
					//System.out.println("Excell cell has cell data: " + timeFromTable);
			    
				} else {
					System.out.println("EMPTY INPUT");
				}
			
			    //flush info at the end
			    distance = "";
			    time = "";
			    
			}
			
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		

	}
	
	public static String returnZIPDataFromCell (int row, int column, XSSFSheet excelWSheet) {
		XSSFCell excelCell = excelWSheet.getRow(row).getCell(column);
		String cellData = excelCell.getStringCellValue();
		
		//replacing the spaces with + so they can be used for the API
		if (cellData.contains(" ")) {
			cellData = cellData.replace(" ", "+");
		}
		
		return cellData;
	}
	
	public static void addRestDataToTable (int row, int column, XSSFSheet excelWSheet, String distance, String time) {
		//distance
		XSSFCell excelCell = excelWSheet.getRow(row).createCell(column+1);
		excelCell.setCellValue(distance);
		//time
		excelCell = excelWSheet.getRow(row).createCell(column+2);
		excelCell.setCellValue(time);
	}

}
