import com.google.gson.JsonArray;
import com.google.gson.JsonParser;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        System.out.println("Start generating automation report");
        //Prepare input data
        String srcJson = "src/Data/cucumber_reporter.html.json";
        String srcExcel = "src/Data/AutomationReport.xlsx";
        String sheetDetailInfo = "Detail Info", sheetDashboard = "Dashboard", sheetDataTable = "Data Table",
                reportDate = ExcelDriven.getCellValue(srcExcel, sheetDashboard, 1, 1),
                sprint = ExcelDriven.getCellValue(srcExcel, sheetDashboard, 0, 1),
                reportName = "[" + sprint + "] " + reportDate;
        int colFeature = 1, colScenario =2;



        System.out.println("Prepare precondition");
        //read the input JSON file THEN parse to JSON Array
        BufferedReader reader = new BufferedReader(new FileReader(srcJson));
        JsonParser parser = new JsonParser();
        JsonArray array = parser.parse(reader).getAsJsonArray();
        //prepare the related classes
        JSONDriven readerJS = new JSONDriven();
        ExcelTemplate template = new ExcelTemplate();
        //create the report header (Report date, status, note) in the Detail Info sheet
//        int statusColumn = ExcelDriven.getTotalColumn(srcExcel, sheetDetailInfo);
        int statusColumn = 3;
        template.setHeader(srcExcel, sheetDetailInfo, statusColumn, reportDate);

        System.out.println("Feed the cucumber data into sheet");
        //define the starting row
        int startRow = 2;
        //feed data to sheet
        for (int i = 0; i < array.size(); i++) {
            //add feature name into sheet
            String featureName = readerJS.getFeatureName(array, i);
            ExcelDriven.setCellValue(srcExcel, sheetDetailInfo, startRow, colFeature, featureName);
            //add all scenarios of feature into sheet
            for (int x = 0; x < readerJS.getTotalScenario(array, i); x++) {
                String scenarioName = readerJS.getScenarioName(array, i, x);
                ExcelDriven.setCellValue(srcExcel, sheetDetailInfo, startRow, colScenario, scenarioName);
                String result = readerJS.getScenarioStatus(array, i, x);
                ExcelDriven.setCellValue(srcExcel, sheetDetailInfo, startRow, statusColumn, result);
                startRow++;
            }
        }

        System.out.println("Generating data table");
        template.addToDataTable(srcExcel,sheetDataTable, reportName);


        System.out.println("Finish generating automation report");
    }
}
