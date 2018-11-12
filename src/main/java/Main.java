import com.google.gson.JsonArray;
import com.google.gson.JsonParser;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        System.out.println("Start!");
        String srcJson = "src/Data/cucumber_reporter_07Nov.json";
        String srcExcel = "src/Data/AutomationReport.xlsx";
        String sheet = "Detail Info";
        int row = 2;

        System.out.println("Creating Classes");
        BufferedReader reader = new BufferedReader(new FileReader(srcJson));
        JsonParser parser = new JsonParser();
        JsonArray array = parser.parse(reader).getAsJsonArray();

        JSONDriven readerJS = new JSONDriven();
        ExcelTemplate template = new ExcelTemplate();

        int statusColumn = ExcelDriven.getTotalColumn(srcExcel, sheet);
        template.setHeader(srcExcel, sheet, statusColumn);

        System.out.println("Feeding data");
        for (int i = 0; i < array.size(); i++) {
            String featureName = readerJS.getFeatureName(array, i);
//            ExcelDriven.setCellValue(srcExcel, sheet, row, 1, featureName);
            int scenarios = readerJS.getTotalScenario(array, i);
            for (int x = 0; x < scenarios; x++) {
                String scenarioName = readerJS.getScenarioName(array, i, x);
//                ExcelDriven.setCellValue(srcExcel, sheet, row, 2, scenarioName);
                String result = readerJS.getScenarioStatus(array, i, x);
                ExcelDriven.setCellValue(srcExcel, sheet, row, statusColumn, result);
                row++;
            }
        }

        System.out.println("Generating data table");
        template.addToDataTable(srcExcel,sheet);

        System.out.println("Completed!");
    }
}
