import java.io.IOException;
import java.util.Date;

public class ExcelTemplate {

    public void setHeader(String src, String sheet, int column) throws IOException {
        String reportDate = ExcelDriven.getCellValue(src, "Dashboard", 1, 1);

        ExcelDriven.setCellValue(src, sheet, 0, column, "Date:");
        ExcelDriven.setCellValue(src, sheet, 0, column+1, reportDate);
        ExcelDriven.setCellValue(src, sheet, 1, column, "Status");
        ExcelDriven.setCellValue(src, sheet, 1, column+1, "Note");

        int totalRow = ExcelDriven.getTotalRow(src, sheet);
        for (int i = 2; i <= totalRow; i++){
            ExcelDriven.setCellValue(src, sheet, i, column, "Not Run on Sauce Lab");
        }
    }

    public void addToDataTable(String src, String sheet) throws IOException{
        String reportName = "[" + ExcelDriven.getCellValue(src, "Dashboard", 0, 1)
                + "] " + ExcelDriven.getCellValue(src, "Dashboard", 1, 1);

        for (int i=1; i<=10; i++){
            if (ExcelDriven.getCellValue(src, "Chart Data", i, 0) == ""){
                ExcelDriven.setCellValue(src, "Chart Data", i, 0, reportName);

                ExcelDriven.setCellFormula(src, "Chart Data", i, 1, "COUNTIF('Detail Info'!D:D, \"Passed on Sauce Lab\")");
                ExcelDriven.setCellFormula(src, "Chart Data", i, 2, "COUNTIF('Detail Info'!D:D, \"Passed After Re-Running\")");
                ExcelDriven.setCellFormula(src, "Chart Data", i, 3, "COUNTIF('Detail Info'!D:D, \"Failed By Bug\")");
                ExcelDriven.setCellFormula(src, "Chart Data", i, 4, "COUNTIF('Detail Info'!D:D, \"Failed By Script\")");
                ExcelDriven.setCellFormula(src, "Chart Data", i, 5, "COUNTIF('Detail Info'!D:D, \"Not Run on Sauce Lab\")");
                ExcelDriven.setCellFormula(src, "Chart Data", i, 6, "COUNTA('Detail Info'!C2:C1000)");

                int j = i+1;
                ExcelDriven.setCellFormula(src, "Chart Data", i, 8, "B"+j+"/G"+j);
                ExcelDriven.setCellFormula(src, "Chart Data", i, 9, "C"+j+"/G"+j);
                ExcelDriven.setCellFormula(src, "Chart Data", i, 10, "D"+j+"/G"+j);
                ExcelDriven.setCellFormula(src, "Chart Data", i, 11, "E"+j+"/G"+j);
                ExcelDriven.setCellFormula(src, "Chart Data", i, 12, "F"+j+"/G"+j);
                ExcelDriven.setCellFormula(src, "Chart Data", i, 13, "SUM(I"+j+":M"+j+")");

                return;
            }
        }



    }
}
