import java.io.IOException;

public class ExcelTemplate {

    public void setHeader(String src, String DetailInfo, int column, String reportDate) throws IOException {
        //create a template for the new report
        ExcelDriven.setCellValue(src, DetailInfo, 0, column, "Date:");
        ExcelDriven.setCellValue(src, DetailInfo, 0, column+1, reportDate);
        ExcelDriven.setCellValue(src, DetailInfo, 1, column, "Status");
        ExcelDriven.setCellValue(src, DetailInfo, 1, column+1, "Note");
        //create a default value
        for (int i = 2; i <= ExcelDriven.getTotalRow(src, DetailInfo); i++){
            ExcelDriven.setCellValue(src, DetailInfo, i, column, "Not Run on Sauce Lab");
        }
    }

    public void addToDataTable(String src, String DataTable, String reportName) throws IOException{

        for (int i=1; i<=10; i++){
            if (ExcelDriven.getCellValue(src, DataTable, i, 0) == "" || ExcelDriven.getCellValue(src, DataTable, i, 0) == null){
                //add report date to data table
                ExcelDriven.setCellValue(src, DataTable, i, 0, reportName);

                //feed data from Detail Info sheet
                ExcelDriven.setCellFormula(src, DataTable, i, 1, "COUNTIF('Detail Info'!D:D, \"Passed on Sauce Lab\")");
                ExcelDriven.setCellFormula(src, DataTable, i, 2, "COUNTIF('Detail Info'!D:D, \"Passed After Re-Running\")");
                ExcelDriven.setCellFormula(src, DataTable, i, 3, "COUNTIF('Detail Info'!D:D, \"Failed By Bug\")");
                ExcelDriven.setCellFormula(src, DataTable, i, 4, "COUNTIF('Detail Info'!D:D, \"Failed By Script\")");
                ExcelDriven.setCellFormula(src, DataTable, i, 5, "COUNTIF('Detail Info'!D:D, \"Not Run on Sauce Lab\")");
                ExcelDriven.setCellFormula(src, DataTable, i, 6, "COUNTA('Detail Info'!C:C)-2");

                //generate the percentage number
                int row = i+1;
                ExcelDriven.setCellFormula(src, DataTable, i, 8, "B"+row+"/G"+row);
                ExcelDriven.setCellFormula(src, DataTable, i, 9, "C"+row+"/G"+row);
                ExcelDriven.setCellFormula(src, DataTable, i, 10, "D"+row+"/G"+row);
                ExcelDriven.setCellFormula(src, DataTable, i, 11, "E"+row+"/G"+row);
                ExcelDriven.setCellFormula(src, DataTable, i, 12, "F"+row+"/G"+row);
                ExcelDriven.setCellFormula(src, DataTable, i, 13, "SUM(I"+row+":M"+row+")");

                return;
            }
        }



    }
}
