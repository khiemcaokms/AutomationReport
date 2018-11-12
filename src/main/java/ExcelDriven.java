import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelDriven {

    public static String getCellValue(String src, String sheetName, int row, int column) {
        try {
            FileInputStream srcFile = new FileInputStream(src);
            XSSFWorkbook wb = new XSSFWorkbook(srcFile);
            XSSFSheet sheet = wb.getSheet(sheetName);
            XSSFCell cell = sheet.getRow(row).getCell(column);
            if (cell != null) {
                String value = cell.getStringCellValue();
                return value;
            }
            return null;
        } catch (IOException e) {
            System.out.println(e.getMessage());
            return null;
        }
    }

    public static void createRow(String src, String sheetName, int row){
        try {
            FileInputStream srcFile = new FileInputStream(src);
            XSSFWorkbook wb = new XSSFWorkbook(srcFile);
            XSSFSheet sheet = wb.getSheet(sheetName);
            sheet.createRow(row);
            FileOutputStream outfile = new FileOutputStream(src);
            wb.write(outfile);
            outfile.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    public static void setCellValue(String src, String sheetName, int row, int column, String newValue) throws IOException{
        try {
            FileInputStream srcFile = new FileInputStream(src);
            XSSFWorkbook wb = new XSSFWorkbook(srcFile);
            XSSFSheet sheet = wb.getSheet(sheetName);
            XSSFCell cell = sheet.getRow(row).getCell(column);
            if (cell == null)
                cell = sheet.getRow(row).createCell(column);
            cell.setCellValue(newValue);
            FileOutputStream outfile = new FileOutputStream(src);
            wb.write(outfile);
            outfile.close();
        } catch (NullPointerException e) {
            createRow(src, sheetName, row);
            setCellValue(src, sheetName, row, column, newValue);
        }
    }
    public static void setCellFormula(String src, String sheetName, int row, int column, String formula) throws IOException{
        try {
            FileInputStream srcFile = new FileInputStream(src);
            XSSFWorkbook wb = new XSSFWorkbook(srcFile);
            XSSFSheet sheet = wb.getSheet(sheetName);
            XSSFCell cell = sheet.getRow(row).getCell(column);
            if (cell == null)
                cell = sheet.getRow(row).createCell(column);
            cell.setCellFormula(formula);
            FileOutputStream outfile = new FileOutputStream(src);
            wb.write(outfile);
            outfile.close();
        } catch (NullPointerException e) {
            createRow(src, sheetName, row);
            setCellValue(src, sheetName, row, column, formula);
        }
    }

    public static int getTotalRow(String src, String sheetName) {
        try {
            FileInputStream file = new FileInputStream(src);
            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheet(sheetName);
            int totalRow = sheet.getLastRowNum();
            return totalRow;
        } catch (IOException e) {
            System.out.println(e.getMessage());
            return -1;
        }
    }
    public static int getTotalColumn(String src, String sheetName) {
        try {
            FileInputStream file = new FileInputStream(src);
            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheet(sheetName);
            int totalColumn = sheet.getRow(0).getLastCellNum();
            return totalColumn;
        } catch (IOException e) {
            System.out.println(e.getMessage());
            return -1;
        }
    }

    public static boolean isCellNull(String src, String sheetName, int row, int col){
        try {
            FileInputStream file = new FileInputStream(src);
            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheet(sheetName);
            XSSFCell cell = sheet.getRow(row).getCell(col);
            if (cell == null){
                return true;
            }
            return false;
        }catch (IOException e){
            System.out.println(e.getMessage());
            return false;
        }
    }
}
