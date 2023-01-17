package Utils;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ExcelUtils {

public static void getColor() throws IOException {

//   String projdir= System.getProperty("user.dir");
//    System.out.print(projdir);
    String path =  "./Data/Book1.xlsx";
    System.out.print(path);
    XSSFWorkbook workbook=new XSSFWorkbook(path);
    XSSFSheet sheet = workbook.getSheet("Color");
    String color = sheet.getRow(0).getCell(0).toString();
    System.out.print(color);
   HSSFColor hssfColor= new HSSFColor();
  //  HSSFFont hssfFont = new HSSFFont();
    XSSFFont xssfFont= new XSSFFont();
  //  short a =


}

    public static void main(String[] args) throws IOException {
        getColor();
    }
}
