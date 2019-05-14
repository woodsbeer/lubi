package invoice;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\johnny\\AppData\\Local\\Temp\\Temp1_鲁碧财务发票Excel处理.zip\\鲁碧财务发票Excel处理";
        ExcelUtil e = new ExcelUtil();
        System.out.println(e.getColList("购方税号", filePath));
    }

    public int getColByName(String name, Sheet sheet) {
        Row row = sheet.getRow(0);
        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (row.getCell(i).getStringCellValue().equals(name)) {
                return i;
            }
        }
        return -1;
    }

   public List<String> getColList(String name, String path) {
        List<String> restList = new ArrayList();
        try {
            FileInputStream fis = new FileInputStream(path);
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            Sheet sheet = wb.getSheetAt(0);
            int colNum = getColByName(name, sheet);
            if (colNum < 0) {
                return null;
            }
            for (int i = 0; i < sheet.getLastRowNum(); i++) {
                Row r = sheet.getRow(i);
                restList.add(r.getCell(colNum).getStringCellValue());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return restList;


    }
}



