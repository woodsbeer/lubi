package invoice;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.*;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class ReadExcel {
 
    public static void main(String[] args) 
    {
        File file1 = new File("C:\\Users\\xynld\\Desktop\\鲁碧财务发票Excel处理\\会计科目明细.xls");
        try 
        {
            List<List> list1 = importExcel(file1);


        } catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
 
 
    public static List<List> importExcel(File file) throws Exception 
    {
        Workbook wb = null;
        String fileName = file.getName();// 读取上传文件(excel)的名字，含后缀后
        // 根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
        Iterator<Sheet> sheets = null;
        List<List> returnlist = new ArrayList<List>();
        try 
        {
            if (fileName.endsWith("xls")) 
            {
                wb = new HSSFWorkbook(new FileInputStream(file));
                sheets = wb.iterator();
            } 
            else if (fileName.endsWith("xlsx")) 
            {
                wb = new XSSFWorkbook(new FileInputStream(file));
                sheets = wb.iterator();
            }
            if (sheets == null) 
            {
                throw new Exception("excel中不含有sheet工作表");
            }
            // 遍历excel里每个sheet的数据。
            int k = 1;
            while (sheets.hasNext()) 
            {
                System.out.println("-----Sheet " + k++ + "-----");
                Sheet sheet = sheets.next();
                List<Map> list = getCellValue(sheet);
                ListIterator<Map> iter = list.listIterator();
                while(iter.hasNext())
                	System.out.println(iter.next());
                returnlist.add(list);
            }
        } 
        catch (Exception ex) 
        {
            throw ex;
        } 
        finally 
        {
            if (wb != null) wb.close();
        }
        return returnlist;
    }
 
 
    // 获取每一个Sheet工作表中的数。
    private static List<Map> getCellValue(Sheet sheet) 
    {
        List<Map> list = new ArrayList<Map>();
        // sheet.getPhysicalNumberOfRows():获取的是物理行数，也就是不包括那些空行（隔行）的情况
        for (int i = sheet.getFirstRowNum(); i < sheet.getPhysicalNumberOfRows(); i++) 
        {
            Map map = new HashMap<>();
            // 获得第i行对象
            Row row = sheet.getRow(i);
            if (row == null) 
            {
                continue;
            } 
            else 
            {
            	int j = row.getFirstCellNum();// 获取第i行第一个单元格的下标
            	String level = row.getCell(j++).getStringCellValue();
            	map.put("科目级别",level);
            	Cell cell = row.getCell(j++);
            	
            	
            	if(cell.getCellTypeEnum().equals(CellType.STRING))
            	{
            		try
            		{
            			String codeStr = cell.getStringCellValue();
            			int code = Integer.parseInt(codeStr);
            			map.put("科目编码",code);
            		}
            		catch(Exception e)
            		{
            			continue;	
            		}
            	}
            	else if(cell.getCellTypeEnum().equals(CellType.NUMERIC))
            	{
            		map.put("科目编码",(int)cell.getNumericCellValue());
            	}
            	
                String name = row.getCell(j++).getStringCellValue();
                map.put("名称",name);
                       
                list.add(map);
            }
        }
        return list;
    }
    
    
}

