package invoice;

import java.io.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class Linkin
{
	public static void main(String[] args)
	{
		FileInputStream in = null;
		HSSFWorkbook workbook = null;
 
 
		try
		{
			in = new FileInputStream("C:\\Users\\xynld\\Desktop\\鲁碧财务发票Excel处理\\会计科目明细.xls");
			POIFSFileSystem fs = new POIFSFileSystem(in);
			workbook = new HSSFWorkbook(fs);
		}
		catch (IOException e)
		{
			System.out.println(e.toString());
		}
		finally
		{
			try
			{
				in.close();
			}
			catch (IOException e)
			{
				System.out.println(e.toString());
			}
		}
 
		int sheetNum = workbook.getNumberOfSheets();
		
		for(int k = 0;k<sheetNum;k++)
		{
			HSSFSheet sheet = workbook.getSheetAt(k);//读取序号为0的sheet
			int rowNum = sheet.getLastRowNum() + 1;

			System.out.println(rowNum);
		}
		
 
 
		//HSSFRow row = sheet.getRow(2);//取得sheet中第二行（行号1）
 
 
		//HSSFCell cell = row.getCell((short) 0);//取得第二行，第二格（单元格号1）
		//System.out.println(cell.getStringCellValue());//cell.getStringCellValue()取值
	}
 
 
}
