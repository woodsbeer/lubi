package invoice;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


public class ReadExcelForHSSF {

	public void read() 
	{
		//File file = new File("C:\\Users\\xynld\\Desktop\\鲁碧财务发票Excel处理\\会计科目明细.xls");
		File file = new File("C:\\Users\\xynld\\Desktop\\鲁碧财务发票Excel处理\\开票信息.xls");

		//System.out.println(fileName);
		
		if (!file.exists())
			System.out.println("文件不存在");
		
		
		try {

			// 读取Excel对象
			POIFSFileSystem poifsFileSystem = new POIFSFileSystem(new FileInputStream(file));
			
			
			// Excel工作簿对象
			HSSFWorkbook hssfWorkbook = new HSSFWorkbook(poifsFileSystem);
			// Excel sheet个数
			int sheetNumber = hssfWorkbook.getNumberOfSheets();
			
			
			for (int k = 0; k < sheetNumber; k++) 
			{
				HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(k);   //获取当前指定sheet
				
				int rowLength = hssfSheet.getLastRowNum() + 1;      // 当前sheet的总行数
				
				HSSFRow hssfRow = hssfSheet.getRow(0);              // 得到Excel工作表的行
				
				System.out.println(hssfSheet.getRow(0));
				int colLength = hssfRow.getLastCellNum();           // 总列数
				
				HSSFCell hssfCell = hssfRow.getCell(0);             // 得到Excell指定单元格中的内容
				
				CellStyle cellStyle = hssfCell.getCellStyle();      // 得到单元格样式

				System.out.println("--------Sheet " + (k+1) + "--------");
				for (int i = 0; i < rowLength; i++) {
					HSSFRow hssfRow1 = hssfSheet.getRow(i); // 获取Excel工作表的行
					for (int j = 0; j < colLength; j++) {
						// 获取指定单元格
						HSSFCell hssfCell1 = hssfRow1.getCell(j);

						// Excel数据Cell有不同的类型，当我们试图从一个数字类型的Cell读取出一个字符串时就有可能报异常：
						// Cannot get a STRING value from a NUMERIC cell
						// 将所有的需要读的Cell表格设置为String格式
						if (hssfCell1 != null) {
							hssfCell1.setCellType(CellType.STRING);
						}

						// 获取每一列中的值
						System.out.printf(hssfCell1.getStringCellValue() + "\t");
					}
					System.out.println();
				}

			}
			hssfWorkbook.close();
		} 
		catch (IOException e) 
		{
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		new ReadExcelForHSSF().read();
	}

}
