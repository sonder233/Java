package checkTool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Check {
	
	//存放文件的文件夹的路径
	static String fileNameOfHomework = "G:\\网课相关\\数据结构\\课代表\\Test";
	//需要填写的Excel的路径
	static String fileNameOfExcel = "G:\\网课相关\\数据结构\\课代表\\Test.xls";
	//已交同学的填充内容
	static String confirmText = "已交";
	//学号所在列
	static int indexOfNumber = 9;
	//填充内容所在列
	static int indexOfContent = 15;
	
	public static void main(String[] args) {
		//建立学号数组
				File file_hw = new File(fileNameOfHomework);
				String[] fileName = file_hw.list();
				String[] str_id = new String[fileName.length];
				int LENGTH_OF_ID = 10;

				for(int i = 0;i<fileName.length;i++) {
					str_id[i] = fileName[i].substring(fileName[i].indexOf("18"), LENGTH_OF_ID);
				}
				
				
				try {
					File file = new File(fileNameOfExcel);
					FileInputStream fin = new FileInputStream(file);
					Workbook workbook= new HSSFWorkbook(fin);
					Sheet sheet = workbook.getSheetAt(0);
							
					for(int i=1;i<9;i++) {
						Row row = sheet.getRow(i);
						Cell cell = row.getCell(indexOfNumber);
						//调用对比函数
						if(check_id(cell.toString(),str_id)) {
							if(cell.getCellType() !=CellType.BLANK) {
								Cell newCell = row.createCell(indexOfContent);
								newCell.setCellValue(confirmText);
							} else {
								cell.setCellValue(confirmText);
							}
							
						}
					
					}
					
					FileOutputStream fout = new FileOutputStream(file);
					workbook.write(fout);
					fin.close();
					fout.close();
					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
	}
	/**
	 * 对比是否存在改学号
	 * @param id 学号
	 * @param str 学号数组
	 * @return 存在返回true，不存在返回false
	 */
	
	public static boolean check_id(String id,String[] str) {
		
		for(int i = 0;i<str.length;i++) {
			if(str[i].equals(id)) {
				return true;
			}
		}
		
		return false;
	}
	
	
}
