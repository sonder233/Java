package checkTool;

import java.util.Iterator;

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

public class Check2 {
	//需要填写的Excel的路径
	static String fileNameOfExcel = "G:\\网课相关\\数据结构\\课代表\\考勤\\考勤-数据结构-18-1计课.xls";
	//考勤excel
	static String fileNameOfExcel2 = "G:\\网课相关\\数据结构\\课代表\\考勤\\每次课签到情况\\5780148_202004071043494071.xls";
	//已交同学的填充内容
	static String confirmText = "√";
	//学号所在列
	static int indexOfNumber = 9;
	//第二个excel学号所在列
	static int indexOfNumber2 = 8;
	//填充内容所在列
	static int indexOfContent = 21;
	//第一个同学所在的行数
	static int indexOfFirstStudent = 2;
	//第一个同学所在的行数(2)
		static int indexOfFirstStudent2 = 1;
	//全体学生的数量（总行数）
	static int numberOfStudent = 146;
	//全体学生的数量（总行数）(2)
		static int numberOfStudent2 = 139;
	//学号的长度
	static int LENGTH_OF_ID = 10;
	
	public static void main(String[] args) {
		//建立学号数组
				String[] str_id = new String[numberOfStudent2];
				
				try {
					File file = new File(fileNameOfExcel);
					File file2 = new File(fileNameOfExcel2);
					FileInputStream fin = new FileInputStream(file);
					FileInputStream fin2 = new FileInputStream(file2);
					Workbook workbook= new HSSFWorkbook(fin);
					Workbook workbook2= new HSSFWorkbook(fin2);
					Sheet sheet = workbook.getSheetAt(0);
					Sheet sheet2 = workbook2.getSheetAt(0);
					
					
					for(int i=indexOfFirstStudent2,j=0;i<=numberOfStudent2;i++,j++) {
						Row row = sheet2.getRow(i);
						str_id[j] = row.getCell(indexOfNumber2).getStringCellValue().trim();
						//ystem.out.println(str_id[i]+"  "+i);
					}
					
					for(int i=indexOfFirstStudent;i<numberOfStudent;i++) {
						Row row = sheet.getRow(i);
						Cell cell = row.getCell(indexOfNumber);
						//System.out.println(cell.toString());
						//调用对比函数
						if(check_id(cell.toString(),str_id)) {
							//单元格为空
							if(cell.getCellType() !=CellType.BLANK) {
								Cell newCell = row.createCell(indexOfContent);
								newCell.setCellValue(confirmText);
							} else {
								cell.setCellValue(confirmText);
							}
							
						}
						System.out.println("已完成"+(int)((i*1.0/numberOfStudent)*100)+"%");
					
					}
					
					FileOutputStream fout = new FileOutputStream(file);
					workbook.write(fout);
					fin.close(); 
					fout.close();
					fin2.close();
					System.out.println("已完成100%");
					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
	}
	/**
	 * 对比是否存在该学号
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
