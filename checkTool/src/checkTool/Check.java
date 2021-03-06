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
	static String fileNameOfHomework = "G:\\网课相关\\数据结构\\课代表\\作业\\第五次作业（第四章）";
	//需要填写的Excel的路径
	static String fileNameOfExcel = "G:\\网课相关\\数据结构\\课代表\\作业\\数据结构1班作业情况 (第五次.xls";
	//考勤excel
	static String fileNameOfExcel2 = "G:\\网课相关\\数据结构\\课代表\\考勤\\每次课签到情况\\5780148_202004071043494071.xls";
	//已交同学的填充内容
	static String confirmText = "已交";
	//学号所在列
	static int indexOfNumber = 1;
	//填充内容所在列
	static int indexOfContent = 4;
	//第一个同学所在的行数
	static int indexOfFirstStudent = 2;
	//全体学生的数量（总行数）
	static int numberOfStudent = 145;
	//学号的长度
	static int LENGTH_OF_ID = 10;
	
	public static void main(String[] args) {
		//建立学号数组
				File file_hw = new File(fileNameOfHomework);
				String[] fileName = file_hw.list();
				String[] str_id = new String[fileName.length];
				
				for(int i = 0;i<fileName.length;i++) {
					
					int index;
					
					if(fileName[i].contains("18") && fileName[i].contains("17")) {
						index = Math.min(fileName[i].indexOf("18"), fileName[i].indexOf("17"));
						str_id[i] = fileName[i].substring(index,index+LENGTH_OF_ID);
					} else if(fileName[i].contains("17")){
						index = fileName[i].indexOf("17");
						str_id[i] = fileName[i].substring(index,index+LENGTH_OF_ID);
					}else if(fileName[i].contains("18")){
						index = fileName[i].indexOf("18");
						str_id[i] = fileName[i].substring(index,index+LENGTH_OF_ID);
					}else {
						System.out.println("这个文件不包含学号"+fileName[i]);
					}
				}
				
//				for(String str:str_id) {
//					System.out.println(str);
//				}
				
				try {
					File file = new File(fileNameOfExcel);
					FileInputStream fin = new FileInputStream(file);
					Workbook workbook= new HSSFWorkbook(fin);
					Sheet sheet = workbook.getSheetAt(0);

							
					for(int i=indexOfFirstStudent;i<=numberOfStudent;i++) {
						Row row = sheet.getRow(i);
						Cell cell = row.getCell(indexOfNumber);
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
