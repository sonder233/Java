package checkTool;

import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.WorkbookUtil; 

import java.io.FileOutputStream;

public class Main {
	public static void main(String[] args) {
		Workbook workbook = new HSSFWorkbook();
		
		//Sheet sheet1 = workbook.createSheet();
		//Sheet sheet2 = workbook.createSheet("Pancakes");
		//注意这里目前仅支持英文，不支持其他的符号
		//要是想使它支持，需要引用其他的包
		//Sheet sheet3 = workbook.createSheet(WorkbookUtil.createSafeSheetName("???asda[]"));
		//这样创建的表名更加安全
		Sheet sheet = workbook.createSheet("Eggs");
		//Row row = sheet.createRow(0);
		Cell cell1 = sheet.createRow(0).createCell(0); //A=0
		Cell cell2 = sheet.createRow(0).createCell(1);
		Cell cell3 = sheet.createRow(0).createCell(2);
		Cell cell4 = sheet.createRow(0).createCell(3);
		Cell cell5 = sheet.createRow(0).createCell(4);
		cell1.setCellValue(33);
		cell2.setCellValue(41);
		cell3.setCellValue(22);
		cell4.setCellValue(12);
		cell5.setCellFormula("SUM(A1:D1)");
		//System.out.println(cell.getRichStringCellValue().toString());
		
		try {
			FileOutputStream output = new FileOutputStream("Test1.xls");
			workbook.write(output);
			output.close();
		} catch(Exception e) {
			e.printStackTrace();
		}
//		
//		String testNumber = "数据结构-1812030039-高龙";
//	
//		//获取文件名
//		List<String> list = new ArrayList();
//		String path = "G:\\网课相关\\数据结构\\课代表\\考勤\\每次课签到情况";
//		File file = new File(path);
//		File[] files = file.listFiles();
//		for(int i=0;i<files.length;i++) {
//			System.out.println(files[i].getName());
//		}
	}
	
	public static boolean checkString (String str) {
		boolean ishave = false;
		ishave = str.contains("18");
		return ishave;
	}
}
