package checkTool;

import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import org.apache.poi.ss.usermodel.*;

import java.util.Iterator;
import java.io.*;
import javax.swing.JFileChooser;


public class Excel2 {
	public static void main(String[] args) {
		//Java自带的文件选择器
		//JFileChooser fileChooser = new JFileChooser();
		//int returnValue = fileChooser.showOpenDialog(null);
		//为点击按钮添加事件
		//if(returnValue == fileChooser.APPROVE_OPTION) {
			
		//}
		
		//建立学号数组
		File file_hw = new File("G:\\网课相关\\数据结构\\课代表\\Test");
		String[] fileName = file_hw.list();
		String[] str_id = new String[fileName.length];
		int LENGTH_OF_ID = 10;

		for(int i = 0;i<fileName.length;i++) {
			str_id[i] = fileName[i].substring(fileName[i].indexOf("18"), LENGTH_OF_ID);
		}
		
		for(int i = 0;i<str_id.length;i++) {
			System.out.print(str_id[i] + "\t");
		}
		
		
		
		try {
			//Workbook workbook = new HSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
			File file = new File("G:\\网课相关\\数据结构\\课代表\\Test.xls");
			FileInputStream fin = new FileInputStream(file);
			Workbook workbook= new HSSFWorkbook(fin);
			Sheet sheet = workbook.getSheetAt(0);
					
			for(int i=1;i<9;i++) {
				Row row = sheet.getRow(i);
				
				if(check_id(row.getCell(9).toString(),str_id)) {
					row.getCell(15).setCellValue("已交");
				}
				else {
					row.getCell(15).setCellValue("");
				}
			}
			
			FileOutputStream fout = new FileOutputStream(file);
			workbook.write(fout);
			
//			for(Iterator<Row> rit = sheet.iterator();rit.hasNext();) {
//				Row row = rit.next();
//				
//				for(Iterator<Cell> cit = row.cellIterator();cit.hasNext();) {
//					Cell cell = cit.next();
//					System.out.print(cell.toString() + "\t");
//				}
//				System.out.println();
//			}
		fin.close();
		fout.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static boolean check_id(String id,String[] str) {
	
		for(int i = 0;i<str.length;i++) {
			if(str[i].equals(id)) {
				return true;
			}
		}
		
		return false;
	}
	
}
