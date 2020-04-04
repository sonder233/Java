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
	
	//����ļ����ļ��е�·��
	static String fileNameOfHomework = "G:\\�������\\���ݽṹ\\�δ���\\Test";
	//��Ҫ��д��Excel��·��
	static String fileNameOfExcel = "G:\\�������\\���ݽṹ\\�δ���\\Test.xls";
	//�ѽ�ͬѧ���������
	static String confirmText = "�ѽ�";
	//ѧ��������
	static int indexOfNumber = 9;
	//�������������
	static int indexOfContent = 15;
	
	public static void main(String[] args) {
		//����ѧ������
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
						//���öԱȺ���
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
	 * �Ա��Ƿ���ڸ�ѧ��
	 * @param id ѧ��
	 * @param str ѧ������
	 * @return ���ڷ���true�������ڷ���false
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
