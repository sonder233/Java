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
	//��Ҫ��д��Excel��·��
	static String fileNameOfExcel = "G:\\�������\\���ݽṹ\\�δ���\\����\\����-���ݽṹ-18-1�ƿ�.xls";
	//����excel
	static String fileNameOfExcel2 = "G:\\�������\\���ݽṹ\\�δ���\\����\\ÿ�ο�ǩ�����\\5780148_202004071043494071.xls";
	//�ѽ�ͬѧ���������
	static String confirmText = "��";
	//ѧ��������
	static int indexOfNumber = 9;
	//�ڶ���excelѧ��������
	static int indexOfNumber2 = 8;
	//�������������
	static int indexOfContent = 21;
	//��һ��ͬѧ���ڵ�����
	static int indexOfFirstStudent = 2;
	//��һ��ͬѧ���ڵ�����(2)
		static int indexOfFirstStudent2 = 1;
	//ȫ��ѧ������������������
	static int numberOfStudent = 146;
	//ȫ��ѧ������������������(2)
		static int numberOfStudent2 = 139;
	//ѧ�ŵĳ���
	static int LENGTH_OF_ID = 10;
	
	public static void main(String[] args) {
		//����ѧ������
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
						//���öԱȺ���
						if(check_id(cell.toString(),str_id)) {
							//��Ԫ��Ϊ��
							if(cell.getCellType() !=CellType.BLANK) {
								Cell newCell = row.createCell(indexOfContent);
								newCell.setCellValue(confirmText);
							} else {
								cell.setCellValue(confirmText);
							}
							
						}
						System.out.println("�����"+(int)((i*1.0/numberOfStudent)*100)+"%");
					
					}
					
					FileOutputStream fout = new FileOutputStream(file);
					workbook.write(fout);
					fin.close(); 
					fout.close();
					fin2.close();
					System.out.println("�����100%");
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
