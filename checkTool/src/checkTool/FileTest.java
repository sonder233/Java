package checkTool;

import java.io.File;

public class FileTest {
	public static void main(String[] args) {
		File file = new File("G:\\�������\\���ݽṹ\\�δ���\\Test");
		String[] list = file.list();
		for(int i=0;i<list.length;i++) {
			System.out.println(list[i]);
		}
	}
}
