package checkTool;

import java.io.File;

public class FileTest {
	public static void main(String[] args) {
		File file = new File("G:\\网课相关\\数据结构\\课代表\\Test");
		String[] list = file.list();
		for(int i=0;i<list.length;i++) {
			System.out.println(list[i]);
		}
	}
}
