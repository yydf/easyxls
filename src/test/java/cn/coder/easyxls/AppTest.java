package cn.coder.easyxls;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Unit test for simple App.
 */
public class AppTest {
	public static void main(String[] args) throws FileNotFoundException {
		Workbook w = new Workbook("test&<>' \"");
		w.addTitle("删&<>' \"掉");
		w.addTitle("订单");
		w.addData("sdf&<>' \"", "1&<>' \"");
		w.addData("sdf", "2");
		w.write(new FileOutputStream(new File("d://sdf.xlsx")));
		w.close();
	}
}
