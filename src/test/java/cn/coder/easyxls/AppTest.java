package cn.coder.easyxls;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import junit.framework.TestCase;

/**
 * Unit test for simple App.
 */
public class AppTest extends TestCase {
	public static void main(String[] args) throws FileNotFoundException {
		Workbook w = new Workbook("test");
		w.addTitle("删掉");
		w.addTitle("订单");
		w.addData("sdf", 1214);
		w.addData("sdf", 679);
		w.write(new FileOutputStream(new File("d://sdf.xlsx")));
		w.close();
	}
}
