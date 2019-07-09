package cn.coder.easyxls;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.zip.ZipOutputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import cn.coder.easyxls.util.XLSUtils;
import cn.coder.easyxls.util.ZipUtils;

public final class Workbook {
	private static final Logger logger = LoggerFactory.getLogger(Workbook.class);

	private Sheet[] sheets;
	private Sheet defaultSheet;

	public Workbook(String sheet) {
		this.defaultSheet = new Sheet(sheet, 1);
		sheets = new Sheet[] { this.defaultSheet };
		if (logger.isDebugEnabled())
			logger.debug("Add default sheet:{}", this.defaultSheet.getName());
	}

	public void addTitle(String title) {
		this.defaultSheet.addTitle(title);
	}

	public void addData(Object... data) {
		this.defaultSheet.addData(data);
	}

	public boolean write(OutputStream outputStream) {
		long start = System.currentTimeMillis();
		ZipOutputStream zipStream = null;
		ArrayList<String> strings = null;
		try {
			byte[] workBook = XLSUtils.getWorkbook(sheets);
			byte[] workBookRels = XLSUtils.getWorkbookRels(sheets);
			byte[] app = XLSUtils.getApp(sheets);
			zipStream = ZipUtils.createZip(outputStream, workBook, workBookRels, app);

			strings = new ArrayList<>();
			for (Sheet sheet : sheets) {
				putSheet(zipStream, strings, sheet);
			}
			putStrings(zipStream, strings);

			return true;
		} catch (IOException e) {
			if (logger.isErrorEnabled())
				logger.error("Create xlsx file faild", e);
			return false;
		} finally {
			if (strings != null) {
				strings.clear();
			}
			ZipUtils.close(zipStream, outputStream);
			if (logger.isDebugEnabled())
				logger.debug("Created xlsx file with {}ms", (System.currentTimeMillis() - start));
		}
	}

	private static void putSheet(ZipOutputStream zipStream, ArrayList<String> strings, Sheet sheet) throws IOException {
		byte[] sheetData = XLSUtils.getSheet(sheet, strings);
		ZipUtils.putStreamEntry(zipStream, "xl/worksheets/sheet" + sheet.getId() + ".xml", sheetData);
	}

	private static void putStrings(ZipOutputStream zipStream, ArrayList<String> strings) throws IOException {
		byte[] temp = XLSUtils.getStrings(strings);
		ZipUtils.putStreamEntry(zipStream, "xl/sharedStrings.xml", temp);
	}

	public synchronized void close() {
		this.defaultSheet.clear();
		this.defaultSheet = null;
		for (Sheet sheet : sheets) {
			sheet.clear();
		}
		this.sheets = null;
	}

	public Sheet addSheet(String name) {
		Sheet sheet = new Sheet(name, sheets.length + 1);
		Sheet[] temp = new Sheet[sheets.length + 1];
		System.arraycopy(sheets, 0, temp, 0, sheets.length);
		temp[sheets.length] = sheet;
		sheets = temp;
		if (logger.isDebugEnabled())
			logger.debug("Add sheet:{}, total:{}", sheet.getName(), temp.length);
		return sheet;
	}

}
