package cn.coder.easyxls;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipOutputStream;

import cn.coder.easyxls.util.XLSUtils;
import cn.coder.easyxls.util.ZipUtils;

public class Workbook {

	private byte[] workBook;
	private final List<Cell> titleCells = new ArrayList<>();
	private final List<Cell[]> dataCells = new ArrayList<>();

	public Workbook(String defaultSheet) {
		this.workBook = XLSUtils.getWorkbook(defaultSheet);
	}

	public void addTitle(String title) {
		this.titleCells.add(new Cell(title, this.titleCells.size()));
	}

	public void addData(Object... data) {
		if (data.length == 0)
			return;
		Cell[] cells = new Cell[data.length];
		for (int i = 0; i < data.length; i++) {
			if (data[i] == null)
				cells[i] = new Cell("", i);
			else
				cells[i] = new Cell(data[i].toString(), i);
		}
		this.dataCells.add(cells);
	}

	public boolean write(OutputStream outputStream) {
		if (workBook == null)
			return false;
		try {
			ArrayList<String> strings = new ArrayList<>();
			byte[] sheetData = XLSUtils.getSheet(titleCells, dataCells, strings);
			byte[] stringsData = XLSUtils.getStrings(strings);
			ZipOutputStream zipStream = new ZipOutputStream(outputStream);
			ZipUtils.putEntry(zipStream, "_rels/.rels", "_rels.rels.xml");
			ZipUtils.putEntry(zipStream, "docProps/app.xml", "app.xml");
			ZipUtils.putEntry(zipStream, "docProps/core.xml", "core.xml");
			ZipUtils.putEntry(zipStream, "docProps/custom.xml", "custom.xml");
			ZipUtils.putStreamEntry(zipStream, "xl/workbook.xml", this.workBook);
			ZipUtils.putEntry(zipStream, "xl/styles.xml", "styles.xml");
			ZipUtils.putStreamEntry(zipStream, "xl/worksheets/sheet1.xml", sheetData);
			ZipUtils.putStreamEntry(zipStream, "xl/sharedStrings.xml", stringsData);
			ZipUtils.putEntry(zipStream, "xl/_rels/workbook.xml.rels", "workbook.xml.rels.xml");
			ZipUtils.putEntry(zipStream, "xl/worksheets/sheet2.xml", "sheet2.xml");
			ZipUtils.putEntry(zipStream, "xl/worksheets/sheet3.xml", "sheet3.xml");
			ZipUtils.putEntry(zipStream, "xl/theme/theme1.xml", "theme1.xml");
			ZipUtils.putEntry(zipStream, "[Content_Types].xml", "Content_Types.xml");
			zipStream.close();
			return true;
		} catch (IOException e) {
			return false;
		} finally {
			try {
				outputStream.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void close() {
		this.titleCells.clear();
		this.dataCells.clear();
		this.workBook = null;
	}

}
