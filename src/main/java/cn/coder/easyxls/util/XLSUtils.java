package cn.coder.easyxls.util;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

import cn.coder.easyxls.Row;
import cn.coder.easyxls.Sheet;

public final class XLSUtils {
	private static final String XML_VERSION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
	private static final String XML_WORKBOOK = "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><fileVersion appName=\"xl\" lastEdited=\"3\" lowestEdited=\"5\"	rupBuild=\"9302\" /><workbookPr /><bookViews><workbookView windowWidth=\"22943\" windowHeight=\"10067\" /></bookViews><sheets>%s</sheets><calcPr calcId=\"144525\" /></workbook>";
	private static final String Relationship_SHEET = "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>";
	private static final String Relationship_THEME = "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>";
	private static final String Relationship_STYLE = "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>";
	private static final String Relationship_STRINGS = "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>";

	public static byte[] getApp(Sheet[] sheets) throws IOException {
		String resource = getResource("app.xml");
		StringBuilder str = new StringBuilder();
		for (Sheet sheet : sheets) {
			str.append("<vt:lpstr>");
			str.append(sheet.getName());
			str.append("</vt:lpstr>");
		}
		return String.format(resource, sheets.length, str).getBytes("utf-8");
	}

	public static byte[] getWorkbook(Sheet[] sheetArray) throws IOException {
		StringBuilder sb = new StringBuilder(XML_VERSION);
		StringBuilder sheets = new StringBuilder();
		for (Sheet sheet : sheetArray) {
			sheets.append("<sheet name=\"");
			sheets.append(sheet.getName());
			sheets.append("\" sheetId=\"");
			sheets.append(sheet.getId());
			sheets.append("\" r:id=\"rId");
			sheets.append(sheet.getId());
			sheets.append("\" />");
		}
		sb.append(String.format(XML_WORKBOOK, sheets));
		return sb.toString().getBytes("utf-8");
	}

	public static byte[] getWorkbookRels(Sheet[] sheets) throws IOException {
		StringBuilder sb = new StringBuilder(XML_VERSION);
		sb.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
		for (Sheet sheet : sheets) {
			sb.append(String.format(Relationship_SHEET, sheet.getId(), sheet.getId()));
		}
		int num = sheets.length + 1;
		sb.append(String.format(Relationship_THEME, num + 1));
		sb.append(String.format(Relationship_STYLE, num + 2));
		sb.append(String.format(Relationship_STRINGS, num + 3));
		sb.append("</Relationships>");
		return sb.toString().getBytes("utf-8");
	}

	public static byte[] getSheet(Sheet sheet, ArrayList<String> strings) throws IOException {
		// 添加表头到第一行
		Row firstRow = new Row(sheet.getTitles(), 1);
		List<Row> rows = sheet.getDataRows();
		rows.add(0, firstRow);

		String resource = getResource((sheet.getId() == 1 ? "sheet1.xml" : "sheet2.xml"));
		StringBuilder body = new StringBuilder();

		Object[] arr;
		for (Row row : rows) {
			body.append("<row r=\"").append(row.getId()).append("\" spans=\"1:2\">");
			arr = row.getData();
			for (int i = 0; i < arr.length; i++) {
				body.append("<c r=\"").append((char) (i + 65)).append(row.getId());
				// 首行加粗
				if (row.getId() == 1)
					body.append("\" s=\"1");
				if (arr[i] instanceof String) {
					body.append("\" t=\"s\"><v>");
					body.append(getString(strings, arr[i].toString())).append("</v></c>");
				} else {
					body.append("\"><v>");
					body.append(arr[i]).append("</v></c>");
				}
			}
			body.append("</row>");
		}
		return String.format(resource, body).getBytes("utf-8");
	}

	public static byte[] getStrings(ArrayList<String> strings) throws IOException {
		String resource = getResource("sharedStrings.xml");
		if (strings.size() > 0) {
			StringBuilder sb2 = new StringBuilder();
			for (String str : strings) {
				sb2.append("<si><t>");
				sb2.append(str);
				sb2.append("</t></si>");
			}
			return String.format(resource, strings.size(), sb2.toString()).getBytes("utf-8");
		}
		return String.format(resource, 0, "").getBytes("utf-8");
	}

	private static String getResource(String name) throws IOException {
		InputStream is = XLSUtils.class.getClassLoader().getResourceAsStream(name);
		BufferedReader reader = new BufferedReader(new InputStreamReader(is, "utf-8"));
		StringBuilder sb = new StringBuilder();
		String line = null;
		while ((line = reader.readLine()) != null) {
			sb.append(line);
			sb.append("\n");
		}
		return sb.toString();
	}

	private static Integer getString(ArrayList<String> strings, String str) {
		int index = strings.indexOf(str);
		if (index == -1) {
			strings.add(str);
			return strings.size() - 1;
		}
		return index;
	}

}
