package cn.coder.easyxls.util;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

import cn.coder.easyxls.Cell;

public class XLSUtils {
	public static byte[] getWorkbook(String sheet) {
		try {
			InputStream is = XLSUtils.class.getClassLoader().getResourceAsStream("workbook.xml");
			BufferedReader reader = new BufferedReader(new InputStreamReader(is, "utf-8"));
			StringBuilder sb = new StringBuilder();
			String line = null;
			while ((line = reader.readLine()) != null) {
				sb.append(line + "\n");
			}
			return String.format(sb.toString(), sheet).getBytes();
		} catch (IOException e) {
			return null;
		}
	}

	public static byte[] getSheet(List<Cell> titleCells, List<Cell[]> dataCells, ArrayList<String> strings)
			throws IOException {
		InputStream is = XLSUtils.class.getClassLoader().getResourceAsStream("sheet1.xml");
		BufferedReader reader = new BufferedReader(new InputStreamReader(is, "utf-8"));
		StringBuilder sb = new StringBuilder();
		String line = null;
		while ((line = reader.readLine()) != null) {
			sb.append(line + "\n");
		}
		StringBuilder body = new StringBuilder();
		if (titleCells.size() > 0) {
			body.append("<row r=\"1\" spans=\"1:2\">");
			for (Cell cell : titleCells) {
				body.append("<c r=\"");
				body.append(cell.getChar());
				if (StringUtils.isEmpty(cell.getTitle()))
					body.append("1\" s=\"1\"><v>");
				else {
					body.append("1\" s=\"1\" t=\"s\"><v>");
					body.append(getString(strings, cell.getTitle()));
				}
				body.append("</v></c>");
			}
			body.append("</row>");
		}
		if (dataCells.size() > 0) {
			Cell[] cells;
			for (int i = 0; i < dataCells.size(); i++) {
				body.append("<row r=\"");
				body.append(i + 2);
				body.append("\" spans=\"1:2\">");
				cells = dataCells.get(i);
				for (Cell cell2 : cells) {
					body.append("<c r=\"");
					body.append(cell2.getChar());
					body.append(i + 2);
					if (StringUtils.isEmpty(cell2.getTitle()))
						body.append("\"><v>");
					else {
						body.append("\" s=\"2\" t=\"s\"><v>");
						body.append(getString(strings, cell2.getTitle()));
					}
					body.append("</v></c>");
				}
				body.append("</row>");
			}
		}
		return String.format(sb.toString(), body.toString()).getBytes();
	}

	private static Integer getString(ArrayList<String> strings, String str) {
		int index = strings.indexOf(str);
		if (index == -1) {
			strings.add(str);
			return strings.size() - 1;
		}
		return index;
	}

	public static byte[] getStrings(ArrayList<String> strings) throws IOException {
		InputStream is = XLSUtils.class.getClassLoader().getResourceAsStream("sharedStrings.xml");
		BufferedReader reader = new BufferedReader(new InputStreamReader(is, "utf-8"));
		StringBuilder sb = new StringBuilder();
		String line = null;
		while ((line = reader.readLine()) != null) {
			sb.append(line + "\n");
		}
		if (strings.size() > 0) {
			StringBuilder sb2 = new StringBuilder();
			for (String str : strings) {
				sb2.append("<si><t>");
				sb2.append(str);
				sb2.append("</t></si>");
			}
			return String.format(sb.toString(), strings.size() + "", sb2.toString()).getBytes();
		}
		return String.format(sb.toString(), "0", "").getBytes();
	}
}
