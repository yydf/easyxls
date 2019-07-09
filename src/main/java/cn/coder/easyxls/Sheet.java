package cn.coder.easyxls;

import java.util.ArrayList;
import java.util.List;

public final class Sheet {

	private final int id;
	private final String name;
	private final List<Object> titles = new ArrayList<>();
	private final List<Row> dataRows = new ArrayList<>();

	public Sheet(String name, int id) {
		this.name = name;
		this.id = id;
	}

	public void addTitle(String title) {
		titles.add(title);
	}

	public void addData(Object... data) {
		if (data.length == 0)
			return;
		this.dataRows.add(new Row(data, this.dataRows.size() + 2));
	}

	public String getName() {
		return this.name;
	}

	public int getId() {
		return this.id;
	}

	public void clear() {
		this.titles.clear();
		this.dataRows.clear();
	}

	public Object[] getTitles() {
		return this.titles.toArray();
	}

	public List<Row> getDataRows() {
		return this.dataRows;
	}

}
