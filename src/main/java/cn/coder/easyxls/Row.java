package cn.coder.easyxls;

public final class Row {

	private final int id;
	private Object[] data;

	public Row(Object[] data, int id) {
		this.id = id;
		this.data = data;
	}

	public int getId() {
		return this.id;
	}

	public Object[] getData() {
		return this.data;
	}

}
