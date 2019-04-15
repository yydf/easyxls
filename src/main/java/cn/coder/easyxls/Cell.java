package cn.coder.easyxls;

public class Cell {
	private String title;
	private int size;

	public Cell(String title, int size) {
		this.title = title;
		this.size = size;
	}

	public String getTitle() {
		return title;
	}

	public char getChar() {
		return (char) (size + 65);
	}
}
