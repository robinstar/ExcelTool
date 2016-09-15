package tools.fjl.excel.poi;

public enum Operation {

	COPY(0, "复制单元格序列", CellListWorker.class.getName()), MULTI_ROW(1, "复制多行记录",
			MultiRowWorker.class.getName());

	int id;
	String title;
	String javaClass;

	private Operation(int id, String title, String javaClass) {
		this.id = id;
		this.title = title;
		this.javaClass = javaClass;
	}

	public String getTitle() {
		return title;
	}

	public String getJavaClass() {
		return javaClass;
	}

	public static Operation getOpertionByTitle(String title) {
		for (Operation op : values()) {
			if (op.title.equals(title)) {
				return op;
			}
		}
		throw new IllegalArgumentException();
	}

	public static Operation getOpertionById(int id) {
		for (Operation op : values()) {
			if (op.id == id) {
				return op;
			}
		}
		throw new IllegalArgumentException();
	}

	public static int getIdByTitle(String title) {
		for (Operation op : values()) {
			if (op.title.equals(title)) {
				return op.id;
			}
		}
		throw new IllegalArgumentException();
	}
}
