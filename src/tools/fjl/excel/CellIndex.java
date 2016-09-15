package tools.fjl.excel;

public class CellIndex {

	public static final int UNDIFINED = -1;

	public static final String MIN_COLUMN_MARK = "A";
	public static final String MAX_COLUMN_MARK = "XFD";

	public static final int MIN_COLUMN_INDEX = 0;
	public static final int MAX_COLUMN_INDEX = 16383;

	public static final int MIN_ROW_INDEX = 0;
	public static final int MAX_ROW_INDEX = 9998;

	public static final int MIN_ROW_MARK = 1;
	public static final int MAX_ROW_MARK = 9999;

	public static final class Regex {
		public static final String ROW = "[1-9]|[1-9]\\d{1,4}";
		public static final String COLUMN = "[a-zA-Z]{1,3}";
		public static final String CELL = "(" + COLUMN + ")(" + ROW + ")";
	}

	public final int row;
	public final int column;

	public CellIndex(int row, int column) {
		if (row < UNDIFINED) {
			throw new IllegalArgumentException("Invalid row index");
		}

		if (column < UNDIFINED) {
			throw new IllegalArgumentException("Invalid column index");
		}

		this.row = row;
		this.column = column;
	}

	public CellIndex(String mark) {
		if (mark == null) {
			throw new NullPointerException();
		}

		int row = CellIndex.UNDIFINED;
		int column = CellIndex.UNDIFINED;

		if (mark.matches(Regex.ROW)) {
			row = rowToIndex(Integer.valueOf(mark));
		}

		if (mark.matches(Regex.COLUMN)) {
			column = columnToIndex(mark);
		}

		if (mark.matches(Regex.CELL)) {
			int firstNumIndex = 0;
			char c;
			for (;; firstNumIndex++) {
				c = mark.charAt(firstNumIndex);
				if (Character.isDigit(c)) {
					break;
				}
			}

			String columePart = mark.substring(0, firstNumIndex);
			column = columnToIndex(columePart);

			String rowPart = mark.substring(firstNumIndex);
			row = rowToIndex(Integer.valueOf(rowPart));
		}

		this.row = row;
		this.column = column;
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null) {
			return false;
		}

		if (obj instanceof CellIndex) {
			CellIndex cell = (CellIndex) obj;
			return cell.row == row && cell.column == column;
		}
		return false;
	}

	public static int columnToIndex(String column) {
		if (!column.matches(Regex.COLUMN)) {
			throw new IllegalArgumentException("Invalid column mark");
		}

		column = column.toUpperCase();

		int index = 0;
		char[] chars = column.toUpperCase().toCharArray();
		for (int i = 0; i < chars.length; i++) {
			index += ((int) chars[i] - (int) 'A' + 1) * (int) Math.pow(26, chars.length - i - 1);
		}

		index--;
		if (index > MAX_COLUMN_INDEX) {
			throw new IllegalArgumentException("unsupport column mark");
		}

		return index;
	}

	public static String indexToColumn(int index) {
		if (index < MIN_COLUMN_INDEX || index > MAX_COLUMN_INDEX) {
			throw new IllegalArgumentException("Invalid column index");
		}

		String column = "";
		do {
			if (column.length() > 0) {
				index--;
			}
			column = ((char) (index % 26 + (int) 'A')) + column;
			index = (int) ((index - index % 26) / 26);
		} while (index > 0);

		return column;
	}

	public static int rowToIndex(int row) {
		if (row < MIN_ROW_MARK || row > MAX_ROW_MARK) {
			throw new IllegalArgumentException();
		}
		return row - 1;
	}

	public static int indexToRow(int index) {
		if (index < MIN_ROW_INDEX || index > MAX_ROW_INDEX) {
			throw new IllegalArgumentException();
		}
		return index + 1;
	}
}