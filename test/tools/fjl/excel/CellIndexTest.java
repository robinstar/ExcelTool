package tools.fjl.excel;

import static tools.fjl.excel.CellIndex.*;

import org.junit.Assert;
import org.junit.Test;

public class CellIndexTest {

	@Test(expected = IllegalArgumentException.class)
	public void testInvalidRow() {
		new CellIndex(-2, 0);
	}

	@Test(expected = IllegalArgumentException.class)
	public void testInvalidColumn() {
		new CellIndex(0, -2);
	}

	@Test
	public void testRowColumn() {
		new CellIndex(0, 0);
	}

	@Test
	public void testEquals() {
		int row = 10;
		int anotherRow = 11;
		int column = 10;
		int anotherColumn = 11;

		CellIndex cell = new CellIndex(row, column);
		CellIndex sameCell = new CellIndex(row, column);
		CellIndex diffRowCell = new CellIndex(anotherRow, column);
		CellIndex diffColumnCell = new CellIndex(row, anotherColumn);

		Assert.assertTrue(cell.equals(sameCell));
		Assert.assertFalse(cell.equals(diffRowCell));
		Assert.assertFalse(cell.equals(diffColumnCell));
	}

	@Test(expected = NullPointerException.class)
	public void testNullMark() {
		new CellIndex(null);
	}

	@Test
	public void testMark() {
		Assert.assertEquals(new CellIndex(UNDIFINED, UNDIFINED), new CellIndex(""));
		Assert.assertEquals(new CellIndex(0, UNDIFINED), new CellIndex("1"));
		Assert.assertEquals(new CellIndex(UNDIFINED, 0), new CellIndex("A"));
		Assert.assertEquals(new CellIndex(0, 0), new CellIndex("A1"));
	}

	@Test(expected = IllegalArgumentException.class)
	public void testIllegalColumnToIndex() {
		CellIndex.columnToIndex("A-a");
	}

	@Test(expected = IllegalArgumentException.class)
	public void testLargeColumnToIndex() {
		CellIndex.columnToIndex("XFE");
	}

	@Test
	public void testColumnToIndex() {
		Assert.assertEquals(MIN_COLUMN_INDEX, CellIndex.columnToIndex(MIN_COLUMN_MARK));
		Assert.assertEquals(MAX_COLUMN_INDEX, CellIndex.columnToIndex(MAX_COLUMN_MARK));
	}

	@Test(expected = IllegalArgumentException.class)
	public void testSmallIndexToColumn() {
		CellIndex.indexToColumn(MIN_COLUMN_INDEX - 1);
	}

	@Test(expected = IllegalArgumentException.class)
	public void testLargeIndexToColumn() {
		CellIndex.indexToColumn(MAX_COLUMN_INDEX + 1);
	}

	@Test
	public void testIndexToColumn() {
		Assert.assertEquals(MIN_COLUMN_MARK, CellIndex.indexToColumn(MIN_COLUMN_INDEX));
		Assert.assertEquals(MAX_COLUMN_MARK, CellIndex.indexToColumn(MAX_COLUMN_INDEX));
	}

	@Test(expected = IllegalArgumentException.class)
	public void testSmallRowToIndex() {
		CellIndex.rowToIndex(MIN_ROW_MARK - 1);
	}

	@Test(expected = IllegalArgumentException.class)
	public void testLargeRowToIndex() {
		CellIndex.rowToIndex(MAX_ROW_MARK + 1);
	}

	@Test
	public void testRowToIndex() {
		Assert.assertEquals(MIN_ROW_INDEX, CellIndex.rowToIndex(MIN_ROW_MARK));
		Assert.assertEquals(MAX_ROW_INDEX, CellIndex.rowToIndex(MAX_ROW_MARK));
	}

	@Test(expected = IllegalArgumentException.class)
	public void testSmallIndexToRow() {
		CellIndex.indexToRow(MIN_ROW_INDEX - 1);
	}

	@Test(expected = IllegalArgumentException.class)
	public void testLargeIndexToRow() {
		CellIndex.indexToRow(MAX_ROW_INDEX + 1);
	}

	@Test
	public void testIndexToRow() {
		Assert.assertEquals(MIN_ROW_MARK, CellIndex.indexToRow(MIN_ROW_INDEX));
		Assert.assertEquals(MAX_ROW_MARK, CellIndex.indexToRow(MAX_ROW_INDEX));
	}
}
