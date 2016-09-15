package tools.fjl.excel;

import org.junit.Assert;
import org.junit.Test;

public class UtilsTest {

	@Test
	public void isExcelFile() {
		String path = "abcd.xls";
		Assert.assertTrue(Utils.isExcelFile(path));

		path = "abcd.xlsx";
		Assert.assertTrue(Utils.isExcelFile(path));

		path = "~$abcd.xls";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "~$abcd.xlsx";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "/User/~$abcd.xls";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "/User/~$abcd.xlsx";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "abcd.exe";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "abcd";
		Assert.assertFalse(Utils.isExcelFile(path));
	}

	@Test
	public void isXlsFile() {
		String path = "abcd.xls";
		Assert.assertTrue(Utils.isXlsFile(path));

		path = "abcd.xlsx";
		Assert.assertFalse(Utils.isXlsFile(path));

		path = "~$abcd.xls";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "~$abcd.xlsx";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "/User/~$abcd.xls";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "/User/~$abcd.xlsx";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "abcd.exe";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "abcd";
		Assert.assertFalse(Utils.isExcelFile(path));
	}

	@Test
	public void isXlsxFile() {
		String path = "abcd.xls";
		Assert.assertFalse(Utils.isXlsxFile(path));

		path = "abcd.xlsx";
		Assert.assertTrue(Utils.isXlsxFile(path));

		path = "~$abcd.xls";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "~$abcd.xlsx";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "/User/~$abcd.xls";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "/User/~$abcd.xlsx";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "abcd.exe";
		Assert.assertFalse(Utils.isExcelFile(path));

		path = "abcd";
		Assert.assertFalse(Utils.isExcelFile(path));
	}
}
