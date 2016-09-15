package tools.fjl.excel;

import org.junit.Assert;
import org.junit.Test;

public class SheetIdentityTest {

	@Test(expected = IllegalArgumentException.class)
	public void testNullBook() {
		new SheetIdentity(null, "sheet");
	}

	@Test(expected = IllegalArgumentException.class)
	public void testEmptyBook() {
		new SheetIdentity("", "sheet");
	}

	@Test(expected = IllegalArgumentException.class)
	public void testNullSheet() {
		new SheetIdentity("book", null);
	}

	@Test(expected = IllegalArgumentException.class)
	public void testEmptySheet() {
		new SheetIdentity("book", "");
	}

	@Test(expected = IllegalArgumentException.class)
	public void testNullBookNullSheet() {
		new SheetIdentity(null, null);
	}

	@Test(expected = IllegalArgumentException.class)
	public void testNullBookEmptySheet() {
		new SheetIdentity(null, "");
	}

	@Test(expected = IllegalArgumentException.class)
	public void testEmptyBookNullSheet() {
		new SheetIdentity("", null);
	}

	@Test(expected = IllegalArgumentException.class)
	public void testEmptyBookEmptySheet() {
		new SheetIdentity("", "");
	}

	@Test
	public void testConstructor() {
		new SheetIdentity("book", "sheet");
	}

	@Test
	public void testEquals() {
		String book = "book";
		String anotherBook = "book2";
		String sheet = "sheet";
		String anotherSheet = "sheet2";

		SheetIdentity param = new SheetIdentity(book, sheet);
		SheetIdentity sameParam = new SheetIdentity(book, sheet);
		SheetIdentity diffBookParam = new SheetIdentity(anotherBook, sheet);
		SheetIdentity diffSheetParam = new SheetIdentity(book, anotherSheet);

		Assert.assertTrue(param.equals(sameParam));
		Assert.assertFalse(param.equals(diffBookParam));
		Assert.assertFalse(param.equals(diffSheetParam));
	}

	@Test
	public void testHashCode() {
		String book = "book";
		String anotherBook = "book2";
		String sheet = "sheet";
		String anotherSheet = "sheet2";

		SheetIdentity param = new SheetIdentity(book, sheet);
		SheetIdentity sameParam = new SheetIdentity(book, sheet);
		SheetIdentity diffBookParam = new SheetIdentity(anotherBook, sheet);
		SheetIdentity diffSheetParam = new SheetIdentity(book, anotherSheet);

		Assert.assertEquals(param.hashCode(), sameParam.hashCode());
		Assert.assertNotEquals(param.hashCode(), diffBookParam.hashCode());
		Assert.assertNotEquals(param.hashCode(), diffSheetParam.hashCode());

		Assert.assertEquals(param, sameParam);
		Assert.assertNotEquals(param, diffBookParam);
		Assert.assertNotEquals(param, diffSheetParam);
	}
}
