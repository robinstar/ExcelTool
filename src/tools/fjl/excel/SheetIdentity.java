package tools.fjl.excel;

public class SheetIdentity {

	final static char MAGIC1 = '\uEF00';
	final static char MAGIC2 = '\uEF01';
	public static final String FIRST_SHEET = "首张工作表" + MAGIC1;

	private String workbookName;
	private String sheetName;

	public SheetIdentity(String workbookName, String sheetName) {
		if (workbookName == null || "".equals(workbookName)) {
			throw new IllegalArgumentException("workbookName can not be empty!");
		}
		if (sheetName == null || "".equals(sheetName)) {
			throw new IllegalArgumentException("sheetName can not be empty!");
		}

		this.workbookName = workbookName;
		this.sheetName = sheetName;
	}

	public String getWorkbookName() {
		return workbookName;
	}

	public String getSheetName() {
		return sheetName;
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null) {
			return false;
		}

		if (obj instanceof SheetIdentity) {
			SheetIdentity sheet = (SheetIdentity) obj;
			return workbookName.equals(sheet.getWorkbookName()) && sheetName.equals(sheet.getSheetName());
		}
		return false;
	}

	@Override
	public int hashCode() {
		return workbookName.hashCode() * sheetName.hashCode();
	}

	@Override
	public String toString() {
		return "SheetIdentity [workbookName=" + workbookName + ", sheetName=" + sheetName + "]";
	}

}
