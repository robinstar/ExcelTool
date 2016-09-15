package tools.fjl.excel.poi;

import java.io.File;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.robin.util.StringUtils;

import tools.fjl.excel.BadParamException;
import tools.fjl.excel.ExcelException;
import tools.fjl.excel.ExcelTool;
import tools.fjl.excel.SheetIdentity;

abstract class BaseTool implements ExcelTool {

	SheetIdentity source;
	SheetIdentity target;
	String paramPath;

	Workbook sourceWorkbook;
	Workbook targetWorkbook;

	Sheet sourceSheet;
	Sheet targetSheet;

	private boolean isTargetSet = false;
	private boolean isParamSet = false;

	public final void setSourceSheet(SheetIdentity source) throws ExcelException {
		if (source == null) {
			throw new NullPointerException("source sheet can not be null");
		}

		if (target != null && target.equals(source)) {
			throw new IllegalArgumentException("source can not same as target");
		}

		this.source = source;

		final String sourcePath = source.getWorkbookName();
		if (target != null && target.getWorkbookName().equals(sourcePath)) {
			sourceWorkbook = targetWorkbook;
		} else {
			sourceWorkbook = PoiUtils.getWorkbookByPath(sourcePath);
		}

		final String sheetName = source.getSheetName();
		if (SheetIdentity.FIRST_SHEET.equals(sheetName)) {
			sourceSheet = sourceWorkbook.getSheetAt(0);
		} else {
			sourceSheet = sourceWorkbook.getSheet(source.getSheetName());
		}
	}

	public final void setTargetSheet(SheetIdentity target) throws ExcelException {
		if (isTargetSet) {
			throw new RuntimeException("target sheet has been set");
		}

		if (target == null) {
			throw new NullPointerException("target sheet can not be null");
		}

		if (source != null && source.equals(target)) {
			throw new IllegalArgumentException("target can not same as source");
		}

		this.target = target;

		final String targetPath = target.getWorkbookName();
		if (source != null && source.getWorkbookName().equals(targetPath)) {
			targetWorkbook = sourceWorkbook;
		} else {
			File file = new File(target.getWorkbookName());
			if (file.exists()) {
				targetWorkbook = PoiUtils.getWorkbookByPath(targetPath);
			} else {
				targetWorkbook = PoiUtils.createWorkbook(targetPath);
			}
		}

		targetSheet = targetWorkbook.getSheet(target.getSheetName());
		if (targetSheet == null) {
			targetSheet = targetWorkbook.createSheet(target.getSheetName());
		}

		isTargetSet = true;
	}

	public final void setParamPath(String paramPath) throws BadParamException {
		if (isParamSet) {
			throw new RuntimeException("target sheet has been set");
		}

		if (StringUtils.isEmpty(paramPath)) {
			throw new NullPointerException("paramPath can not be empty");
		}

		this.paramPath = paramPath;
		loadParam();

		isParamSet = true;
	}

	private void checkSourceSheet() {
		if (source == null) {
			throw new IllegalStateException("source has not been set");
		}
	}

	private void checkTargetSheet() {
		if (target == null) {
			throw new IllegalStateException("target has not been set");
		}
	}

	private void checkParamPath() {
		if (paramPath == null || "".equals(paramPath)) {
			throw new IllegalStateException("params have not been set");
		}
	}

	abstract void loadParam() throws BadParamException;

	abstract void work();

	public final void execute() throws ExcelException {
		checkSourceSheet();
		checkTargetSheet();
		checkParamPath();
		work();
	}

	public void store() throws ExcelException {
		checkTargetSheet();
		PoiUtils.writeWorkbookToFile(targetWorkbook, target.getWorkbookName());
	}
}
