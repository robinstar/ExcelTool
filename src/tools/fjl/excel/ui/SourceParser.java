package tools.fjl.excel.ui;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import tools.fjl.excel.ExcelException;
import tools.fjl.excel.Logger;
import tools.fjl.excel.SheetIdentity;
import tools.fjl.excel.Utils;
import tools.fjl.excel.poi.PoiUtils;

public class SourceParser {
	private static final String EXCLUDE_SHEET_NAME_REGEX = "^Sheet\\d+$";

	private final Map<String, HashSet<SheetIdentity>> sheets = new HashMap<String, HashSet<SheetIdentity>>();
	private final List<String> sheetNames = new ArrayList<String>();
	private final List<String> workbookPathes = new ArrayList<String>();

	private String selectedSheetName;
	private boolean isParsing = false;

	private Callback callback;

	public interface Callback {
		void preParse();

		void onParseError(String msg);

		void onParseDone(SourceParser parser);
	}

	public synchronized void parse(String folder) {
		if (isParsing) {
			throw new IllegalStateException("isParsing");
		}

		if (folder == null || folder.equals("")) {
			throw new IllegalArgumentException("empty folder!");
		}

		isParsing = true;
		if (callback != null) {
			callback.preParse();
		}

		clearSheets();

		try {
			parseSheets(folder);
		} catch (Exception e) {
			Logger.log(e);

			if (callback != null) {
				callback.onParseError(e.getMessage());
			}

			isParsing = false;
			return;
		}

		selectDefaultSheet();
		isParsing = false;

		if (callback != null) {
			callback.onParseDone(this);
		}
	}

	public void setCallback(Callback callback) {
		this.callback = callback;
	}

	private void clearSheets() {
		Logger.log("SourceParser clearSheets");
		for (Entry<String, HashSet<SheetIdentity>> entry : sheets.entrySet()) {
			entry.getValue().clear();
		}
		sheets.clear();
		sheetNames.clear();
	}

	private void parseSheets(String folder) throws ExcelException {
		Logger.log("parseSheets folder :" + folder);
		File root = new File(folder);
		if (!root.exists()) {
			throw new IllegalArgumentException("no such file :" + folder);
		}

		if (!root.isDirectory()) {
			throw new IllegalArgumentException(folder + " mush be a folder");
		}

		File[] childrenFile = root.listFiles();
		workbookPathes.clear();

		for (File file : childrenFile) {
			String path = file.getAbsolutePath();
			if (Utils.isExcelFile(path)) {
				Logger.log("parseSheets Workbook path :" + path);
				workbookPathes.add(path);
				Workbook book = PoiUtils.getWorkbookByPath(path);
				for (Sheet sheet : book) {
					String sheetName = sheet.getSheetName();
					if (sheetName.matches(EXCLUDE_SHEET_NAME_REGEX)) {
						continue;
					}

					HashSet<SheetIdentity> namedSheetSet = sheets.get(sheetName);
					if (namedSheetSet == null) {
						namedSheetSet = new HashSet<SheetIdentity>();
						sheets.put(sheetName, namedSheetSet);
					}

					SheetIdentity sheetId = new SheetIdentity(path, sheetName);
					namedSheetSet.add(sheetId);
				}
			}
		}
	}

	private void selectDefaultSheet() {
		for (Entry<String, HashSet<SheetIdentity>> entry : sheets.entrySet()) {
			final String sheetName = entry.getKey();
			final int sheetCount = entry.getValue().size();

			final int N = sheetNames.size();
			int targetIndex;
			for (targetIndex = 0; targetIndex < N; targetIndex++) {
				String tempSheetName = sheetNames.get(targetIndex);
				int tempSheetCount = sheets.get(tempSheetName).size();
				if (sheetCount < tempSheetCount) {
					continue;
				} else {
					break;
				}
			}
			sheetNames.add(targetIndex, sheetName);
		}

		sheetNames.add(0, SheetIdentity.FIRST_SHEET);
		selectedSheetName = sheetNames.size() == 0 ? null : sheetNames.get(0);
		Logger.log("selectDefaultSheet sheet name :" + selectedSheetName);
	}

	public synchronized void setSelectedSheetName(String sheetName) {
		if (!sheetNames.contains(sheetName)) {
			throw new IllegalArgumentException();
		}
		this.selectedSheetName = sheetName;
	}

	public synchronized String getSelectedSheetName() {
		return selectedSheetName;
	}

	public synchronized Set<SheetIdentity> getSelectedSheets() {
		HashSet<SheetIdentity> result = new HashSet<SheetIdentity>();
		if (sheets.get(selectedSheetName) != null) {
			result.addAll(sheets.get(selectedSheetName));
		}
		return result;
	}

	public synchronized List<String> getSheetNames() {
		return new ArrayList<String>(sheetNames);
	}

	public synchronized List<String> getWorkbookPathes() {
		return new ArrayList<String>(workbookPathes);
	}

	Map<String, HashSet<SheetIdentity>> getSheets() {
		return sheets;
	}
}
