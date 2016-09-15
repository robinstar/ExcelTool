package tools.fjl.excel;

import java.io.File;

public class Utils {

	public static boolean isExcelFile(String path) {
		File file = new File(path);
		String fileName = file.getName();
		return fileName.matches("^[^(~$)][\\s\\S]*(.xls|.xlsx)$");
	}

	public static boolean isXlsFile(String path) {
		File file = new File(path);
		String fileName = file.getName();
		return fileName.matches("^[^(~$)][\\s\\S]*(.xls)$");
	}

	public static boolean isXlsxFile(String path) {
		File file = new File(path);
		String fileName = file.getName();
		return fileName.matches("^[^(~$)][\\s\\S]*(.xlsx)$");
	}
}