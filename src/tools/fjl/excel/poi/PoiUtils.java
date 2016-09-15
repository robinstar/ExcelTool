package tools.fjl.excel.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import tools.fjl.excel.ExcelException;
import tools.fjl.excel.Logger;
import tools.fjl.excel.Utils;

public class PoiUtils {

	public static Workbook getWorkbookByPath(String path) throws ExcelException {
		InputStream is = null;
		try {
			if (Utils.isXlsFile(path)) {
				is = new FileInputStream(path);
				return new HSSFWorkbook(is);

			} else if (Utils.isXlsxFile(path)) {
				is = new FileInputStream(path);
				return new XSSFWorkbook(is);
			} else {
				throw new ExcelException("path must be a '.xls' or '.xlsx' file path");
			}
		} catch (IOException e) {
			Logger.log(path);
			throw new PoiException(e);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	public static Workbook createWorkbook(String path) throws ExcelException {
		if (Utils.isXlsFile(path)) {
			return new HSSFWorkbook();
		} else if (Utils.isXlsxFile(path)) {
			return new XSSFWorkbook();
		} else {
			throw new ExcelException("path must be a '.xls' or '.xlsx' file path");
		}
	}

	public static FormulaEvaluator getFormulaEvaluator(Workbook book) throws PoiException {
		if (book == null) {
			return null;
		}

		if (book instanceof HSSFWorkbook) {
			return new HSSFFormulaEvaluator((HSSFWorkbook) book);
		} else if (book instanceof XSSFWorkbook) {
			return new XSSFFormulaEvaluator((XSSFWorkbook) book);
		} else {
			throw new PoiException("unsupport Workbook type");
		}
	}

	public static void writeWorkbookToFile(Workbook workbook, String path) throws PoiException {
		OutputStream os = null;
		try {
			os = new FileOutputStream(path);
			workbook.write(os);
		} catch (IOException e) {
			throw new PoiException(e);
		} finally {
			if (os != null) {
				try {
					os.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
