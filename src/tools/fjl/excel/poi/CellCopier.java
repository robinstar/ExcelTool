package tools.fjl.excel.poi;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

public class CellCopier {

	public static void copy(Workbook targetBook, Cell target, Cell source, FormulaEvaluator evaluator) {
		switch (source.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			break;

		case Cell.CELL_TYPE_BOOLEAN:
			target.setCellValue(source.getBooleanCellValue());
			break;

		case Cell.CELL_TYPE_ERROR:
			target.setCellValue(source.getErrorCellValue());
			break;

		case Cell.CELL_TYPE_FORMULA:
			copyFormulaValue(target, evaluator.evaluate(source));
			break;

		case Cell.CELL_TYPE_NUMERIC:
			copyNumericValue(targetBook, target, source);
			copyNumericStyle(targetBook, target, source);
			break;

		case Cell.CELL_TYPE_STRING:
			target.setCellValue(source.getStringCellValue());
			break;

		default:
			break;
		}
	}

	private static void copyNumericValue(Workbook targetBook, Cell target, Cell source) {
		if (HSSFDateUtil.isCellDateFormatted(source)) {
			Date date = source.getDateCellValue();

			Calendar cal = Calendar.getInstance();
			final long currentTime = cal.getTimeInMillis();
			final int thisYear = cal.get(Calendar.YEAR);
			cal.setTime(date);
			final int sourceYear = cal.get(Calendar.YEAR);
			if (sourceYear < 1970) {
				cal.set(Calendar.YEAR, thisYear);
				final long time = cal.getTimeInMillis();
				if (time > currentTime) {
					cal.set(Calendar.YEAR, thisYear - 1);
				}
			}

			date = cal.getTime();
			target.setCellValue(date);
		} else {
			target.setCellValue(source.getNumericCellValue());
		}
	}

	private static void copyNumericStyle(Workbook targetBook, Cell target, Cell source) {
		final CellStyle sourceStyle = source.getCellStyle();
		final String sourceFormat = sourceStyle.getDataFormatString();

		final short targetFormatIndex;
		if (sourceFormat == null || "".equals(sourceFormat)) {
			final short sourceFormatIndex = sourceStyle.getDataFormat();
			targetFormatIndex = sourceFormatIndex;
		} else {
			DataFormat targetDataFormat = targetBook.createDataFormat();
			targetFormatIndex = targetDataFormat.getFormat(sourceFormat);
		}

		CellStyle targetStyle = targetBook.createCellStyle();
		targetStyle.setDataFormat(targetFormatIndex);

		target.setCellStyle(targetStyle);
	}

	public static void copyFormulaValue(Cell target, CellValue source) {
		switch (source.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN:
			target.setCellValue(source.getBooleanValue());
			break;

		case Cell.CELL_TYPE_ERROR:
			target.setCellValue(source.getErrorValue());
			break;

		case Cell.CELL_TYPE_NUMERIC:
			target.setCellValue(source.getNumberValue());
			break;

		case Cell.CELL_TYPE_STRING:
			target.setCellValue(source.getStringValue());
			break;

		default:
			break;
		}
	}

}
