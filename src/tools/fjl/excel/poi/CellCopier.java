package tools.fjl.excel.poi;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class CellCopier {

	public static void copy(Cell target, Cell source, FormulaEvaluator evaluator) {
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
			if (HSSFDateUtil.isCellDateFormatted(source)) {
				Date date = source.getDateCellValue();
				Calendar cal = Calendar.getInstance();
				final int thisYear = cal.get(Calendar.YEAR);
				final long currentTime = cal.getTimeInMillis();

				cal.setTime(date);
				final int year = cal.get(Calendar.YEAR);
				if (year < 1970) {
					cal.set(Calendar.YEAR, thisYear);
					final long time = cal.getTimeInMillis();
					if (time > currentTime) {
						cal.set(Calendar.YEAR, thisYear - 1);
					}
				}
				date = cal.getTime();

				target.setCellValue(date);
				target.setCellType(source.getCellType());
				int type = target.getCellType();
				assert type == 0;
			} else {
				target.setCellValue(source.getNumericCellValue());
			}
			break;

		case Cell.CELL_TYPE_STRING:
			target.setCellValue(source.getStringCellValue());
			break;

		default:
			break;
		}
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
