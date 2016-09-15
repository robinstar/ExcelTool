package tools.fjl.excel.poi;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.Reader;
import java.util.ArrayList;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import tools.fjl.excel.BadParamException;
import tools.fjl.excel.CellIndex;
import tools.fjl.excel.Logger;

public class CellListWorker extends BaseTool {

	static final String PARAM_REGEX = "^(" + CellIndex.Regex.CELL + ")$";
	static final String HYPHEN_PARAM_REGEX = "(" + CellIndex.Regex.CELL + ")(-(" + CellIndex.Regex.CELL + "))?";

	static final Pattern COLUME_PATTERN = Pattern.compile("^(" + CellIndex.Regex.COLUMN + ")");

	final ArrayList<CellIndex> cellsForCopy = new ArrayList<CellIndex>();

	@Override
	void loadParam() throws BadParamException {
		Reader reader = null;
		BufferedReader br = null;
		try {
			reader = new FileReader(paramPath);
			br = new BufferedReader(reader);
			String line;
			while ((line = br.readLine()) != null) {
				if (line.matches(PARAM_REGEX)) {
					CellIndex index = new CellIndex(line);
					cellsForCopy.add(index);
				} else if (line.matches(HYPHEN_PARAM_REGEX)) {
					String[] temp = line.split("-");
					CellIndex lt = new CellIndex(temp[0]);
					CellIndex rb = new CellIndex(temp[1]);
					for (int i = lt.row; i <= rb.row; i++) {
						for (int j = lt.column; j <= rb.column; j++) {
							CellIndex index = new CellIndex(i, j);
							cellsForCopy.add(index);
						}
					}
				}
			}

			if (cellsForCopy.isEmpty()) {
				throw new BadParamException();
			}
		} catch (IOException e) {
			throw new BadParamException();
		} finally {
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}

			if (reader != null) {
				try {
					reader.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	@Override
	void work() {
		FormulaEvaluator evaluator;
		try {
			evaluator = PoiUtils.getFormulaEvaluator(sourceWorkbook);
		} catch (PoiException e) {
			return; // TODO
		}

		int lastRowIndexInTarget = targetSheet.getLastRowNum();
		final int targetRowIndex = lastRowIndexInTarget + 1;
		Row targetRow = targetSheet.createRow(targetRowIndex);
		File file = new File(source.getWorkbookName());
		final String simpleName = file.getName();
		Logger.log(simpleName);

		int targetColumnIndex = 0;
		targetRow.createCell(targetColumnIndex++).setCellValue(simpleName);

		if (sourceSheet == null) {
			String errorLog = String.format("there is no such shell  %s in %s", source.getSheetName(),
					source.getWorkbookName());
			System.err.println(errorLog);
			String error = "没有对应的工作表";
			targetRow.createCell(targetColumnIndex).setCellValue(error);
			Logger.log(error);
			return;
		}

		for (CellIndex cellIndex : cellsForCopy) {
			Cell sourceCell = null;
			Cell targetCell = null;
			String error = null;
			Row sourceRow = sourceSheet.getRow(cellIndex.row);
			if (sourceRow == null) {
				String errorLog = String.format("there is no such row in %s of %s, row : %d", source.getSheetName(),
						source.getWorkbookName(), cellIndex.row);
				System.err.println(errorLog);
				error = "单元格缺失";
			} else {
				sourceCell = sourceRow.getCell(cellIndex.column);
				if (sourceCell == null) {
					String errorLog = String.format("there is no such column in %s of %s, column : %d",
							source.getSheetName(), source.getWorkbookName(), cellIndex.column);
					System.err.println(errorLog);
					error = "单元格缺失";
				}
			}

			if (sourceCell != null && sourceCell.getCellType() == Cell.CELL_TYPE_BLANK) {
				continue;
			}

			targetCell = targetRow.createCell(targetColumnIndex++);
			if (sourceCell != null) {
				CellCopier.copy(targetWorkbook, targetCell, sourceCell, evaluator);
			} else {
				targetCell.setCellValue(error);
			}
		}
	}
}
