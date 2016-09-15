package tools.fjl.excel.poi;

import static tools.fjl.excel.CellIndex.UNDIFINED;
import static tools.fjl.excel.CellIndex.Regex.COLUMN;
import static tools.fjl.excel.CellIndex.Regex.ROW;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.Reader;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import com.robin.util.StringUtils;

import tools.fjl.excel.CellIndex;
import tools.fjl.excel.BadParamException;

public class MultiRowWorker extends BaseTool {

	static final String HYPHEN_ROW_REGEX = "(" + ROW + ")(-(" + ROW + "))?";
	static final String COMMA_ROW_REGEX = getCommaRegex(HYPHEN_ROW_REGEX);

	static final String UNDEFINED_CELL_REGEX = "[a-zA-Z]{1,3}(" + ROW + ")?";
	static final String COMMA_COLUMN_REGEX = getCommaRegex(UNDEFINED_CELL_REGEX);

	static final String CHECK_COLUMN_REGEX = getCommaRegex(COLUMN);

	private static String getCommaRegex(String singleRegex) {
		return "(" + singleRegex + ")|((" + singleRegex + "),)+" + "|(,(" + singleRegex + "))+" + "|,((" + singleRegex
				+ "),)+" + "|(" + singleRegex + ")(,(" + singleRegex + "))+";
	}

	ArrayList<Integer> rowIndexs = new ArrayList<Integer>();
	ArrayList<CellIndex> cells = new ArrayList<CellIndex>();
	List<Integer> checkedColumnIndexs = new ArrayList<Integer>();

	@Override
	void loadParam() throws BadParamException {
		Reader reader = null;
		BufferedReader br = null;
		try {
			reader = new FileReader(paramPath);
			br = new BufferedReader(reader);

			String rowParamLine;
			String columnParamLine;
			String checkParamLine;

			rowParamLine = br.readLine();
			if (rowParamLine == null || !rowParamLine.matches(COMMA_ROW_REGEX)) {
				throw new BadParamException("第一行参数格式错误");
			}

			columnParamLine = br.readLine();
			if (columnParamLine == null || !columnParamLine.matches(COMMA_COLUMN_REGEX)) {
				throw new BadParamException("第二行参数格式错误");
			}

			checkParamLine = br.readLine();
			if (!StringUtils.isEmpty(checkParamLine) && !checkParamLine.matches(CHECK_COLUMN_REGEX)) {
				throw new BadParamException("第三行参数格式错误");
			}

			parseRowParam(rowParamLine);
			parseCells(columnParamLine);
			parseCheckColumn(checkParamLine);

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

	void parseRowParam(String paramLine) {
		rowIndexs.clear();
		Set<Integer> paramsSet = new HashSet<Integer>();

		String[] params = paramLine.split(",");
		for (String hyphenParam : params) {
			if (StringUtils.isEmpty(hyphenParam)) {
				continue;
			}

			if (hyphenParam.matches(ROW)) {
				paramsSet.add(Integer.valueOf(hyphenParam) - 1);
			} else if (hyphenParam.matches(HYPHEN_ROW_REGEX)) {
				String[] be = hyphenParam.split("-");
				int begin = Integer.valueOf(be[0]);
				int end = Integer.valueOf(be[1]);

				for (; begin <= end; begin++) {
					paramsSet.add(begin - 1);
				}
			} else {
				continue;
			}
		}

		rowIndexs.addAll(paramsSet);
		Collections.sort(rowIndexs);
	}

	void parseCells(String paramLine) {
		cells.clear();

		String[] params = paramLine.split(",");
		for (String param : params) {
			if (StringUtils.isEmpty(param)) {
				continue;
			}

			cells.add(new CellIndex(param));
		}
	}

	void parseCheckColumn(String paramLine) {
		checkedColumnIndexs.clear();
		String[] params = paramLine.split(",");
		for (String param : params) {
			if (StringUtils.isEmpty(param)) {
				continue;
			}

			int columnIndex = CellIndex.columnToIndex(param);
			checkedColumnIndexs.add(columnIndex);
		}
	}

	@Override
	void work() {
		if (sourceSheet == null) {
			String errorLog = String.format("there is no such shell  %s in %s", source.getSheetName(),
					source.getWorkbookName());
			System.err.println(errorLog);
			return;
		}

		FormulaEvaluator evaluator;
		try {
			evaluator = PoiUtils.getFormulaEvaluator(sourceWorkbook);
		} catch (PoiException e) {
			return; // TODO
		}

		File file = new File(source.getWorkbookName());
		final String simpleName = file.getName();

		final int checkedColumn = checkedColumnIndexs.size() > 0 ? checkedColumnIndexs.get(0) : UNDIFINED;

		for (int rowIndex : rowIndexs) {
			Row row = sourceSheet.getRow(rowIndex);
			if (row == null)
				continue;

			if (checkedColumn != UNDIFINED) {
				Cell checkedCell = row.getCell(checkedColumn);
				if (checkedCell == null)
					continue;

				if (checkedCell.getCellType() != Cell.CELL_TYPE_STRING)
					continue;
				String cellValue = checkedCell.getStringCellValue();
				if (StringUtils.isEmpty(cellValue))
					continue;
			}

			int lastRowIndexInTarget = targetSheet.getLastRowNum();
			final int targetRowIndex = lastRowIndexInTarget + 1;
			Row targetRow = targetSheet.createRow(targetRowIndex);
			int targetColumnIndex = 0;
			targetRow.createCell(targetColumnIndex++).setCellValue(simpleName);

			for (CellIndex cell : cells) {
				Cell sourceCell = null;
				Cell targetCell = null;
				String error = null;
				if (cell.row != UNDIFINED && cell.column != UNDIFINED) {
					Row cellRow = sourceSheet.getRow(cell.row);
					if (cellRow == null) {
						String errorLog = String.format("there is no such row in %s of %s, row : %d",
								source.getSheetName(), source.getWorkbookName(), cell.row);
						System.err.println(errorLog);
						error = "单元格缺失";
					} else {
						sourceCell = cellRow.getCell(cell.column);
						if (sourceCell == null) {
							String errorLog = String.format("there is no such column in %s of %s, column : %d",
									source.getSheetName(), source.getWorkbookName(), cell.column);
							System.err.println(errorLog);
							error = "单元格缺失";
						}
					}

					targetCell = targetRow.createCell(targetColumnIndex++);
					if (sourceCell != null) {
						CellCopier.copy(targetCell, sourceCell, evaluator);
					} else {
						targetCell.setCellValue(error);
					}
				} else if (cell.row == UNDIFINED && cell.column != UNDIFINED) {
					sourceCell = row.getCell(cell.column);
					if (sourceCell == null) {
						String errorLog = String.format("there is no such column in %s of %s, column : %d",
								source.getSheetName(), source.getWorkbookName(), cell.column);
						System.err.println(errorLog);
						error = "单元格缺失";
					}

					targetCell = targetRow.createCell(targetColumnIndex++);
					if (sourceCell != null) {
						CellCopier.copy(targetCell, sourceCell, evaluator);
					} else {
						targetCell.setCellValue(error);
					}
				} else {
					continue;
				}
			}
		}
	}
}
