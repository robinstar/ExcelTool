package tools.fjl.excel;

public interface ExcelTool {

	public void setSourceSheet(SheetIdentity source) throws ExcelException;

	public void setTargetSheet(SheetIdentity target) throws ExcelException;

	public void setParamPath(String paramPath) throws BadParamException;

	public void execute() throws ExcelException;

	public void store() throws ExcelException;
}
