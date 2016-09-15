package tools.fjl.excel;

import java.util.Set;

public class WorkCluster implements Runnable {

	private Set<SheetIdentity> sourceSheets;
	private SheetIdentity targetSheet;
	private String paramPath;

	private ExcelTool excelTool;
	boolean exception = false;

	public interface Callback {
		void preExecute(WorkCluster worker);

		void onError(WorkCluster worker);

		void PostExecute(WorkCluster worker);
	}

	private Callback callback;

	public void setCallback(Callback callback) {
		this.callback = callback;
	}

	public void setSourceSheets(Set<SheetIdentity> sourceSheets) {
		this.sourceSheets = sourceSheets;
	}

	public void setTargetSheet(SheetIdentity targetSheet) {
		this.targetSheet = targetSheet;
	}

	public void setParamPath(String paramPath) {
		this.paramPath = paramPath;
	}

	public void setExcelTool(ExcelTool tool) {
		this.excelTool = tool;
	}

	private void checkSourceSheets() throws SetupException {
		if (sourceSheets == null) {
			exception = true;
			throw new SetupException("source sheet set has not been set");
		}
	}

	private void checkTargetSheet() throws SetupException {
		if (targetSheet == null) {
			exception = true;
			throw new SetupException("target sheet set has not been set");
		}
	}

	private void checkWorkParams() throws SetupException {
		if (paramPath == null || "".equals(paramPath)) {
			exception = true;
			throw new SetupException("param path has not been set");
		}
	}

	private void checkExcelTool() throws SetupException {
		if (excelTool == null) {
			exception = true;
			throw new SetupException("ExcelTool has not been set");
		}
	}

	public void run() {
		onStart();
		exception = false;

		try {
			checkSourceSheets();
			checkTargetSheet();
			checkWorkParams();
			checkExcelTool();
		} catch (SetupException e) {
			onError(e);
			return;
		}

		try {
			excelTool.setParamPath(paramPath);
		} catch (BadParamException e) {
			onError(e);
			return;
		}

		try {
			excelTool.setTargetSheet(targetSheet);
		} catch (Exception e) {
			String msg = String.format("catch an exception when set tarege :%s; %s", targetSheet.toString(),
					e.getMessage());
			onError(new RuntimeException(msg));
			return;
		}

		for (SheetIdentity source : sourceSheets) {
			try {
				excelTool.setSourceSheet(source);
			} catch (Exception e) {
				String msg = String.format("catch an exception when set source :%s; %s", source.toString(),
						e.getMessage());
				Logger.log(msg);
				Logger.log(e);
				continue;
			}

			try {
				excelTool.execute();
			} catch (Exception e) {
				String msg = String.format("catch an exception when execute :%s; %s", source.toString(),
						e.getMessage());
				Logger.log(msg);
				Logger.log(e);
				continue;
			}
		}

		try {
			excelTool.store();
		} catch (Exception e) {
			String msg = String.format("catch an exception when store; %s", e.getMessage());
			onError(new RuntimeException(msg));
			return;
		}

		onStop();
	}

	private void onStart() {
		logStart();

		if (callback != null) {
			callback.preExecute(this);
		}
	}

	private void onError(Exception e) {
		Logger.log(e);

		if (callback != null) {
			callback.onError(this);
		}

		logStop();
	}

	private void onStop() {
		if (callback != null) {
			callback.PostExecute(this);
		}

		logStop();
	}

	private void logStart() {
		Logger.log("");
		Logger.log("--------begin--------");
	}

	private void logStop() {
		Logger.log("--------end--------");
	}
}
