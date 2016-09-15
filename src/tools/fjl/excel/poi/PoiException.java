package tools.fjl.excel.poi;

import tools.fjl.excel.ExcelException;

public class PoiException extends ExcelException {

	private static final long serialVersionUID = 1L;

	public PoiException() {
		super();
	}

	public PoiException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	public PoiException(String message, Throwable cause) {
		super(message, cause);
	}

	public PoiException(String message) {
		super(message);
	}

	public PoiException(Throwable cause) {
		super(cause);
	}
}
