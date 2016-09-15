package tools.fjl.excel.poi;

import org.junit.Test;

import tools.fjl.excel.BadParamException;

public class CellListWorkerTest {

	private static final String PARAM_NAME = "/cellListParam.txt";
	CellListWorker worker = new CellListWorker();
	
	@Test
	public void testParam() {
		String path = getClass().getResource(PARAM_NAME).getPath();
		try {
			worker.setParamPath(path);
			
			assert worker != null;
		} catch (BadParamException e) {
			e.printStackTrace();
		}
	}
}
