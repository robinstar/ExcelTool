package tools.fjl.excel;

import java.util.HashSet;
import java.util.Set;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

public class WorkClusterTest {

	private Set<SheetIdentity> sheets;
	private SheetIdentity target;
	private String paramPath;
	private ExcelTool tool = new ExcelTool() {

		public void store() {
		}

		public void setTargetSheet(SheetIdentity target) {
		}

		public void setSourceSheet(SheetIdentity source) {
		}

		public void setParamPath(String paramPath) throws BadParamException {
		}

		public void execute() {
		}
	};

	private WorkCluster cluster;

	@Before
	public void setUp() {
		cluster = new WorkCluster();

		sheets = new HashSet<SheetIdentity>();
		String path = getClass().getResource("/list1.xlsx").getPath();
		SheetIdentity sheet = new SheetIdentity(path, "copyData");
		sheets.add(sheet);

		path = getClass().getResource("/list2.xlsx").getPath();
		sheet = new SheetIdentity(path, "copyData");
		sheets.add(sheet);

		String targetPath = getClass().getResource("/target.xls").getPath();
		target = new SheetIdentity(targetPath, "target");

		paramPath = getClass().getResource("/cellListParam.txt").getPath();
	}

	@Test
	public void testNullSourceSheets() {
		cluster.setSourceSheets(null);
		cluster.setTargetSheet(target);
		cluster.setParamPath(paramPath);
		cluster.setExcelTool(tool);
		cluster.run();
		Assert.assertTrue(cluster.exception);
	}

	@Test
	public void testNullTargetSheet() {
		cluster.setSourceSheets(sheets);
		cluster.setTargetSheet(null);
		cluster.setParamPath(paramPath);
		cluster.setExcelTool(tool);
		cluster.run();
		Assert.assertTrue(cluster.exception);
	}

	@Test
	public void testNullParamPath() {
		cluster.setSourceSheets(sheets);
		cluster.setTargetSheet(target);
		cluster.setParamPath(null);
		cluster.setExcelTool(tool);
		cluster.run();
		Assert.assertTrue(cluster.exception);
	}

	@Test
	public void testEmptyParamPath() {
		cluster.setSourceSheets(sheets);
		cluster.setTargetSheet(target);
		cluster.setParamPath("");
		cluster.setExcelTool(tool);
		cluster.run();
		Assert.assertTrue(cluster.exception);
	}

	@Test
	public void testNullToolPath() {
		cluster.setSourceSheets(sheets);
		cluster.setTargetSheet(target);
		cluster.setParamPath(paramPath);
		cluster.setExcelTool(null);
		cluster.run();
		Assert.assertTrue(cluster.exception);
	}

	@Test
	public void testWork() {
		cluster.setSourceSheets(sheets);
		cluster.setTargetSheet(target);
		cluster.setParamPath(paramPath);
		cluster.setExcelTool(tool);
		cluster.run();
		Assert.assertFalse(cluster.exception);
	}
}
