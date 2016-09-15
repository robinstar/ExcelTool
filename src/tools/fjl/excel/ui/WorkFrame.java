package tools.fjl.excel.ui;

import static tools.fjl.excel.ui.Constants.CONTENT_WIDTH;
import static tools.fjl.excel.ui.Constants.FRAME_WIDTH;
import static tools.fjl.excel.ui.Constants.GROUP_GAP;
import static tools.fjl.excel.ui.Constants.ITEM_HEIGHT;
import static tools.fjl.excel.ui.Constants.PADDING_BOTTOM;
import static tools.fjl.excel.ui.Constants.PADDING_LEFT;
import static tools.fjl.excel.ui.Constants.PADDING_TOP;
import static tools.fjl.excel.ui.Constants.TITLE;

import java.awt.Container;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.io.IOException;
import java.net.URL;
import java.util.HashSet;
import java.util.List;
import java.util.Properties;
import java.util.Set;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import com.robin.util.PropUtils;
import com.robin.util.StringUtils;

import tools.fjl.excel.ExcelTool;
import tools.fjl.excel.SheetIdentity;
import tools.fjl.excel.WorkCluster;
import tools.fjl.excel.WorkCluster.Callback;
import tools.fjl.excel.ui.SourceDataPanel.SourceParseStateListener;

final class WorkFrame extends JFrame implements ActionListener, Callback {

	private static final long serialVersionUID = 1L;

	private Dimension screenSize;
	private Dimension frameDimention;
	private String locationPropPath;
	private Properties locationProperties;

	private SourceDataPanel sourceDataPanel;
	private TargetFilePanel targetFilePanel;
	private OperationPanel operationPanel;
	private JButton startButton;

	WorkFrame() {
		super(TITLE);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		screenSize = Toolkit.getDefaultToolkit().getScreenSize();

		// content component
		setLayout(null);
		Container container = getContentPane();

		int top = PADDING_TOP;
		Rectangle rectangle = new Rectangle();
		rectangle.setLocation(PADDING_LEFT, top);
		rectangle.setSize(CONTENT_WIDTH, 0);
		sourceDataPanel = new SourceDataPanel(rectangle);
		container.add(sourceDataPanel);

		top = top + sourceDataPanel.getHeight() + GROUP_GAP;
		rectangle = new Rectangle();
		rectangle.setLocation(PADDING_LEFT, top);
		rectangle.setSize(CONTENT_WIDTH, 0);
		targetFilePanel = new TargetFilePanel(this, rectangle);
		container.add(targetFilePanel);

		top = top + targetFilePanel.getHeight() + GROUP_GAP;
		rectangle = new Rectangle();
		rectangle.setLocation(PADDING_LEFT, top);
		rectangle.setSize(CONTENT_WIDTH, 0);
		operationPanel = new OperationPanel(rectangle);
		container.add(operationPanel);

		top = top + operationPanel.getHeight() + GROUP_GAP;
		int buttonWidth = 100;
		int buttonX = (FRAME_WIDTH - buttonWidth) / 2;
		startButton = new JButton("开始");
		startButton.setBounds(buttonX, top, buttonWidth, ITEM_HEIGHT);
		startButton.addActionListener(this);
		container.add(startButton);

		// width height
		int frameHeight;
		frameHeight = top + startButton.getHeight() + PADDING_BOTTOM;
		frameHeight += 30; // title height
		frameDimention = new Dimension(FRAME_WIDTH, frameHeight);
		checkFrameSize();
		setSize(frameDimention);
		setResizable(false);

		// location
		try {
			URL url = getClass().getResource("location.properties");
			locationPropPath = url.getPath().toString();
			locationProperties = PropUtils.read(locationPropPath);
			String propX = locationProperties.getProperty("x");
			String propY = locationProperties.getProperty("y");
			int x = StringUtils.isEmpty(propX) ? 0 : Integer.valueOf(propX);
			int y = StringUtils.isEmpty(propY) ? 0 : Integer.valueOf(propY);
			removeComponentListener(movedFrameAdapter);
			setLocation(x, y);
			addComponentListener(movedFrameAdapter);
		} catch (IOException e) {
			e.printStackTrace();
		}

		// listener
		sourceDataPanel.setSourceParseStateListener(new SourceParseStateListener() {
			public void onSourceParseStateChanged(int state, String msg) {
				if (state == SourceParseStateListener.STATE_ERROR) {
					alertMessage(msg);
				}
				startButton.setEnabled(state != SourceParseStateListener.STATE_PARSING);
			}
		});
	}

	private void checkFrameSize() {
		if (frameDimention.getWidth() > screenSize.getWidth() || frameDimention.getHeight() > screenSize.getHeight()) {
			System.exit(0);
		}
	}

	private ComponentAdapter movedFrameAdapter = new ComponentAdapter() {
		@Override
		public void componentMoved(ComponentEvent e) {
			adjustFrameLocation();
			recordFrameLocation();
		}
	};

	private void adjustFrameLocation() {
		final int x = getX();
		final int y = getY();

		int adjustX = Math.max(0, x);
		int adjustY = Math.max(0, y);

		int maxX = (int) (screenSize.getWidth() - frameDimention.getWidth());
		int maxY = (int) (screenSize.getHeight() - frameDimention.getHeight());
		adjustX = Math.min(adjustX, maxX);
		adjustY = Math.min(adjustY, maxY);

		if (adjustX != x || adjustY != y) {
			removeComponentListener(movedFrameAdapter);
			setLocation(adjustX, adjustY);
			addComponentListener(movedFrameAdapter);
		}
	}

	private void recordFrameLocation() {
		locationProperties.setProperty("x", getX() + "");
		locationProperties.setProperty("y", getY() + "");
		try {
			PropUtils.write(locationPropPath, locationProperties);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void actionPerformed(ActionEvent e) {
		if (e.getSource().equals(startButton)) {
			if (!sourceDataPanel.isSourceSelected()) {
				alertMessage(Error.SOURCE_NOT_SELECTED);
				return;
			}

			if (!sourceDataPanel.hasAvaiableSourceBooks()) {
				alertMessage(Error.SOURCE_FILES_NOT_EXIST);
				return;
			}

			if (!sourceDataPanel.hasAvaialbeSourceSheets()) {
				alertMessage(Error.NO_AVAIALBE_SHEET);
				return;
			}

			if (!targetFilePanel.isTargetWorkbookSelected()) {
				alertMessage(Error.TARGET_NOT_SELECTED);
				return;
			}

			if (!targetFilePanel.isTargetSheetNameSelected()) {
				alertMessage(Error.TARGET_SHEET_NOT_SELECTED);
				return;
			}

			if (!operationPanel.isParamPathSelected()) {
				alertMessage(Error.PARAM_NOT_SELECTED);
				return;
			}

			WorkCluster cluster = new WorkCluster();

			String targetPath = targetFilePanel.getTargetPath();
			String targetSheet = targetFilePanel.getTargetSheetName();
			SheetIdentity target = new SheetIdentity(targetPath, targetSheet);
			cluster.setTargetSheet(target);

			List<String> allBooks = sourceDataPanel.getAllWorkbooks();
			if (!targetFilePanel.includeTargetWorkbook()) {
				allBooks.remove(targetPath);
			}

			final String sheetName = sourceDataPanel.getSourceSheetName();
			Set<SheetIdentity> sourceSheets = new HashSet<SheetIdentity>();
			for (String bookName : allBooks) {
				SheetIdentity s = new SheetIdentity(bookName, sheetName);
				sourceSheets.add(s);
			}
			cluster.setSourceSheets(sourceSheets);

			cluster.setParamPath(operationPanel.getParamPath());
			ExcelTool tool;
			try {
				tool = (ExcelTool) Class.forName(operationPanel.getOperation().getJavaClass()).newInstance();
			} catch (Exception ex) {
				alertMessage("功能开发中，请耐心等待～");
				return;
			}

			cluster.setExcelTool(tool);

			cluster.setCallback(this);
			cluster.run();
		}
	}

	static final class Error {
		static final String SOURCE_NOT_SELECTED = "还未指定数据源";
		static final String SOURCE_FILES_NOT_EXIST = "源目录中没有Excel文件";
		static final String NO_AVAIALBE_SHEET = "数据源文件中没有可分析的工作表";

		static final String TARGET_NOT_SELECTED = "还未指定目标工作簿";
		static final String TARGET_SHEET_NOT_SELECTED = "还未指定目标工作表的名称";

		static final String PARAM_NOT_SELECTED = "还未指定分析参数";
	}

	void alertMessage(String msg) {
		JOptionPane.showMessageDialog(this, msg, "警告", JOptionPane.ERROR_MESSAGE);
	}

	public void preExecute(WorkCluster worker) {
		enabled(false);
	}

	public void onError(WorkCluster worker) {
		alertMessage("内部错误！");
		enabled(true);
		worker.setCallback(null);
	}

	public void PostExecute(WorkCluster worker) {
		JOptionPane.showMessageDialog(this, "完成", operationPanel.getOperation().getTitle(), JOptionPane.PLAIN_MESSAGE);
		enabled(true);
		worker.setCallback(null);
	}

	private void enabled(boolean enabled) {
		sourceDataPanel.setEnabled(enabled);
		targetFilePanel.setEnabled(enabled);
		operationPanel.setEnabled(enabled);
		startButton.setEnabled(enabled);
	}
}
