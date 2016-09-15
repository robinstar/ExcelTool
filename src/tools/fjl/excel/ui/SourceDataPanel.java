package tools.fjl.excel.ui;

import static tools.fjl.excel.ui.Constants.ITEM_GAP;
import static tools.fjl.excel.ui.Constants.ITEM_HEIGHT;

import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;
import java.util.Set;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileFilter;

import com.robin.util.StringUtils;

import tools.fjl.excel.Logger;
import tools.fjl.excel.SheetIdentity;
import tools.fjl.excel.Utils;
import tools.fjl.excel.ui.SourceParser.Callback;

final class SourceDataPanel extends BasePanel implements ActionListener, Callback {

	private static final long serialVersionUID = 1L;

	private SourceParser parser = new SourceParser();
	private JButton buttonForSource;
	private JTextField sourceFolder;
	private JComboBox<String> sheetNameBox;

	private JFileChooser jfc = new JFileChooser();

	interface SourceParseStateListener {
		static final int STATE_PARSING = 0;
		static final int STATE_DONE = 1;
		static final int STATE_ERROR = 2;

		void onSourceParseStateChanged(int state, String msg);
	}

	private SourceParseStateListener stateListener;

	void setSourceParseStateListener(SourceParseStateListener listener) {
		stateListener = listener;
	}

	public SourceDataPanel(Rectangle r) {
		parser.setCallback(this);

		jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		jfc.setAcceptAllFileFilterUsed(false);
		jfc.setFileFilter(new FileFilter() {

			@Override
			public String getDescription() {
				return "目录或Excel文件";
			}

			@Override
			public boolean accept(File f) {
				return f.isDirectory() || Utils.isExcelFile(f.getAbsolutePath());
			}
		});

		final int gapDelta = 5;

		setBorder(BorderFactory.createTitledBorder("源数据"));
		int panelHeight = ITEM_HEIGHT * 2 + ITEM_GAP * 3;
		r.setSize((int) r.getWidth(), panelHeight);
		setBounds(r);
		setLayout(null);

		int x, y, width, height;

		x = ITEM_GAP;
		y = ITEM_GAP + gapDelta;
		width = 70;
		height = ITEM_HEIGHT;
		JLabel sourceLabel = new JLabel("数据文件夹");
		sourceLabel.setBounds(x, y, width, height);
		add(sourceLabel);

		width = 40;
		x = getWidth() - ITEM_GAP - width;
		buttonForSource = new JButton("...");
		buttonForSource.setBounds(x, y, width, height);
		buttonForSource.addActionListener(this);
		add(buttonForSource);

		x = sourceLabel.getX() + sourceLabel.getWidth() + ITEM_GAP;
		width = buttonForSource.getX() - ITEM_GAP - x;
		sourceFolder = new JTextField();
		sourceFolder.setEditable(false);
		sourceFolder.setBounds(x, y, width, height);
		sourceFolder.getDocument().addDocumentListener(new DocumentListener() {

			public void removeUpdate(DocumentEvent e) {
			}

			public void insertUpdate(DocumentEvent e) {
				final String path = sourceFolder.getText();
				if (path == null || path.equals("")) {
					return;
				}

				Logger.log("");
				Logger.log("SourceDataPanel parse folder :" + path);
				try {
					parser.parse(path);
				} catch (Exception ex) {
					Logger.log(ex);

					if (stateListener != null) {
						stateListener.onSourceParseStateChanged(SourceParseStateListener.STATE_ERROR, "源数据错误");
					}
				}
			}

			public void changedUpdate(DocumentEvent e) {
			}
		});
		add(sourceFolder);

		x = ITEM_GAP;
		y += ITEM_HEIGHT;
		y += ITEM_GAP;
		y -= gapDelta;
		width = sourceLabel.getWidth();
		height = ITEM_HEIGHT;
		JLabel sheetLabel = new JLabel("数据表名称");
		sheetLabel.setBounds(x, y, width, height);
		add(sheetLabel);

		x = sourceFolder.getX();
		width = sourceFolder.getWidth();
		sheetNameBox = new JComboBox<String>();
		sheetNameBox.setEnabled(false);
		sheetNameBox.setBounds(x, y, width, height);
		add(sheetNameBox);
	}

	public String getSourceSheetName() {
		return parser.getSelectedSheetName();
	}

	public Set<SheetIdentity> getSourceSheets() {
		return parser.getSelectedSheets();
	}

	boolean isSourceSelected() {
		return !StringUtils.isEmpty(sourceFolder.getText().trim());
	}

	boolean hasAvaiableSourceBooks() {
		return parser.getWorkbookPathes().size() > 0;
	}

	boolean hasAvaialbeSourceSheets() {
		return parser.getSheetNames().size() > 0;
	}

	List<String> getAllWorkbooks() {
		return parser.getWorkbookPathes();
	}

	public void actionPerformed(ActionEvent e) {
		if (e.getSource().equals(buttonForSource)) {
			int state = jfc.showOpenDialog(null);
			if (state == 1) {
				return;
			} else {
				File f = jfc.getSelectedFile();
				if (f.isFile()) {
					f = f.getParentFile();
				}
				sourceFolder.setText(f.getAbsolutePath());
			}
		} else if (e.getSource().equals(sheetNameBox)) {
			parser.setSelectedSheetName((String) sheetNameBox.getSelectedItem());
		}
	}

	public void preParse() {
		buttonForSource.setEnabled(false);
		sheetNameBox.setEnabled(false);

		if (stateListener != null) {
			stateListener.onSourceParseStateChanged(SourceParseStateListener.STATE_PARSING, null);
		}
	}

	public void onParseError(String msg) {
		buttonForSource.setEnabled(true);
		sheetNameBox.setEnabled(true);

		if (stateListener != null) {
			stateListener.onSourceParseStateChanged(SourceParseStateListener.STATE_ERROR, msg);
		}
	}

	public void onParseDone(SourceParser parser) {
		buttonForSource.setEnabled(true);
		sheetNameBox.setEnabled(true);

		sheetNameBox.removeActionListener(this);
		sheetNameBox.removeAllItems();

		if (parser.getSelectedSheetName() == null) {
			return;
		}

		List<String> sheetNames = parser.getSheetNames();
		for (String sheet : sheetNames) {
			sheetNameBox.addItem(sheet);
		}

		if (parser.getSelectedSheetName() != null) {
			sheetNameBox.setSelectedItem(parser.getSelectedSheetName());
		}

		sheetNameBox.addActionListener(this);

		if (stateListener != null) {
			stateListener.onSourceParseStateChanged(SourceParseStateListener.STATE_DONE, null);
		}
	}
}
