package tools.fjl.excel.ui;

import static tools.fjl.excel.ui.Constants.ITEM_GAP;
import static tools.fjl.excel.ui.Constants.ITEM_HEIGHT;

import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.robin.util.StringUtils;

import tools.fjl.excel.Logger;
import tools.fjl.excel.Utils;
import tools.fjl.excel.poi.PoiUtils;

final class TargetFilePanel extends BasePanel implements ActionListener {

	private static final long serialVersionUID = 1L;

	private JButton buttonForTarget;
	private JTextField targetFiled;
	private JComboBox<String> sheetNameBox;
	private JCheckBox targetAsSourceCheckBox;

	private JFileChooser jfc = new JFileChooser();

	private WorkFrame frame;

	public TargetFilePanel(WorkFrame frame, Rectangle r) {
		this.frame = frame;

		jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		jfc.setAcceptAllFileFilterUsed(false);
		jfc.setFileFilter(new FileFilter() {

			@Override
			public String getDescription() {
				return "Excel文件";
			}

			@Override
			public boolean accept(File f) {
				return f.isDirectory() || Utils.isExcelFile(f.getAbsolutePath());
			}
		});

		final int gapDelta = 5;

		setBorder(BorderFactory.createTitledBorder("要生成的目标文件"));
		int panelHeight = ITEM_HEIGHT * 3 + ITEM_GAP * 4 - gapDelta;
		r.setSize((int) r.getWidth(), panelHeight);
		setBounds(r);
		setLayout(null);

		int x, y, width, height;

		x = ITEM_GAP;
		y = ITEM_GAP + gapDelta;
		width = 70;
		height = ITEM_HEIGHT;
		JLabel sourceLabel = new JLabel("目标工作簿");
		sourceLabel.setBounds(x, y, width, height);
		add(sourceLabel);

		width = 40;
		x = getWidth() - ITEM_GAP - width;
		buttonForTarget = new JButton("...");
		buttonForTarget.setBounds(x, y, width, height);
		buttonForTarget.addActionListener(this);
		add(buttonForTarget);

		x = sourceLabel.getX() + sourceLabel.getWidth() + ITEM_GAP;
		width = buttonForTarget.getX() - ITEM_GAP - x;
		targetFiled = new JTextField();
		targetFiled.setEditable(false);
		targetFiled.setBounds(x, y, width, height);
		targetFiled.getDocument().addDocumentListener(new DocumentListener() {

			public void removeUpdate(DocumentEvent e) {
			}

			public void insertUpdate(DocumentEvent e) {
				String targetPath = targetFiled.getText().trim();
				Logger.log("");
				Logger.log("select target Workbook :" + targetPath);
				Workbook target = null;
				try {
					target = PoiUtils.getWorkbookByPath(targetPath);
				} catch (Exception ex) {
					String msg = String.format("目标工作簿 :%s，获取Workbook对象时发生错误;%s", targetPath, ex.getMessage());
					Logger.log(new Exception(msg));
					TargetFilePanel.this.frame.alertMessage(msg);
					return;
				}

				sheetNameBox.removeAllItems();

				final int sheetCount = target.getNumberOfSheets();
				for (int i = 0; i < sheetCount; i++) {
					Sheet sheet = null;
					try {
						sheet = target.getSheetAt(i);
						sheetNameBox.addItem(sheet.getSheetName());
					} catch (Exception ex) {
						String msg = String.format("目标工作簿 :%s，获取index为%d的Sheet的名称时发生错误;%s", targetPath, i,
								ex.getMessage());
						Logger.log(new Exception(msg));
						TargetFilePanel.this.frame.alertMessage(msg);

						sheetNameBox.removeAllItems();
						return;
					}
				}
			}

			public void changedUpdate(DocumentEvent e) {
			}
		});
		add(targetFiled);

		x = ITEM_GAP;
		y += ITEM_HEIGHT;
		y += ITEM_GAP;
		y -= gapDelta;
		width = sourceLabel.getWidth();
		JLabel sheetLabel = new JLabel("目标工作表");
		sheetLabel.setBounds(x, y, width, height);
		add(sheetLabel);

		x = targetFiled.getX();
		width = targetFiled.getWidth();
		sheetNameBox = new JComboBox<String>();
		sheetNameBox.setEditable(true);
		sheetNameBox.setBounds(x, y, width, height);
		add(sheetNameBox);

		x = ITEM_GAP;
		y += ITEM_HEIGHT;
		y += ITEM_GAP;
		y -= gapDelta;
		width = 230;
		JLabel includeTargetFileAsSourceLabel = new JLabel("是否将目标工作簿视为源数据进行分析");
		includeTargetFileAsSourceLabel.setBounds(x, y, width, height);
		add(includeTargetFileAsSourceLabel);

		x = includeTargetFileAsSourceLabel.getX() + includeTargetFileAsSourceLabel.getWidth() + ITEM_GAP;
		width = ITEM_HEIGHT;
		targetAsSourceCheckBox = new JCheckBox();
		targetAsSourceCheckBox.setBounds(x, y, width, height);
		add(targetAsSourceCheckBox);
	}

	public void actionPerformed(ActionEvent e) {
		if (e.getSource().equals(buttonForTarget)) {
			int state = jfc.showOpenDialog(null);
			if (state == 1) {
				return;
			} else {
				File f = jfc.getSelectedFile();
				targetFiled.setText(f.getAbsolutePath());
			}
		} else if (e.getSource().equals(sheetNameBox)) {
		}
	}

	boolean isTargetWorkbookSelected() {
		return !StringUtils.isEmpty(targetFiled.getText().trim());
	}

	String getTargetPath() {
		return targetFiled.getText().trim();
	}

	boolean isTargetSheetNameSelected() {
		return !StringUtils.isEmpty((String) sheetNameBox.getSelectedItem());
	}

	String getTargetSheetName() {
		return (String) sheetNameBox.getSelectedItem();
	}

	boolean includeTargetWorkbook() {
		return targetAsSourceCheckBox.isSelected();
	}
}
