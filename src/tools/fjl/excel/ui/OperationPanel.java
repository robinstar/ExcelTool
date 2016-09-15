package tools.fjl.excel.ui;

import static tools.fjl.excel.ui.Constants.ITEM_GAP;
import static tools.fjl.excel.ui.Constants.ITEM_HEIGHT;

import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.io.File;
import java.io.IOException;
import java.util.Properties;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;

import com.robin.util.PropUtils;
import com.robin.util.StringUtils;

import tools.fjl.excel.Logger;
import tools.fjl.excel.poi.Operation;

final class OperationPanel extends BasePanel implements ActionListener {

	private static final long serialVersionUID = 1L;

	private JComboBox<String> operationBox;
	private JButton paramButton;
	private JTextField paramField;

	private String profilePath;

	private JFileChooser jfc = new JFileChooser();

	public OperationPanel(Rectangle r) {
		profilePath = getClass().getResource("profile.properties").getPath();

		jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		jfc.setAcceptAllFileFilterUsed(false);
		jfc.setFileFilter(new FileFilter() {

			@Override
			public String getDescription() {
				return "txt文件";
			}

			@Override
			public boolean accept(File f) {
				return f.isDirectory() || f.getPath().endsWith(".txt");
			}
		});

		final int gapDelta = 5;

		setBorder(BorderFactory.createTitledBorder("选择分析方式"));
		int panelHeight = ITEM_HEIGHT * 2 + ITEM_GAP * 3;
		r.setSize((int) r.getWidth(), panelHeight);
		setBounds(r);
		setLayout(null);

		int x, y, width, height;

		x = ITEM_GAP;
		y = ITEM_GAP + gapDelta;
		width = 70;
		height = ITEM_HEIGHT;
		JLabel operationLabel = new JLabel("分析方式");
		operationLabel.setBounds(x, y, width, height);
		add(operationLabel);

		x = operationLabel.getX() + operationLabel.getWidth() + ITEM_GAP;
		width = getWidth() - ITEM_GAP - 40 - ITEM_GAP - x;
		operationBox = new JComboBox<String>();
		operationBox.setBounds(x, y, width, height);
		add(operationBox);

		for (Operation op : Operation.values()) {
			operationBox.addItem(op.getTitle());
		}

		try {
			Properties prop = PropUtils.read(profilePath);
			String preferOp = (String) prop.get("preferOpertion");
			int preferId = StringUtils.isEmpty(preferOp) ? 0 : Integer.valueOf(preferOp);
			Operation op = Operation.getOpertionById(preferId);
			operationBox.setSelectedItem(op.getTitle());
		} catch (IOException e) {
			Logger.log(e);
		}

		operationBox.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent event) {
				if (event.getStateChange() != ItemEvent.SELECTED) {
					return;
				}

				File file = new File("profile.properties");
				Logger.log(file.exists() + "");
				Logger.log(file.exists() + "");

				String preferOp = (String) event.getItem();
				Logger.log("");
				Logger.log("select operation :" + preferOp);

				int preferId = Operation.getIdByTitle(preferOp);
				try {
					PropUtils.write(profilePath, "preferOpertion", preferId + "");
				} catch (IOException e) {
					Logger.log(profilePath);
					Logger.log(e);
				}
			}
		});

		x = ITEM_GAP;
		y += ITEM_HEIGHT;
		y += ITEM_GAP;
		y -= gapDelta;
		width = 70;
		height = ITEM_HEIGHT;
		JLabel sourceLabel = new JLabel("分析参数");
		sourceLabel.setBounds(x, y, width, height);
		add(sourceLabel);

		width = 40;
		x = getWidth() - ITEM_GAP - width;
		paramButton = new JButton("...");
		paramButton.setBounds(x, y, width, height);
		paramButton.addActionListener(this);
		add(paramButton);

		x = sourceLabel.getX() + sourceLabel.getWidth() + ITEM_GAP;
		width = paramButton.getX() - ITEM_GAP - x;
		paramField = new JTextField();
		paramField.setEditable(false);
		paramField.setBounds(x, y, width, height);
		add(paramField);
	}

	public void actionPerformed(ActionEvent e) {
		if (e.getSource().equals(paramButton)) {
			int state = jfc.showOpenDialog(null);
			if (state == 1) {
				return;
			} else {
				File f = jfc.getSelectedFile();
				paramField.setText(f.getAbsolutePath());
			}
		}
	}

	boolean isParamPathSelected() {
		return !StringUtils.isEmpty((String) paramField.getText().trim());
	}

	String getParamPath() {
		return paramField.getText().trim();
	}

	Operation getOperation() {
		return Operation.getOpertionByTitle((String) operationBox.getSelectedItem());
	}
}
