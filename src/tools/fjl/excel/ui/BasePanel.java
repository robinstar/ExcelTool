package tools.fjl.excel.ui;

import java.awt.Component;

import javax.swing.JPanel;

class BasePanel extends JPanel {

	private static final long serialVersionUID = 1L;

	@Override
	public void setEnabled(boolean enabled) {
		super.setEnabled(enabled);

		Component[] children = getComponents();
		for (Component component : children) {
			component.setEnabled(enabled);
		}
	}
}
