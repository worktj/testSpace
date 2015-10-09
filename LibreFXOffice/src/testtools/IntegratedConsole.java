package testtools;


import javafx.fxml.FXML;
import javafx.scene.control.TextArea;

public class IntegratedConsole  {
	@FXML
	private TextArea area;

	public void print(String t) {
		area.appendText(t + "\n");
	}
}
