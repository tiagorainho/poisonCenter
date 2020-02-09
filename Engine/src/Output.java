import javax.swing.JFrame;
import javax.swing.JOptionPane;

public class Output {
	private static JFrame frame = new JFrame();
	
	public static void print(String string) {
		JOptionPane.showMessageDialog(frame, string);
	}
}
