import javax.swing.JFrame;
import javax.swing.JOptionPane;

public class Output {
	private static JFrame frame = new JFrame();
	
	public static void print(String string) {
		if(string == null || string.length() == 0) {
			string = "Erro!";
		}
		JOptionPane.showMessageDialog(frame, string);
		System.exit(0);
	}
	
	public static void print(String string, boolean continueProgram) {
		if(string == null || string.length() == 0) {
			string = "Erro!";
		}
		JOptionPane.showMessageDialog(frame, string);
		if(!continueProgram) {
			System.exit(0);
		}
	}
}
