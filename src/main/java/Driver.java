import javax.swing.*;
import java.io.File;

public class Driver {
    static File dataSheet;
    public static void main(String[] args) {
        do {
            String filePath = JOptionPane.showInputDialog("Please enter the path to the data sheet\n" +
                    "/Desktop/Folder1/Folder2/DataSheet.xlsx");
            if (filePath.isEmpty()) {
                JOptionPane.showMessageDialog(null, "Incorrect file path. Please try again.");
            } else {
                dataSheet = new File(filePath);
            }
        } while (!dataSheet.canRead());

        String[] options = {"60 - 69", "50 - 59", "40 - 49", "20 - 39"};

        //Ask for age range
        int choice = JOptionPane.showOptionDialog(null, "What is your age range?", "Age Range",
                JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, 0);
        System.out.println(options[choice]);

        //Ask for time
        options = new String[]{"Post", "Mid", "Baseline"};
        choice = JOptionPane.showOptionDialog(null, "What test stage?", "Test Stage",
                JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, 0);
        System.out.println(options[choice]);

        //Ask for kind of test
        options = new String[]{"Delayed Memory", "Attention", "Language", "Visuospatial/Constructional", "Immediate " +
                "Memory"};
        choice = JOptionPane.showOptionDialog(null, "What is the test type?", "Test Type",
                JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, 0);
        System.out.println(options[choice]);
    }
}
