import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Driver {
    static File dataSheetFile;
    public static void main(String[] args) {
        //TODO: uncomment below
//        do {
//            String filePath = JOptionPane.showInputDialog("Please enter the path to the data sheet\n" +
//                    "/Desktop/Folder1/Folder2/DataSheet.xlsx");
//            if (filePath.isEmpty()) {
//                JOptionPane.showMessageDialog(null, "Incorrect file path. Please try again.");
//            } else {
//                dataSheetFile = new File(filePath);
//            }
//        } while (!dataSheetFile.canRead());

        //Create stream to file
        FileInputStream fis;
        try {
            //TODO: Change to datasheetfile
            fis = new FileInputStream("/Users/coltenglover/Downloads/DataEntry.xlsx");
        } catch (FileNotFoundException e) {
            JOptionPane.showMessageDialog(null, "Could not find file.");
            e.printStackTrace();
            return;
        }

        //Use Apache POI to create workbook
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(1);
            XSSFCell c=  sheet.getRow(2).createCell(19);
            c.setCellValue("fjkasdfl");

        } catch (IOException e) {

        }


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
