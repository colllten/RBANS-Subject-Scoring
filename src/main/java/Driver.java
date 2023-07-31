import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.*;

public class Driver {
    static File dataSheetFile;

    private enum TestStage {
        BASELINE, MID, POST
    }

    private enum Test {
        DELAYED_MEMORY, ATTENTION, LANGUAGE, VISUOSPATIAL_CONSTRUCTIONAL, IMMEDIATE_MEMORY
    }

    public static void main(String[] args) {
        //File f = fileChooser.getSelectedFile();
        //System.out.println(f.getName());
        //TODO: uncomment below
        //TODO: Add file selector, take out paths
        do {
            String filePath = JOptionPane.showInputDialog("Please enter the path to the data sheet\n" +
                    "/Desktop/Folder1/Folder2/DataSheet.xlsx");
            if (filePath.isEmpty()) {
                JOptionPane.showMessageDialog(null, "Incorrect file path. Please try again.");
            } else {
                dataSheetFile = new File(filePath);
            }
        } while (!dataSheetFile.canRead());

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
            //Represents book containing Excel sheets
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            //Get Input 1 sheet
            XSSFSheet sheet = workbook.getSheetAt(1);

            //TODO: Find score for each person in BASELINE, 20-39, 1st Test
            for (Row row : sheet) {
                if (row.getCell(2).getCellType() == CellType.NUMERIC && row.getCell(3).getCellType() == CellType.STRING) {
                    if (row.getCell(2).getNumericCellValue() >= 20 && row.getCell(2).getNumericCellValue() <= 39 && row.getCell(3).getStringCellValue().equals("Baseline")) {
                        //TODO: write the intersection of both x and y
                        row.createCell(16).setCellValue(getImmediateMemoryScore((int) row.getCell(4).getNumericCellValue(), (int) row.getCell(5).getNumericCellValue()));
                    }
                }
            }
            FileOutputStream fos = new FileOutputStream("/Users/coltenglover/Downloads/DataEntry.xlsx");
            workbook.write(fos);
            fos.close();
            System.out.println("Wrote the values");
        } catch (IOException e) {
            e.printStackTrace();
        } catch  (Exception e) {
            e.printStackTrace();
        }


        String[] options = {"60 - 69", "50 - 59", "40 - 49", "20 - 39"};

        //Ask for age range
        int choice = JOptionPane.showOptionDialog(null, "What is your age range?", "Age Range",
                JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, 0);
        int ageLowerBound;
        int ageUpperBound;

        switch (choice) {
            case 3 -> {
                ageLowerBound = 20;
                ageUpperBound = 39;
            }
            case 2 -> {
                ageLowerBound = 40;
                ageUpperBound = 49;
            }
            case 1 -> {
                ageLowerBound = 50;
                ageUpperBound = 59;
            }
            case 0 -> {
                ageLowerBound = 60;
                ageUpperBound = 69;
            }
            default -> {
                //TODO: Add error handling
                ageLowerBound = 0;
                ageUpperBound = 1;
            }
        }

        //Ask for time
        options = new String[]{"Post", "Mid", "Baseline"};
        choice = JOptionPane.showOptionDialog(null, "What test stage?", "Test Stage",
                JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, 0);


        TestStage testStage;
        switch (choice) {
            case 2 -> {
                testStage = TestStage.BASELINE;
            }
            case 1 -> {
                testStage = TestStage.MID;
            }
            case 0 -> {
                testStage = TestStage.POST;
            }
            default -> {
                //TODO: Add error handling
                testStage = TestStage.POST;
            }
        }

        //Ask for kind of test
        options = new String[]{"Delayed Memory", "Attention", "Language", "Visuospatial/Constructional", "Immediate " +
                "Memory"};
        choice = JOptionPane.showOptionDialog(null, "What is the test type?", "Test Type",
                JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, 0);
        Test test;
        switch (choice) {
            case 4 -> {
                test = Test.IMMEDIATE_MEMORY;
            }
            case 3 -> {
                test = Test.VISUOSPATIAL_CONSTRUCTIONAL;
            }
            case 2 -> {
                test = Test.LANGUAGE;
            }
            case 1 -> {
                test = Test.ATTENTION;
            }
            case 0 -> {
                test = Test.DELAYED_MEMORY;
            }
            default -> {
                //TODO: Add error handling
                test = Test.IMMEDIATE_MEMORY;
            }
        }
        System.out.printf("Lower: %d\nUpper: %d\nTestType: %s\nTest: %s\n", ageLowerBound, ageUpperBound, testStage,
                test);
    }

    /**
     * Finds intersection of listLearning (y) and storyMemory (x)
     * @param listLearning y-value
     * @param storyMemory x-value
     * @return Intersection of x and y
     */
    private static int getImmediateMemoryScore(int listLearning, int storyMemory) {
        //TODO: find intersection of both values
        //TODO: implement file explorer
        JFileChooser fileChooser = new JFileChooser();
        //TODO: Change to user settings
        fileChooser.setCurrentDirectory(new File("/Users/coltenglover/Desktop"));
        //Present the window
        fileChooser.showOpenDialog(null);

        try {
            FileInputStream fis = new FileInputStream(fileChooser.getSelectedFile());
            //New workbook for new table
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(listLearning + 1);
            return (int) row.getCell(storyMemory + 1).getNumericCellValue();
        } catch (FileNotFoundException e) {
            JOptionPane.showMessageDialog(null, "File could not be found.", "Error", JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
        } catch (IOException e) {
            //TODO: Do better error handling
            throw new RuntimeException(e);
        }
        return -1;
    }
}
