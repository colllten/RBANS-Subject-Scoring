import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;

public class Driver {

    //Where subjects' scores are stored
    static File dataSheetFile = new File("/Users/coltenglover/Desktop/SubjectData/DataEntry.xlsx");

    private enum TestStage {
        BASELINE, MID, POST
    }

    private enum Test {
        DELAYED_MEMORY, ATTENTION, LANGUAGE, VISUOSPATIAL_CONSTRUCTIONAL, IMMEDIATE_MEMORY
    }

    public static void main(String[] args) {
        //TODO: Add file selector, take out paths
        //TODO: Add logging
        try {
            EventQueue.invokeAndWait(new Runnable() {
                @Override
                public void run() {

                    //Set the dataSheetFile to the MS Excel sheet holding all subject data
                    //TODO: Uncomment this
                    //setSubjectsDataFile();

                    //Open input stream to data file
                    FileInputStream fis;
                    try {
                        fis = new FileInputStream(dataSheetFile); //Stream to the subject's data
                    } catch (FileNotFoundException e) {
                        JOptionPane.showMessageDialog(null, "Could not find file.");
                        e.printStackTrace();
                        return;
                    } catch (Exception e) {
                        JOptionPane.showMessageDialog(null, "An unexpected error ocurred.");
                        e.printStackTrace();
                        return;
                    }

                    //Use Apache POI to create workbook
                    try {
                        //Represents book containing Excel sheets
                        XSSFWorkbook workbook = new XSSFWorkbook(fis);
                        //Get Input 1 subjectDataSheet because index 0 contains notes
                        //TODO: Take out '1' and let user choose where the sheet is located, or ask Emily to remove index 0
                        XSSFSheet subjectDataSheet = workbook.getSheetAt(1);

                        //TODO: Start reading tables
                        //Subject's data starts on 3rd row (index 2)
                        for (int i = 2; i <= subjectDataSheet.getLastRowNum(); i++) { //iterate over rows
                            int age = Integer.parseInt(subjectDataSheet.getRow(i).getCell(2).getRawValue());

                            //Time/Test (Baseline, Mid, Post)
                            String time = subjectDataSheet.getRow(i).getCell(3).getStringCellValue();
                            //Manipulate time to have capital letter, lowercase remainder
                            time = time.toUpperCase().charAt(0) + time.substring(1).toLowerCase();

                            int listLearning = 0;
                            int storyMemory = 0;
                            try {
                                //Get List Learning score (y)
                                listLearning = (int) subjectDataSheet.getRow(i).getCell(4).getNumericCellValue();

                                //Get Story Memory score (x)
                                storyMemory = (int) subjectDataSheet.getRow(i).getCell(5).getNumericCellValue();
                            } catch (IllegalStateException e) {
                                System.err.println(e.getMessage());
                                System.err.printf("Encountered: %s%n", subjectDataSheet.getRow(i).getCell(4).getStringCellValue());
                                continue;
                            }

                            XSSFWorkbook indexScoreTable = null;

                            if (age <= 39) {
                                indexScoreTable = new XSSFWorkbook(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/" + time + "/20-39-" + time +".xlsx"));
                            } else if (age <= 49) {
                                indexScoreTable = new XSSFWorkbook(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/" + time + "/40-49-" + time +".xlsx"));
                            } else if (age <= 59) {
                                indexScoreTable = new XSSFWorkbook(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/" + time + "/50-59-" + time +".xlsx"));
                            } else if (age <= 69) {
                                indexScoreTable = new XSSFWorkbook(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/" + time + "/60-69-" + time +".xlsx"));
                            }

                            XSSFSheet table = indexScoreTable.getSheetAt(0);
                            int score = (int) table.getRow(listLearning + 1).getCell(storyMemory + 1).getNumericCellValue();
                        }



                        //Prompt user to pick the desired table
                        JOptionPane.showMessageDialog(null, "Choose the table for your test and age group (Baseline, " +
                                "Test 1, 20-39)");
                        //Iterate through every row
                        for (Row row : subjectDataSheet) {
                            if (row.getCell(2).getCellType() == CellType.NUMERIC && row.getCell(3).getCellType() == CellType.STRING) {
                                if (row.getCell(2).getNumericCellValue() >= 20 && row.getCell(2).getNumericCellValue() <= 39 && row.getCell(3).getStringCellValue().equals("Baseline")) {
                                    //TODO: write the intersection of both x and y
//                                    row.createCell(16).setCellValue(getImmediateMemoryScore((int) row.getCell(4)
//                                            .getNumericCellValue(), (int) row.getCell(5).getNumericCellValue(), fileChooser.getSelectedFile()));
                                }
                            }
                        }
                        //Write all changes back to the file
                        FileOutputStream fos = new FileOutputStream(dataSheetFile);
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
                private static int getImmediateMemoryScore(int listLearning, int storyMemory, File table) {
                    //TODO: find intersection of both values
                    //TODO: implement file explorer

                    try {
                        //New workbook for new table
                        XSSFWorkbook workbook = new XSSFWorkbook(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/Baseline/20-39-Baseline.xlsx"));
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
            });
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Prompts user to select the file where all subjects scores are listed for each test. The selected file must be
     * a Microsoft Excel file (xlsx) and is used throughout the remainder of the program.
     */
    static void setSubjectsDataFile() {
        do {
            //Prompt user to select the excel file where participant's scores are
            JOptionPane.showMessageDialog(null, "Choose the excel file where subject data is stored", "Welcome",
                    JOptionPane.PLAIN_MESSAGE);
            //Choose the data set
            //TODO: Change path
            JFileChooser fileChooser = new JFileChooser("/Users/coltenglover/Downloads/");
            //Filter for Excel files only
            fileChooser.setFileFilter(new FileNameExtensionFilter("XLSX", "xlsx"));
            fileChooser.setDialogTitle("Choose data Excel file");
            //Displays the GUI
            fileChooser.showOpenDialog(null);
            dataSheetFile = fileChooser.getSelectedFile();
            if (!dataSheetFile.canRead()) {
                JOptionPane.showMessageDialog(null, "Error: File is unreadable", "Error", JOptionPane.ERROR_MESSAGE);
            }
        } while (!dataSheetFile.canRead());
    }
}
