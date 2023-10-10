import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.util.Objects;

public class Driver {

    //Where subjects' scores are stored
    //TODO: Change location
    static File dataSheetFile = new File("/Users/coltenglover/Desktop/SubjectData/DataEntry.xlsx");

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
                        //TODO: Change back to <= subjectDataSheet.getLastRowNum()
                        for (int i = 2; i <= 81; i++) { //iterate over rows
                            int age = Integer.parseInt(subjectDataSheet.getRow(i).getCell(ColumnLabels.AGE_COLUMN).getRawValue());

                            //Time/Test (Baseline, Mid, Post)
                            String time = subjectDataSheet.getRow(i).getCell(ColumnLabels.TEST_TIME).getStringCellValue();
                            //Manipulate time to have capital letter, lowercase remainder
                            time = time.toUpperCase().charAt(0) + time.substring(1).toLowerCase();

                            int listLearning = 0;
                            int storyMemory = 0;
                            try {
                                //Get List Learning score (y)
                                listLearning = (int) subjectDataSheet.getRow(i).getCell(ColumnLabels.LIST_LEARNING).getNumericCellValue();

                                //Get Story Memory score (x)
                                storyMemory = (int) subjectDataSheet.getRow(i).getCell(ColumnLabels.STORY_MEMORY).getNumericCellValue();
                            } catch (IllegalStateException e) {
                                System.err.println(e.getMessage());
                                System.err.printf("Encountered: %s%n", subjectDataSheet.getRow(i).getCell(ColumnLabels.LIST_LEARNING).getStringCellValue());
                                continue;
                            }

                            XSSFWorkbook indexScoreTable = null;

                            if (age <= 39) {
                                indexScoreTable = new XSSFWorkbook(Objects.requireNonNull(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/" + time + "/20-39-" + time + ".xlsx")));
                            } else if (age <= 49) {
                                indexScoreTable = new XSSFWorkbook(Objects.requireNonNull(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/" + time + "/40-49-" + time + ".xlsx")));
                            } else if (age <= 59) {
                                indexScoreTable = new XSSFWorkbook(Objects.requireNonNull(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/" + time + "/50-59-" + time + ".xlsx")));
                            } else if (age <= 69) {
                                indexScoreTable = new XSSFWorkbook(Objects.requireNonNull(Driver.class.getClassLoader().getResourceAsStream("IndexScoreTables/" + time + "/60-69-" + time + ".xlsx")));
                            }

                            XSSFSheet table = indexScoreTable.getSheetAt(0);
                            int score = (int) table.getRow(listLearning + 1).getCell(storyMemory + 1).getNumericCellValue();

                            //Write score to column
                            subjectDataSheet.getRow(i).getCell(ColumnLabels.IMMEDIATE_MEMORY_COLUMN).setCellValue(score);
                            FileOutputStream fos = new FileOutputStream("/Users/coltenglover/Desktop/SubjectData/DataEntry.xlsx");
                            workbook.write(fos);
                            fos.close();
                        }
                        System.out.println("Finished writing scores!");
                        workbook.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    } catch  (Exception e) {
                        e.printStackTrace();
                    }
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
