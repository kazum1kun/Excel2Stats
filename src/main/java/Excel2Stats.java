import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;


/**
 * Excel2Stats processes an external MS Excel (.xls) format and calculates the average, variance, and
 * standard deviation from the given data
 */
public class Excel2Stats extends Application {
    // User selected Excel file
    private File userFile;

    public static void main(String[] args) {
        launch(args);
    }

    // Do the calculation
    private static void doCalc(File file, String colRange, String rowRange) {
        int[] colNum = new int[0];  // Do this just to bypass the Nullity check
        int rowStart = 0, rowEnd = 0;

        // First split the column into individual numbers
        // Assign 2,3,4 to the colNum if colRange is null
        if (colRange.isEmpty()) {
            colNum = new int[3];
            colNum[0] = 2;
            colNum[1] = 3;
            colNum[2] = 4;
        } else {
            try {
                // First obtain column numbers as Strings
                String[] colNumsStr = colRange.trim().split(",");
                // Then convert them to int's
                colNum = Arrays.stream(colNumsStr).mapToInt(Integer::parseInt).toArray();
            } catch (Exception ex) {
                System.out.println("Err");
            }
        }

        // Indicator to show if there is END as input
        boolean endExists = false;

        // Split the row into two numbers
        // Assign 1 as start and END as end, if rowRange is empty
        if (rowRange.isEmpty()) {
            rowStart = 1;
            endExists = true;
        } else {
            try {
                // Obtain the range as String
                String[] rowNumStr = rowRange.trim().split("-");
                // Convert to two numbers
                rowStart = Integer.parseInt(rowNumStr[0]);
                // If latter is END, do not process it yet
                if (rowNumStr[1].equals("END"))
                    endExists = true;
                else
                    rowEnd = Integer.parseInt(rowNumStr[1]);
            } catch (Exception ex) {
                System.out.println("ErrParseInt");
            }
        }


        // The file that data being saved to
        File saveToFile = new File(file.getParent() + "/" + file.getName() + "计算结果.txt");
        // This contains the results in String
        String resultStr = "";

        try {
            // Load the Excel file and default sheet (0)
            FileInputStream fis = new FileInputStream(file);
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            HSSFSheet sheet = workbook.getSheetAt(0);

            // Get the row # when END is present
            if (endExists)
                rowEnd = sheet.getPhysicalNumberOfRows();

            // Results are recorded here, with the following format
            /*  ---------
             * |  Col#  |   0
             * |  Avg.  |   1
             * |  Var.  |   2
             * |Std.Dev.|   3
             *  ---------
             */

            double results[][] = new double[colNum.length][4];

            // Iterate through the Excel and store information in an list
            int count = 0;
            for (Integer i : colNum) {
                DescriptiveStatistics stats = new DescriptiveStatistics();
                // Detect the format of the cell and convert them when necessary
                // Check the first cell
                Cell indicatorCell = sheet.getRow(rowStart).getCell(i);

                // Numeric cells
                if (indicatorCell.getCellTypeEnum() == CellType.NUMERIC) {
                    for (int j = rowStart; j < rowEnd; j++) {
                        Double currentCellVal = sheet.getRow(j).getCell(i).getNumericCellValue();
                        stats.addValue(currentCellVal);
                    }
                }
                // String cells (percentage stored as cell)
                else if (indicatorCell.getCellTypeEnum() == CellType.STRING) {
                    for (int j = rowStart; j < rowEnd; j++) {
                        String currCellContent = sheet.getRow(j).getCell(i).getStringCellValue();
                        Double numPercentage = new Double(currCellContent.trim().replace("%", ""))
                                / 100.0;
                        stats.addValue(numPercentage);
                    }
                }

                // Calc the average
                double average = stats.getMean();
                // Calc the variance
                double variance = stats.getVariance();
                // Calc the std. deviation
                double stdDev = stats.getStandardDeviation();
                // Store them into the results array
                results[count][0] = i;
                results[count][1] = average;
                results[count][2] = variance;
                results[count][3] = stdDev;

                count++;
            }

            // Format the result
            for (int i = 0; i < count; i++) {
                int currColNum = ((int) results[i][0]);
                resultStr = resultStr.concat("列[" + currColNum + "]，" + sheet.getRow(0)
                        .getCell(currColNum).getStringCellValue() + "\n");
                resultStr = resultStr.concat("平均值：\t" + results[i][1] + "\n");
                resultStr = resultStr.concat("方差：\t" + results[i][2] + "\n");
                resultStr = resultStr.concat("标准差：\t" + results[i][3] + "\n\n");

            }
        } catch (IOException ex) {
            System.out.println("ErrReadExcel");
        }

        // Save the results
        try {
            Files.write(Paths.get(saveToFile.getPath()), resultStr.getBytes());
        } catch (IOException ex) {
            System.out.println("ErrWriteResult");
        }

        // Show a prompt with analysis result
        Alert showResult = new Alert(Alert.AlertType.INFORMATION);
        showResult.setTitle("计算结果");
        showResult.setContentText(resultStr + "结果已保存到 " + saveToFile.getPath());
        showResult.setHeaderText(null);
        showResult.showAndWait();
    }

    // FileChooser to pick an Excel Worksheet
    private static File chooseFile(Stage stage) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("请选择要计算的 Excel 工作簿");
        // Initial dir set to user's desktop
        fileChooser.setInitialDirectory(new File(
                System.getProperty("user.home") + "/Desktop"));
        // Allow only xls file here
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel 工作簿", "*.xls")
        );

        return fileChooser.showOpenDialog(stage);
    }

    // Main interface
    @Override
    public void start(Stage primaryStage) throws Exception {
        // Set the app name and default dimension
        primaryStage.setTitle("Excel 数据计算器");
        primaryStage.setWidth(450);
        primaryStage.setHeight(300);

        // Default scene
        Scene scene = new Scene(new VBox());

        // Top part: a file picker which consists of a unmodifiable TextField and a button
        HBox filePickerBox = new HBox();

        // TextField
        TextField selectedFileField = new TextField();
        // Make it unmodifiable and suggest user to pick a file
        selectedFileField.setEditable(false);
        selectedFileField.setPromptText("请选择 Excel 工作簿");
        // Make selectedFileField longer
        selectedFileField.setPrefWidth(500);

        // Button
        Button selectFileButton = new Button("...");
        // Invoke a FileChooser when pressed
        selectFileButton.setOnAction((ActionEvent event) -> {
            userFile = chooseFile(primaryStage);
            // Check for the nullity of the file
            if (userFile != null)
                // Update the TextField and process the Excel file
                selectedFileField.setText(userFile.getPath());
        });

        // Assemble the Top part and set its spacing
        filePickerBox.getChildren().addAll(selectedFileField, selectFileButton);
        filePickerBox.setSpacing(3);


        // Middle part: user may enter the range of the data to get data from
        // Includes a 3 Labels and 2 TextFields
        VBox rangeSelectBox = new VBox();

        // Description of this section
        Label rangeSelectDesc = new Label("输入采样范围");

        // Desc of the column input field
        Label colRangeSelectDesc = new Label("请输入需要采样的列，用英文逗号分隔\n" +
                "列A = 0，列B = 1，依此类推。如采样列A、B、D，则输入 0,1,3\n" +
                "留空则默认采样列C、D、E（2,3,4）。");
        // Column input field
        TextField colRangeSelectField = new TextField();
        colRangeSelectField.setPromptText("2,3,4");

        // Desc of the row input field
        Label rowRangeSelectDesc = new Label("请输入需要采样的行区间，用英文连字符分隔。最后一行可用 END 表示\n" +
                "注意！行1 = 0，行2 = 1，以此类推。如需采样行2-100，请输入 1-99\n" +
                "留空则默认采样行2到最后一行（1-END）。暂不支持多个区间。");
        // Row input field
        TextField rowRangeSelectField = new TextField();
        rowRangeSelectField.setPromptText("1-END");

        // Assemble the middle part and set is spacing
        rangeSelectBox.getChildren().addAll(rangeSelectDesc, colRangeSelectDesc, colRangeSelectField,
                rowRangeSelectDesc, rowRangeSelectField);
        rangeSelectBox.setSpacing(5);


        // Final piece: a tempting "Calculate" button!
        Button calcButton = new Button("计算");
        calcButton.setOnAction((ActionEvent event) -> doCalc(userFile,
                colRangeSelectField.getText(), rowRangeSelectField.getText()));

        // Assemble the app into a different VBox for better manageability
        VBox appBox = new VBox(filePickerBox, rangeSelectBox, calcButton);
        ((VBox) scene.getRoot()).getChildren().addAll(appBox);

        // Set spacing and padding of the app
        appBox.setSpacing(10);
        appBox.setPadding(new Insets(10, 10, 10, 10));

        // Set focus on the Calculate button
        calcButton.requestFocus();

        // Show the app
        primaryStage.setScene(scene);
        primaryStage.show();
    }
}

