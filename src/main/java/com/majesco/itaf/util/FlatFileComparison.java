package com.majesco.itaf.util;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.WebHelper;

public class FlatFileComparison {
static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
private final static Logger log = LogManager.getLogger(WebHelper.class.getName());

    public static void FlatFileCompare(String destPath, String expectedFlatFile, String localFlatFileLocation, String testCaseID) throws Exception {
        try (BufferedReader file1Reader = new BufferedReader(new FileReader(destPath));
             BufferedReader file2Reader = new BufferedReader(new FileReader(expectedFlatFile))) {

            List<String> file1Lines = new ArrayList<>();
            List<String> file2Lines = new ArrayList<>();
            String line;

            // Read lines from File 1
            while ((line = file1Reader.readLine()) != null) {
                file1Lines.add(line.trim());
            }

            // Read lines from File 2 and group them into data sets
            List<List<String>> file2DataSets = new ArrayList<>();
            List<String> currentDataSet = new ArrayList<>();
            while ((line = file2Reader.readLine()) != null) {
                line = line.trim();
                if (line.startsWith("FH")) {
                    // Skip lines starting with "FH"
                    continue;
                }
                if (line.startsWith("TD") || line.startsWith("PY")) {
                    currentDataSet.add(line);
                } else {
                    // If the line doesn't start with "TD" or "PY", add it to the current data set
                    currentDataSet.add(line);
                }
                if (line.startsWith("PY")) {
                    file2DataSets.add(new ArrayList<>(currentDataSet));
                    currentDataSet.clear();
                }
            }

            if (!currentDataSet.isEmpty()) {
                file2DataSets.add(currentDataSet);
            }

            // Read lines from File 2 and store them in file2Lines list
            file2Lines.clear();
            try (BufferedReader file2ReaderCopy = new BufferedReader(new FileReader(expectedFlatFile))) {
                while ((line = file2ReaderCopy.readLine()) != null) {
                    file2Lines.add(line.trim());
                }
            }

            // Compare data sets between File 1 and File 2
            boolean allDataExists = true;
            List<List<String>> missingDataSets = new ArrayList<>();

            for (List<String> dataSet : file2DataSets) {
                if (!file1Lines.containsAll(dataSet)) {
                    allDataExists = false;
                    missingDataSets.add(dataSet);
                }
            }

            if (allDataExists) {
                log.info("All data sets from File 2 exist in File 1.");
                return;
            }

            StringBuilder mismatchBuilder = new StringBuilder();
            mismatchBuilder.append("The below data provided in the expected flat file is either missing or doesn't match the data in the response flat file: ").append("\n");

            for (String file2Line : file2Lines) {
                // Read each line from File 2
                String dataToFind = file2Line.trim();
                boolean found = false;

                for (String file1Line : file1Lines) {
                    if (file1Line.equals(dataToFind)) {
                        // Data from File 2 exists in File 1
                        found = true;
                        break;
                    }
                }

                if (!found) {
                    // Data from File 2 does not exist in File 1
                    mismatchBuilder.append(file2Line).append("\n");
                }
            }

            if (!mismatchBuilder.toString().equals(mismatchBuilder.append("\n").toString())) {
                Workbook workbook = new XSSFWorkbook();

                // Create "Mismatches" sheet
                Sheet mismatchesSheet = workbook.createSheet("Differences");
                int mismatchesRowNum = 0;

                String[] mismatchLines = mismatchBuilder.toString().split("\n");

                // Write header to "Mismatches" sheet
                Row headerRow = mismatchesSheet.createRow(mismatchesRowNum++);
                Cell headerCell = headerRow.createCell(0);
                headerCell.setCellValue(mismatchLines[0]);

                // Apply red font color to header
                XSSFFont headerFont = ((XSSFWorkbook) workbook).createFont();
                headerFont.setColor(IndexedColors.RED.getIndex());
                CellStyle headerStyle = workbook.createCellStyle();
                headerStyle.setFont(headerFont);
                headerCell.setCellStyle(headerStyle);

                // Write mismatches to "Mismatches" sheet
                for (List<String> dataSet : missingDataSets) {
                    for (String dataSetLine  : dataSet) {
                        Row row = mismatchesSheet.createRow(mismatchesRowNum++);
                        Cell cell = row.createCell(0);
                        cell.setCellValue(dataSetLine );
                    }
                }

                // Create "ResponseFlatFile" sheet
                Sheet responseFlatFileSheet = workbook.createSheet("ResponseFlatFile");
                int responseFlatFileRowNum = 0;

                // Write file1Lines to "ResponseFlatFile" sheet
                for (String file1Line : file1Lines) {
                    Row row = responseFlatFileSheet.createRow(responseFlatFileRowNum++);
                    row.createCell(0).setCellValue(file1Line);
                }

                // Create "ExpectedFlatFile" sheet
                Sheet expectedFlatFileSheet = workbook.createSheet("ExpectedFlatFile");
                int expectedFlatFileRowNum = 0;

                // Write file2Lines to "ExpectedFlatFile" sheet
                for (String file2Line : file2Lines) {
                    Row row = expectedFlatFileSheet.createRow(expectedFlatFileRowNum++);
                    row.createCell(0).setCellValue(file2Line);
                }

                // Autosize columns for better visibility in all sheets
                for (int i = 0; i < 3; i++) {
                    workbook.getSheetAt(i).autoSizeColumn(0);
                }

                // Save the Excel file
                String excelFilePath = (localFlatFileLocation + "/" + testCaseID + "_"+ "Differences.xlsx");
                try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                    workbook.write(fos);
                    log.info("Differences.xlsx created: " + excelFilePath);
                	webDriver.getReport().setMessage("Expected file doesn't match Response file, refer the Differences.xlsx for details");
                	webDriver.getReport().setStatus("FAIL");
                } catch (IOException e) {
                    e.printStackTrace();
                } finally {
                    workbook.close();
                }
            } else {
                log.info("All data from File 2 exists in File 1.");
				webDriver.getReport().setMessage("All the data from Expected file exists in Actual file");
				webDriver.getReport().setStatus("PASS");
            }

        } catch (IOException e) {
            // e.printStackTrace();
            throw new Exception("Error while reading the flat files: " + e.getMessage());
        } catch (Exception e) {
            // Handle any other unexpected errors
            throw new Exception("An unexpected error occurred: " + e.getMessage());
        }
    }
}
