package com.aandds.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Xlsx2OT {

    private static String inputFileName = "";
    private static String outputFileName = "";
    private static StringBuilder dataSb = new StringBuilder();

    private static Logger logger = LoggerFactory.getLogger(Xlsx2OT.class.getName());

    private static void printUsageAndExit() {
        System.err.println("prog input.xlsx [output.txt]");
        System.exit(1);
    }

    /*-
     * Convert xlsx to org-mode table (text-based table) , an example of org-mode table:
     * | Name  | Phone | Age |
     * |-------+-------+-----|
     * | Peter |  1234 |  17 |
     * | Anna  |  4321 |  25 |
     */
    public static void main(String[] args) throws Exception {
        System.setProperty("java.util.logging.SimpleFormatter.format", "%4$s: %5$s%n");

        switch (args.length) {
        case 1:
            inputFileName = args[0];
            break;
        case 2:
            inputFileName = args[0];
            outputFileName = args[1];
            break;
        default:
            printUsageAndExit();
        }

        try {

            FileInputStream excelFile = new FileInputStream(new File(inputFileName));

            XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = sheet.iterator();

            int rowIdx = 0;
            int maxRowNum = 0;
            int maxColNum = 0;

            long[][] widthForEachCell = new long[ConstXlsx.MAX_ROW][ConstXlsx.MAX_COL];
            String[][] eachCell = new String[ConstXlsx.MAX_ROW][ConstXlsx.MAX_COL];

            // each xlsx line
            while (iterator.hasNext() && rowIdx < ConstXlsx.MAX_ROW) {

                Row currentRow = iterator.next();

                logger.debug("Current row: " + rowIdx);

                // Don't use currentRow.iterator(), it skip blank cell!

                int colIdx = 0;

                boolean isThisLineEmpty = true;
                // each xlsx column
                for (; (colIdx < currentRow.getLastCellNum() && colIdx < ConstXlsx.MAX_COL);) {

                    Cell currentCell = currentRow.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    CellType currentCellType = currentCell.getCellTypeEnum();

                    String currentCellContent = "";

                    if (currentCellType == CellType.STRING) {
                        currentCellContent = currentCell.getStringCellValue();
                    } else if (currentCellType == CellType.NUMERIC) {
                        currentCellContent = String.valueOf(currentCell.getNumericCellValue());
                    } else {
                        currentCellContent = currentCell.toString();
                    }

                    if (!currentCellContent.trim().isEmpty()) {
                        // If one column of this line is not empty, this line is not empty.
                        isThisLineEmpty = false;
                    }

                    logger.debug("Current column " + colIdx + ": [" + currentCellContent + "]");

                    // Escape characters
                    currentCellContent = escapeOrgmodeTableChars(currentCellContent);

                    // normalize multiple lines to one line, remove leading/trailing space(s)
                    String[] lines = currentCellContent.split("\r\n|\r|\n");
                    String newCell = Arrays.stream(lines).map(String::trim).filter(s -> s.length() > 0)
                            .collect(Collectors.joining(" "));
                    if (!newCell.equals(currentCellContent)) {
                        logger.warn("Newline(s) and leading/trailing space(s) in row {} column {} [{}] is removed",
                                rowIdx, colIdx, currentCellContent);
                        currentCellContent = newCell;
                    }

                    widthForEachCell[rowIdx][colIdx] = Util.countMaxWidthInLines(currentCellContent);
                    eachCell[rowIdx][colIdx] = currentCellContent;

                    colIdx++;
                    maxColNum = Util.getGreater(maxColNum, colIdx);
                }
                rowIdx++;

                if (!isThisLineEmpty) {
                    maxRowNum = rowIdx;
                }
            }

            logger.debug("maxColNum:" + maxColNum);
            logger.debug("maxRowNum:" + maxRowNum);

            // Compute maxWidthInEachCol
            List<Integer> maxWidthInEachCol = new ArrayList<Integer>();
            for (int col = 0; col < maxColNum; col++) {
                int maxWidthCurrentCol = 0;
                for (int row = 0; row < maxRowNum; row++) {
                    maxWidthCurrentCol = Util.getGreater((int) widthForEachCell[row][col], maxWidthCurrentCol);
                }
                maxWidthInEachCol.add(maxWidthCurrentCol);
            }

            logger.debug("maxWidthInEachCol: " + maxWidthInEachCol.toString());

            while (maxWidthInEachCol.get(maxWidthInEachCol.size() - 1) == 0) {
                maxWidthInEachCol.remove(maxWidthInEachCol.size() - 1);
                maxColNum = maxWidthInEachCol.size();
            }

            logger.debug("normalize maxColNum:" + maxColNum);
            logger.debug("normalize maxRowNum:" + maxRowNum);
            logger.debug("normalize maxWidthInEachCol: " + maxWidthInEachCol.toString());

            // Output
            printOrgModeTable(eachCell, maxRowNum, maxColNum, maxWidthInEachCol);

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        if (outputFileName.isEmpty()) {
            // print to stdout
            System.out.print(dataSb.toString());
        } else {
            // print to file
            try (PrintWriter out = new PrintWriter(outputFileName, "UTF-8")) {
                out.write(dataSb.toString());
            }
        }
    }

    private static String escapeOrgmodeTableChars(String str) {
        String bak = str;
        String newStr = str.replaceAll("\\|", " \\\\vert "); // change ‘|’ to ‘ \vert ’
        if (!bak.equals(newStr)) {
            logger.debug("[" + bak + "] is changed to [" + newStr + "]");
        }
        return newStr;
    }

    private static void printOrgModeTable(String[][] eachCell, int maxRowNum, int maxColNum,
            List<Integer> maxWidthInEachCol) {
        for (int row = 0; row < maxRowNum; row++) {
            if (row == 1) {
                Util.printOrgmodeTableHline(dataSb, maxWidthInEachCol);
            }

            for (int col = 0; col < maxColNum; col++) {
                String cell = eachCell[row][col];
                logger.trace("printOrgModeTable cell[{}][{}] is {}", row, col, cell);
                if (cell == null) {
                    eachCell[row][col] = Util.padString("", maxWidthInEachCol.get(col));
                } else {
                    eachCell[row][col] = Util.padString(cell, maxWidthInEachCol.get(col));
                }
            }

            for (int col = 0; col < maxColNum; col++) {
                dataSb.append("| ");
                String[] lines = eachCell[row][col].split("\r\n|\r|\n");
                assert lines.length <= 1;
                dataSb.append(lines[0]);
                dataSb.append(" ");
            }
            dataSb.append("|\n");
        }
    }
}
