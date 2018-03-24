package com.aandds.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Xlsx2ET {

    private static String inputFileName = "";
    private static String outputFileName = "";
    private static StringBuilder dataSb = new StringBuilder();

    private static Logger logger = LoggerFactory.getLogger(Xlsx2ET.class.getName());

    private static void printUsageAndExit() {
        System.err.println("prog input.xlsx [output.txt]");
        System.exit(1);
    }

    /*-
     * Convert xlsx to emacs table (text-based table) , an example of emacs table:
     * +-----+-----+-----+-----+
     * |     |A    |B    |C    |
     * +-----+-----+-----+-----+
     * |1    |1a   |1b   |1c   |
     * +-----+-----+-----+-----+
     * |2    |2a   |2b   |2c   |
     * +-----+-----+-----+-----+
     * |3    |3a   |3b   |3c   |
     * +-----+-----+-----+-----+
     */
    public static void main(String[] args) throws Exception {
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

            List<Integer> maxHeightInEachRow = new ArrayList<Integer>();

            long[][] widthForEachCell = new long[ConstXlsx.MAX_ROW][ConstXlsx.MAX_COL];
            String[][] eachCell = new String[ConstXlsx.MAX_ROW][ConstXlsx.MAX_COL];

            // each xlsx line
            while (iterator.hasNext() && rowIdx < ConstXlsx.MAX_ROW) {

                Row currentRow = iterator.next();

                logger.debug("Current row: " + rowIdx);

                // Don't use currentRow.iterator(), it skip blank cell!

                int colIdx = 0;

                int maxLinesInCurrentRow = 0;

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

                    // Convert leading tabs to spaces, remove tailing tabs.
                    currentCellContent = Util.untablify(currentCellContent);

                    logger.debug("Current column " + colIdx + ": [" + currentCellContent + "]");

                    maxLinesInCurrentRow = Util.getGreater(Util.countLines(currentCellContent), maxLinesInCurrentRow);

                    widthForEachCell[rowIdx][colIdx] = Util.countMaxWidthInLines(currentCellContent);
                    eachCell[rowIdx][colIdx] = currentCellContent;

                    colIdx++;
                    maxColNum = Util.getGreater(maxColNum, colIdx);
                }

                maxHeightInEachRow.add(maxLinesInCurrentRow);

                rowIdx++;
            }
            maxRowNum = rowIdx;

            // Util.print2dArray(widthForEachCell);

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
            logger.debug("maxHeightInEachRow: " + maxHeightInEachRow.toString());

            while (maxWidthInEachCol.get(maxWidthInEachCol.size() - 1) == 0) {
                maxWidthInEachCol.remove(maxWidthInEachCol.size() - 1);
                maxColNum = maxWidthInEachCol.size();
            }
            while (maxHeightInEachRow.get(maxHeightInEachRow.size() - 1) == 0) {
                maxHeightInEachRow.remove(maxHeightInEachRow.size() - 1);
                maxRowNum = maxHeightInEachRow.size();
            }

            logger.debug("normalize maxColNum:" + maxColNum);
            logger.debug("normalize maxRowNum:" + maxRowNum);
            logger.debug("normalize maxWidthInEachCol: " + maxWidthInEachCol.toString());
            logger.debug("normalize maxHeightInEachRow: " + maxHeightInEachRow.toString());

            // Output
            printEmacsTable(eachCell, maxRowNum, maxColNum, maxHeightInEachRow, maxWidthInEachCol);

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

    private static void printEmacsTable(String[][] eachCell, int maxRowNum, int maxColNum,
            List<Integer> maxHeightInEachRow, List<Integer> maxWidthInEachCol) {
        for (int row = 0; row < maxRowNum; row++) {
            Util.printEmacsTableHline(dataSb, maxWidthInEachCol);

            for (int col = 0; col < maxColNum; col++) {
                String cell = eachCell[row][col];
                eachCell[row][col] = expandString(cell, maxWidthInEachCol.get(col), maxHeightInEachRow.get(row));
            }

            for (int i = 0; i < maxHeightInEachRow.get(row); i++) {
                for (int col = 0; col < maxColNum; col++) {
                    dataSb.append("| ");
                    String[] lines = eachCell[row][col].split("\r\n|\r|\n");
                    assert lines.length == maxHeightInEachRow.get(row);
                    dataSb.append(lines[i]);
                    dataSb.append(" ");
                }
                dataSb.append("|\n");
            }
        }
        Util.printEmacsTableHline(dataSb, maxWidthInEachCol);
    }

    /*-
     * Expand string to specified width and height, for example:
     * Input: str = "abc", width = 4, height = 3
     * Output: "abc \n    \n    "
     */
    private static String expandString(String str, int width, int height) {
        logger.debug("expandIt width: " + width + ", height: " + height + ", input: [" + str + "]");
        String ret = "";

        if (str == null) {
            for (int i = 0; i < height; i++) {
                ret += Util.padString("", width);
                ret += "\n";
            }
        } else {
            String[] lines = str.split("\r\n|\r|\n");
            for (String line : lines) {
                ret += Util.padString(line, width);
                ret += "\n";
            }
            for (int i = 0; i < height - lines.length; i++) {
                ret += Util.padString("", width);
                ret += "\n";
            }
        }

        if (ret.endsWith("\n")) {
            ret = ret.substring(0, ret.length() - "\n".length());
        }

        logger.debug("expandIt output: [" + ret.replace(' ', '_') + "]");
        return ret;
    }
}
