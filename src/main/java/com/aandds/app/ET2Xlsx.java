package com.aandds.app;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ET2Xlsx {

    private static String inputFileName = "";
    private static String outputFileName = "";

    private static void printUsageAndExit() {
        System.err.println("prog input.txt output.xlsx");
        System.exit(1);
    }

    private static Logger logger = LoggerFactory.getLogger(ET2Xlsx.class);

    public static void main(String[] args) throws IOException {

        switch (args.length) {
        case 2:
            inputFileName = args[0];
            outputFileName = args[1];
            break;
        default:
            printUsageAndExit();
        }

        BufferedReader br = new BufferedReader(new FileReader(inputFileName));
        String fileLine = null;
        boolean isTableStart = false;
        boolean isTableEnd = false;
        List<Integer> widthOfEachCol = new ArrayList<Integer>();
        List<String> tableLines = new ArrayList<String>();
        while ((fileLine = br.readLine()) != null) {
            if (fileLine.startsWith("+--")) {
                isTableStart = true;
                if (widthOfEachCol.size() == 0) {
                    // parse table header line, for example:
                    // +-----+-----+-----+-----+
                    String[] toks = fileLine.split("\\+");
                    logger.debug("header toks length: " + toks.length);
                    for (String tok : toks) {
                        if (!tok.isEmpty()) {
                            widthOfEachCol.add(tok.length());
                        }
                        logger.debug("header tok: " + tok);
                    }
                }
            } else if (fileLine.startsWith("|")) {
                // do nothing
            } else {
                if (isTableStart) {
                    isTableEnd = true;
                }
            }

            // parsing end
            if (isTableEnd) {
                break;
            }

            if (isTableStart) {
                tableLines.add(fileLine);
            }
        }
        br.close();

        for (String str : tableLines) {
            logger.debug("table line: " + str);
        }

        for (int i : widthOfEachCol) {
            logger.debug("width of each col: " + i);
        }

        int rowNum = 0;
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("output");
        Row row = sheet.createRow(0);

        for (int i = 0; i < tableLines.size(); i++) {
            String line = tableLines.get(i);

            if (Util.widthOfString(line) != widthOfHeader(widthOfEachCol)) {
                throw new RuntimeException("Bad width of table line: " + line);
            }

            float maxRowHeigthInPoints = 0;
            if (line.startsWith("+--")) {
                row = sheet.createRow(rowNum++);
                logger.debug("new xlsx table line begin.");
                maxRowHeigthInPoints = 0;
            } else if (line.startsWith("|")) {
                for (int colIndex = 0; colIndex < widthOfEachCol.size(); colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        // cell in new xlsx row
                        cell = row.createCell(colIndex);
                    }
                    String cellExsitStr = cell.toString();
                    String newCellStr;
                    if (cellExsitStr.isEmpty()) {
                        newCellStr = getColunmFromLine(line, widthOfEachCol, colIndex);
                    } else {
                        newCellStr = cellExsitStr + "\n" + getColunmFromLine(line, widthOfEachCol, colIndex);
                    }

                    /*- remove first space (if it exist in all lines in current row) for cell, for example:
                     *  +-----+-----+
                     *  | ABCD| 1234|   (emacs table)
                     *  | ABCD| 1234|
                     *  +-----+-----+
                     *  to
                     *  +----+----+
                     *  |ABCD|1234|  (xlsx table)
                     *  |ABCD|1234|
                     *  +----+----+
                     */
                    if (i + 1 < tableLines.size()) {
                        // if (i+1) is not last line in table
                        boolean isNextLineNewRow = tableLines.get(i + 1).startsWith("+--");
                        if (isNextLineNewRow) {
                            // check each line in current row
                            boolean linesStartWithSpace = true;
                            for (String li : newCellStr.split("\n")) {
                                if (!li.startsWith(" ")) {
                                    linesStartWithSpace = false;
                                    break;
                                }
                            }
                            // remove first space in each line in current row
                            if (linesStartWithSpace) {
                                newCellStr = newCellStr.replaceAll("(?m)^ ", "");
                            }
                        }
                    }

                    cell.setCellValue((String) newCellStr);

                    // See
                    // http://poi.apache.org/spreadsheet/quick-guide.html#NewLinesInCells
                    CellStyle cs = workbook.createCellStyle();
                    cs.setWrapText(true);
                    cell.setCellStyle(cs);

                    // increase row height to accommodate lines of text
                    float rowHeigthInPoints = newCellStr.split("\n").length * sheet.getDefaultRowHeightInPoints();
                    maxRowHeigthInPoints = Util.getGreater(maxRowHeigthInPoints, rowHeigthInPoints);
                    logger.debug("maxRowHeigthInPoints: " + maxRowHeigthInPoints);
                    row.setHeightInPoints(maxRowHeigthInPoints);

                    // adjust column width to fit the content
                    sheet.autoSizeColumn(colIndex);
                }
            } else {
                throw new RuntimeException("Unrecognize table line: " + line);
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(outputFileName);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static int widthOfHeader(List<Integer> widthOfEachCol) {
        return widthOfEachCol.stream().reduce(0, Integer::sum) + widthOfEachCol.size() + 1;
    }

    /*-
     *  line example: |     |A    |B    |C    |
     */
    private static String getColunmFromLine(String line, List<Integer> widthOfEachCol, int currentCol) {
        logger.debug("getColunmFromLine line: " + line + ", currentCol: " + currentCol);
        int offWidth = "+".length();
        for (int i = 0; i < widthOfEachCol.size() && i < currentCol; i++) {
            offWidth += widthOfEachCol.get(i) + "+".length();
        }
        int width = widthOfEachCol.get(currentCol);

        logger.debug("getColunmFromLine offWidth: " + offWidth);
        logger.debug("getColunmFromLine width: " + width);

        String ret = "";
        int passWidth = 0;
        for (int i = 0; i < line.length(); i++) {
            if (Util.isFullwidth(Character.codePointAt(line, i))) {
                passWidth += 2;
            } else {
                passWidth += 1;
            }
            if (passWidth > offWidth) {
                ret += Character.toString(line.charAt(i));
            }
            if (Util.widthOfString(ret) >= width) {
                break;
            }
        }
        logger.debug("getColunmFromLine ret (before trim): [" + ret + "]");
        return Util.trimRight(ret);
    }
}