package com.aandds.app;

import java.util.Arrays;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.ibm.icu.lang.UCharacter;
import com.ibm.icu.lang.UProperty;

public class Util {
    private static Logger logger = LoggerFactory.getLogger(Util.class);

    public static void print2dArray(Object[] a) {
        System.err.println(Arrays.deepToString(a));
    }

    public static <T extends Comparable<T>> T getGreater(T a, T b) {
        return a.compareTo(b) > 0 ? a : b;
    }

    /**
     * Pad string to specified width. It also support Chinese.
     * 
     * @param line
     * @param width
     * @return
     */
    public static String padString(String line, int width) {
        logger.trace("padString, width: " + width + ", line: [" + line + "]");
        long w = widthOfString(line);
        if (w > width) {
            logger.warn("width of {} is {}, no need pad to {}, do nothing.", line, w, width);
            return line;
        }
        for (long i = w; i < width; i++) {
            line += " ";
        }
        logger.trace("padString, output: [" + line + "]");
        return line;
    }

    public static long widthOfString(String str) {
        long numOfFullwidth = str.codePoints().filter(codePoint -> isFullwidth(codePoint)).count();
        return str.length() + numOfFullwidth;
    }

    /*
     * Test codePoint is fullwidth or not.
     * 
     * See http://www.unicode.org/reports/tr11/
     */
    public static boolean isFullwidth(int codePoint) {
        int value = UCharacter.getIntPropertyValue(codePoint, UProperty.EAST_ASIAN_WIDTH);

        String str = String.valueOf(Character.toChars(codePoint));
        switch (value) {
        case UCharacter.EastAsianWidth.NARROW:
            logger.trace("character [{}] width is NARROW", str);
            return false;
        case UCharacter.EastAsianWidth.NEUTRAL:
            logger.trace("character [{}] width is NEUTRAL", str);
            return false;
        case UCharacter.EastAsianWidth.HALFWIDTH:
            logger.trace("character [{}] width is HALFWIDTH", str);
            return false;
        case UCharacter.EastAsianWidth.AMBIGUOUS:
            // “”‘’ is AMBIGUOUS
            logger.trace("character [{}] width is AMBIGUOUS", str);
            return false;

        case UCharacter.EastAsianWidth.FULLWIDTH:
            logger.trace("character [{}] width is FULLWIDTH", str);
            return true;
        case UCharacter.EastAsianWidth.WIDE:
            logger.trace("character [{}] width is WIDE", str);
            return true;
        default:
            throw new RuntimeException("Unrecognize UProperty.EAST_ASIAN_WIDTH: " + value);
        }
    }

    public static int countLines(String str) {
        if (str.trim().isEmpty()) {
            return 0;
        }
        String[] lines = str.split("\r\n|\r|\n");
        return lines.length;
    }

    public static long countMaxWidthInLines(String str) {
        if (str.trim().isEmpty()) {
            return 0;
        }
        String[] lines = str.split("\r\n|\r|\n");
        long maxWidth = 0;
        for (String line : lines) {
            maxWidth = Util.getGreater(widthOfString(line), maxWidth);
        }
        return maxWidth;
    }

    public static String trimRight(String s) {
        int i = s.length() - 1;
        while (i >= 0 && Character.isWhitespace(s.charAt(i))) {
            i--;
        }
        return s.substring(0, i + 1);
    }

    /*-
     * Example of emacs table hline:
     *  +-----+-----+---------+
     * Example of org-mode table hline:
     *  |-----+-----+---------|
     */
    public static void printEmacsTableHline(StringBuilder sb, List<Integer> maxWidthInEachCol) {
        for (int i = 0; i < maxWidthInEachCol.size(); i++) {
            sb.append("+");
            for (int j = 0; j < maxWidthInEachCol.get(i) + 2; j++) {
                sb.append("-");
            }
        }
        sb.append("+\n");
    }

    /*-
     * Example of org-mode table hline:
     *  |-----+-----+---------|
     */
    public static void printOrgmodeTableHline(StringBuilder sb, List<Integer> maxWidthInEachCol) {
        for (int i = 0; i < maxWidthInEachCol.size(); i++) {
            sb.append("|");
            for (int j = 0; j < maxWidthInEachCol.get(i) + 2; j++) {
                sb.append("-");
            }
        }
        sb.append("|\n");
    }
}
