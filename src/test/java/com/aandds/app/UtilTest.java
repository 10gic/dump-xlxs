package com.aandds.app;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class UtilTest {

    @Test
    public void testPadString() {
        assertEquals("     ", Util.padString("", 5));
        assertEquals("abc  ", Util.padString("abc", 5));
        assertEquals("abcde", Util.padString("abcde", 5));
        assertEquals("中文测试 ", Util.padString("中文测试", 9));
        assertEquals("abc中英文测试", Util.padString("abc中英文测试", 13));
        assertEquals("abc中英文测试  ", Util.padString("abc中英文测试", 15));

        // 全角字母测试
        assertEquals("ａｂｃｏｍ", Util.padString("ａｂｃｏｍ", 10));
        assertEquals("ａｂｃｏｍ ", Util.padString("ａｂｃｏｍ", 11));

        // 中文标点测试
        assertEquals("、，。", Util.padString("、，。", 6));
    }

    @Test
    public void testWidthOfString() {
        assertEquals(1, Util.widthOfString("1"));
        assertEquals(1, Util.widthOfString("a"));
        assertEquals(3, Util.widthOfString("abc"));
        assertEquals(8, Util.widthOfString("中文测试"));
        assertEquals(13, Util.widthOfString("abc中英文测试"));

        // 全角数字测试
        assertEquals(2, Util.widthOfString("１"));
        assertEquals(12, Util.widthOfString("１２３４５６"));

        // 全角字母测试
        assertEquals(2, Util.widthOfString("ａ"));
        assertEquals(6, Util.widthOfString("ａｂｃ"));

        // 中文标点、，。测试
        assertEquals(2, Util.widthOfString("、"));
        assertEquals(2, Util.widthOfString("，"));
        assertEquals(2, Util.widthOfString("。"));

        // 中文标点“”‘’测试
        assertEquals(1, Util.widthOfString("“"));
        assertEquals(1, Util.widthOfString("”"));
        assertEquals(1, Util.widthOfString("‘"));
        assertEquals(1, Util.widthOfString("’"));
    }

    @Test
    public void testCountMaxWidthInLines() {
        assertEquals(3, Util.countMaxWidthInLines("123"));
        assertEquals(4, Util.countMaxWidthInLines("123\n1234"));
        assertEquals(8, Util.countMaxWidthInLines("中文测试"));
        assertEquals(13, Util.countMaxWidthInLines("abc中英文测试"));
        assertEquals(13, Util.countMaxWidthInLines("abc中英文测试\n1234"));
    }
}
