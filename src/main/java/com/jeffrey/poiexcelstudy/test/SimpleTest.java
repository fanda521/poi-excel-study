package com.jeffrey.poiexcelstudy.test;

import com.jeffrey.poiexcelstudy.simple.OutputExcelDemo;
import org.junit.Test;

import java.io.IOException;

/**
 * @version 1.0
 * @Aythor jeffrey 王吉慧
 * @date 2021/9/23 14:51
 * @description
 */
public class SimpleTest {

    @Test
    public void outputExcel() throws Exception {
        OutputExcelDemo outputExcelDemo = new OutputExcelDemo();
        outputExcelDemo.outputExcel();
    }

    @Test
    public void readExel() throws IOException {
        OutputExcelDemo outputExcelDemo = new OutputExcelDemo();
        outputExcelDemo.readExel();
    }

    @Test
    public void testExcelStyle() throws IOException {
        OutputExcelDemo outputExcelDemo = new OutputExcelDemo();
        outputExcelDemo.testExcelStyle();
    }

}
