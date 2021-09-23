package com.jeffrey.poiexcelstudy.simple;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by huangMP on 2017/8/20.
 * decription :
 */
public class OutputExcelDemo {

    /**
     *  从 工作簿中写入 数据 OutputExcelDemo
     * 07 版本及之前的版本写法
     * @throws IOException
     */
    public void outputExcel() throws IOException {
        // 1. 创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 2. 创建工作类
        HSSFSheet sheet = workbook.createSheet("hello world");

        // 3. 创建行 , 第三行 注意:从0开始
        HSSFRow row = sheet.createRow(2);

        // 4. 创建单元格, 第三行第三列 注意:从0开始
        HSSFCell cell = row.createCell(2);
        cell.setCellValue("Hello World");
        String projectPath = System.getProperty("user.dir");
        String fileName = projectPath+"\\document\\OutputExcelDemo.xls";
        FileOutputStream fileOutputSteam = new FileOutputStream(fileName);

        workbook.write(fileOutputSteam);
        workbook.close();

        fileOutputSteam.close();
    }


    /**
     * 从 工作簿中读取 数据 ReadExcelDemo
     * @throws IOException
     */
    public void readExel() throws IOException {

        String projectPath = System.getProperty("user.dir");
        String fileName = projectPath+"\\document\\OutputExcelDemo.xls";
        FileInputStream fileInputStream = new FileInputStream(fileName);

        // 1. 创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);

        // 2. 创建工作类
        HSSFSheet sheet = workbook.getSheetAt(0);

        // 3. 创建行 , 第三行 注意:从0开始
        HSSFRow row = sheet.getRow(2);

        // 4. 创建单元格, 第三行第三列 注意:从0开始
        HSSFCell cell = row.getCell(2);
        String cellString = cell.getStringCellValue();

        System.out.println("第三行第三列的值为 : " + cellString );

        workbook.close();
        fileInputStream.close();
    }


    public void testExcelStyle() throws IOException {
        // 1. 创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 1.1 创建单元格对象 合并第三行第三列到5列
        // 构造参数 起始行号 结束行号 起始列号 结束列号
        CellRangeAddress cellRangeAddress = new CellRangeAddress(
                2, 2, 2, 4 );
        // 1.2 创建单元格样式
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setAlignment(HSSFCellStyle.VERTICAL_CENTER);

        // 1.3 创建字体
        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints((short)16);
        // 将字体加载到样式中
        style.setFont(font);

        // 1.4 设置背景色为黄色
        // 1.4.1 设置填充模式
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillBackgroundColor(HSSFColor.YELLOW.index);
        style.setFillForegroundColor(HSSFColor.GREEN.index);

        // 2. 创建工作类
        HSSFSheet sheet = workbook.createSheet("hello world");
        // 2.1 加入合并单元格对象
        sheet.addMergedRegion(cellRangeAddress);

        // 3. 创建行 , 第三行 注意:从0开始
        HSSFRow row = sheet.createRow(2);

        // 4. 创建单元格, 第三行第三列 注意:从0开始
        HSSFCell cell = row.createCell(2);
        cell.setCellValue("Hello World");
        // 4.1 单元格添加样式
        cell.setCellStyle(style);
        String projectPath = System.getProperty("user.dir");
        String fileName = projectPath+"\\document\\HelloExcelStyle.xls";
        FileOutputStream fileOutputSteam = new FileOutputStream(fileName);

        workbook.write(fileOutputSteam);
        workbook.close();

        fileOutputSteam.close();

    }

}