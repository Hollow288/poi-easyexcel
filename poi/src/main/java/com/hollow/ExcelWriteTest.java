package com.hollow;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author 刘继涛
 * @version 1.0
 */
public class ExcelWriteTest {

    //注意路径最后的\\
    String PATH = "F:\\Java .Projects\\easy-excel\\";

    @Test
    public void testWrite03() throws IOException {
        //1.创建一个工作簿,03版本
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("统计表");
        //3.创建一个行,从0行开始
        Row row1 = sheet.createRow(0);
        //第一行的第一列的那个格子
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        //第一行的第二列的那个格子
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //第二行
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表(IO)流 现在是03版本 就必须以xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "观众统计表03.xls");
        //输出
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();

        System.out.println("03文件生成完毕");


    }

    @Test
    public void testWrite07() throws IOException {
        //1.创建一个工作簿,07版本
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("统计表");
        //3.创建一个行,从0行开始
        Row row1 = sheet.createRow(0);
        //第一行的第一列的那个格子
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        //第一行的第二列的那个格子
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //第二行
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表(IO)流 现在是03版本 就必须以xlsx结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "观众统计表07.xlsx");
        //输出
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();

        System.out.println("07文件生成完毕");


    }

    @Test
    public void testWrite03BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            //这里创建了当前行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                //再来一个循环，创建当前行的列，这就是每行的单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite03BigData.xls");
        workbook.write(fileOutputStream);

        System.out.println("03大数据表生成完毕");

        long end = System.currentTimeMillis();

        System.out.println((double)(end-begin)/1000);

    }

    @Test
    public void testWrite07BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            //这里创建了当前行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                //再来一个循环，创建当前行的列，这就是每行的单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite07BigData.xlsx");
        workbook.write(fileOutputStream);

        System.out.println("07大数据表生成完毕");

        long end = System.currentTimeMillis();

        System.out.println((double)(end-begin)/1000);

    }

    @Test
    public void testWrite07BigDataS() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new SXSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            //这里创建了当前行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                //再来一个循环，创建当前行的列，这就是每行的单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite07BigDataS.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();

        System.out.println("07大数据表缓冲版生成完毕");

        long end = System.currentTimeMillis();

        System.out.println((double)(end-begin)/1000);

    }

}
