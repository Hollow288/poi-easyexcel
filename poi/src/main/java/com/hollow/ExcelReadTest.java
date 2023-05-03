//package com.hollow;
//
//import org.apache.poi.hssf.usermodel.HSSFCell;
//import org.apache.poi.hssf.usermodel.HSSFDateUtil;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.joda.time.DateTime;
//import org.junit.Test;
//
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.Date;
//
///**
// * @author 刘继涛
// * @version 1.0
// */
//public class ExcelReadTest {
//    String PATH = "F:\\Java .Projects\\easy-excel\\";
//
//
//    @Test
//    public void testRead03() throws IOException {
//        //获取文件流
//        FileInputStream fileInputStream = new FileInputStream(PATH + "观众统计表03.xls");
//        //1.创建一个工作簿,03版本，使用excel可以操作的这里都可以操作
//        Workbook workbook = new HSSFWorkbook(fileInputStream);
//        //得到表
//        Sheet sheet = workbook.getSheetAt(0);
//        //得到行
//        Row row = sheet.getRow(0);
//        //得到列
//        Cell cell = row.getCell(1);
//
//        //读取数值的时候一定要注意类型！
//        //获取字符串类型
////        System.out.println(cell.getStringCellValue());
//        //因为上面列要获取第二个列了，所以这里变成number
//        System.out.println(cell.getNumericCellValue());
//        fileInputStream.close();
//
//
//    }
//
//
//    @Test
//    public void testRead07() throws IOException {
//        //获取文件流
//        FileInputStream fileInputStream = new FileInputStream(PATH + "观众统计表07.xlsx");
//        //1.创建一个工作簿,03版本，使用excel可以操作的这里都可以操作
//        Workbook workbook = new XSSFWorkbook(fileInputStream);
//        //得到表
//        Sheet sheet = workbook.getSheetAt(0);
//        //得到行
//        Row row = sheet.getRow(0);
//        //得到列
//        Cell cell = row.getCell(1);
//
//        //读取数值的时候一定要注意类型！
//        //获取字符串类型
//        System.out.println(cell.getStringCellValue());
//        //因为上面列要获取第二个列了，所以这里变成number
////        System.out.println(cell.getNumericCellValue());
//        fileInputStream.close();
//
//
//    }
//
//    @Test
//    public void testCellType() throws IOException {
//        //获取文件流
//        FileInputStream fileInputStream = new FileInputStream(PATH + "观众统计表03.xls");
//
//        //1.创建一个工作簿，使用excel能操作的这边他都可以操作
//        Workbook workbook= new HSSFWorkbook(fileInputStream);
//        Sheet sheet = workbook.getSheetAt(0);
//        //获取标题内容
//        Row rowTitle = sheet.getRow(0);
//        if (rowTitle != null) {
//            //这个方法可以查询到这一行中列的数量
//            int cellCount = rowTitle.getPhysicalNumberOfCells();
//            //遍历这一行的每个
//            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                Cell cell = rowTitle.getCell(cellNum);
//                if (cell != null) {
//                    int cellType = cell.getCellType();
//                    String cellValue = cell.getStringCellValue();
//                    System.out.print(cellValue + "|");
//                }
//            }
//            System.out.println();
//        }
//
//        //获取表中的内容
//        //这里拿到了表中行的数量
//        int rowCount = sheet.getPhysicalNumberOfRows();
//        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
//            Row rowData = sheet.getRow(rowNum);
//            if (rowData != null) {
//                //读取行中的列
//                int cellCount = rowTitle.getPhysicalNumberOfCells();
//                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                    System.out.print("[" + (rowNum + 1) + "-" + (cellNum + 1) + "]");
//
//                    Cell cell = rowData.getCell(cellNum);
//                    String cellValue = "";
//                    //匹配列的数据类型
//                    if(cell != null) {
//                        int cellType = cell.getCellType();
//                        switch (cellType) {
//                            case Cell.CELL_TYPE_STRING://字符串
//                                System.out.print("[String]");
//                                cellValue = cell.getStringCellValue();
//                                break;
//                            case Cell.CELL_TYPE_BOOLEAN:// 布尔
//                                System.out.print("[BOOLEAN]");
//                                cellValue = String.valueOf(cell.getBooleanCellValue());
//                                break;
//                            case Cell.CELL_TYPE_BLANK: //空
//                                System.out.print("[BLANK]");
//                                break;
//                            case Cell.CELL_TYPE_NUMERIC: //数字(日期、普通数字)
//                                System.out.print("[NUMERIC]");
//                                if (HSSFDateUtil.isCellDateFormatted(cell)){ //日期
//                                    System.out.print("[日期]");
//                                    Date date = cell.getDateCellValue();
//                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
//                                }else {
//                                    //如果不是日期格式，防止数据过长
//                                    System.out.print("[转化字符串输出]");
//                                    cell.setCellType(Cell.CELL_TYPE_STRING);
//                                    cellValue = cell.toString();
//                                }
//                                break;
//                            case Cell.CELL_TYPE_ERROR: //错误
//                                System.out.print("[数据类型错误]");
//                                break;
//                        }
//                        System.out.println(cellValue);
//                    }else {
//                        System.out.print("[BLANK]");
//                        System.out.println(cellValue);
//                    }
//                }
//            }
//
//        }
//        fileInputStream.close();
//    }
//}
