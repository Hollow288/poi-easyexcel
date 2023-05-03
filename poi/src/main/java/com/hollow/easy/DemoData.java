package com.hollow.easy;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.Data;

import java.util.Date;

/**
 * @author 刘继涛
 * @version 1.0
 */
@Data
public class DemoData {

    @ExcelProperty("字符串标题")
    private String string;

//    /**
//     * 这里用string 去接日期才能格式化。我想接收年月日格式
//     */
//    @DateTimeFormat("yyyy年MM月dd日HH时mm分ss秒")
    @ExcelProperty("日期标题")
    private Date date;

    @ExcelProperty("数字标题")
    private Double doubleData;

    /**
     * 忽略这个字段
     */

    @ExcelIgnore
    private String ignore;


}
