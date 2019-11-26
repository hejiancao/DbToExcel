package com.example.demo.poi_2;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;

/**
 * @author created by shaos on 2019/11/26
 */
public class ExcelUtils {


    public static String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";//默认日期格式
    public static int DEFAULT_COLOUMN_WIDTH = 17;
    public static String[] p_static;
    public static String[] h_static;


    /**
     * @param titles      标题
     * @param headMaps    列名
     * @param jsonArray   数据集
     * @param datePattern 日期格式
     * @param colWidth    0
     * @param workbook    SXSSFWorkbook对象
     * @param sheetNum    sheet页码
     * @author shaos
     * @date 2019/11/25 16:51
     */
    public static void exportExcelByMultiSheet(String[] titles, Map<String, String>[] headMaps, JSONArray[] jsonArray, String datePattern, int colWidth, SXSSFWorkbook workbook, int sheetNum) {
        if (datePattern == null) {
            datePattern = DEFAULT_DATE_PATTERN;
        }
        workbook.setCompressTempFiles(true);
        //表头样式
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 20);
        titleFont.setBoldweight((short) 700);
        titleStyle.setFont(titleFont);
        // 列头样式
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerStyle.setFont(headerFont);
        // 单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        Font cellFont = workbook.createFont();
        cellFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        cellStyle.setFont(cellFont);
        // 生成一个(带标题)表格
        SXSSFSheet sheet = workbook.createSheet();
        sheet = assistant(headMaps[0], sheet, colWidth);
        workbook.setSheetName(sheetNum, titles[0] + sheetNum);
        // 遍历集合数据，产生数据行
        int rowIndex = 0;
        int currentCount = 0;
        boolean initSheet = false;
        while (currentCount < titles.length) {
            for (Object obj : jsonArray[currentCount]) {
                if (rowIndex - 1 > jsonArray[currentCount].size()) {
                    rowIndex = 0;
                    currentCount++;
                    initSheet = true;
                    break;
                }
                if (rowIndex == 65535 || rowIndex == 0) {
                    if (rowIndex != 0 || initSheet) {
                        sheet = workbook.createSheet();//如果数据超过了，则在第二页显示或操作对象发生变化
                        sheet = assistant(headMaps[currentCount], sheet, colWidth);
                        if (rowIndex != 0) {
                            workbook.setSheetName(workbook.getNumberOfSheets() - 1, titles[currentCount]);
                        } else {
                            workbook.setSheetName(currentCount, titles[currentCount]);
                        }
                    }

                    SXSSFRow titleRow = sheet.createRow(0);//表头 rowIndex=0
                    titleRow.createCell(0).setCellValue(titles[currentCount]);
                    titleRow.getCell(0).setCellStyle(titleStyle);
                    int size = headMaps[currentCount].size();
                    System.out.println(size);
                    //合并单元格 如果导出列数刚好是1列，则会抛出Merged Region异常 故添加单列导出判断
                    if (headMaps[currentCount].size() - 1 == 0) {
                        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
                    } else {
                        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headMaps[currentCount].size() - 1));
                    }

                    SXSSFRow headerRow = sheet.createRow(1); //列头 rowIndex =1
                    for (int i = 0; i < h_static.length; i++) {
                        headerRow.createCell(i).setCellValue(h_static[i]);
                        headerRow.getCell(i).setCellStyle(headerStyle);

                    }
                    rowIndex = 2;//数据内容从 rowIndex=2开始
                }
                JSONObject jo = (JSONObject) JSONObject.toJSON(obj);
                SXSSFRow dataRow = sheet.createRow(rowIndex);
                for (int i = 0; i < p_static.length; i++) {
                    SXSSFCell newCell = dataRow.createCell(i);

                    Object o = jo.get(p_static[i]);
                    String cellValue = "";
                    if (o == null) {
                        cellValue = "";
                    } else if (o instanceof Date) {
                        cellValue = new SimpleDateFormat(datePattern).format(o);
                    } else if (o instanceof Float || o instanceof Double) {
                        cellValue = new BigDecimal(o.toString()).setScale(2, BigDecimal.ROUND_HALF_UP).toString();
                    } else {
                        cellValue = o.toString();
                    }

                    newCell.setCellValue(cellValue);
                    newCell.setCellStyle(cellStyle);
                }
                rowIndex++;
            }
        }
    }


    public static SXSSFSheet assistant(Map<String, String> headMap, SXSSFSheet sheet, int colWidth) {
        //设置列宽
        int minBytes = colWidth < DEFAULT_COLOUMN_WIDTH ? DEFAULT_COLOUMN_WIDTH : colWidth;//至少字节数
        int[] arrColWidth = new int[headMap.size()];
        // 产生表格标题行,以及设置列宽
        p_static = new String[headMap.size()];
        h_static = new String[headMap.size()];
        int ii = 0;
        for (Iterator<String> iter = headMap.keySet().iterator(); iter
                .hasNext(); ) {
            String fieldName = iter.next();

            p_static[ii] = fieldName;
            h_static[ii] = headMap.get(fieldName);

            int bytes = fieldName.getBytes().length;
            arrColWidth[ii] = bytes < minBytes ? minBytes : bytes;
            sheet.setColumnWidth(ii, arrColWidth[ii] * 256);
            ii++;
        }
        return sheet;
    }
}
