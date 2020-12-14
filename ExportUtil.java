package com.genuine.mes.system.util;

import cn.hutool.core.map.MapUtil;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * @className: ExportUtil
 * @description:
 * @author: wwj
 *
 */
public class ExportUtil {

    private ExportUtil() {
    }

    /**
     * 导出excel，xlsx格式
     *
     * @param sheetName sheet名称, 参数类型：数组
     * @param title     标题, 参数类型：数组
     * @param list      需要导出内容数据, 参数类型：List<List<String>> 需要将所有数据转换成字符串类型
     * @param fileName  excel文件名，参数类型：字符串183
     */
    @SneakyThrows
    public static void getSXSSFWorkbook(HttpServletResponse response, String sheetName, String[] title,
                                        List<List<String>> list, String fileName, String userName, String companyName) {

        SXSSFWorkbook workbook = new SXSSFWorkbook();
        // 设置日期格式
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        // 在workbook中添加一个sheet,对应Excel文件中的sheet
        Sheet sheet = workbook.createSheet(sheetName);

        // 在sheet中添加表头第0行
        Row row = null;

        // 声明列对象
        Cell cell = null;
        // ----------------一级标题样式---------------------
        CellStyle titleStyle = getCellStyle1(workbook);

        // ----------------二级标题格样式----------------------------------
        CellStyle titleStyle2 = getCellStyle2(workbook);

        // 创建标题单元格样式
        CellStyle style = getCellStyle(workbook);

        /* 设置内容样式 */
        CellStyle tableStyle = getTableCellStyle(workbook);

        // ----------------------创建第一行---------------
        // 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        row = sheet.createRow(0);
        // 创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
        cell = row.createCell(0);
        // 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(0, 3, 0, title.length - 1));
        // 设置单元格内容
        cell.setCellValue(companyName);
        cell.setCellStyle(titleStyle);

        // ------------------创建第二行(单位、填表日期)---------------------
        int num = title.length - 1;
        row = sheet.createRow(4);
        cell = row.createCell(0);
        cell.setCellStyle(titleStyle2);
        cell.setCellValue("打印人：" + userName);
        // 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(4, 5, 0, (num - 1) / 2));
        cell = row.createCell((num - 1) / 2 + 1);
        sheet.addMergedRegion(new CellRangeAddress(4, 5, (num - 1) / 2 + 1, num));
        cell.setCellValue("打印时间：" + df.format(new Date()));
        cell.setCellStyle(titleStyle2);

        // 创建标题
        row = sheet.createRow(6);
        cell = row.createCell(0);
        cell.setCellStyle(style);
        for (int i = 0; i < title.length; i++) {
            // 设置列宽
            sheet.setColumnWidth(i, 20 * 256);
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        // 创建内容
        for (int i = 6; i < list.size() + 6; i++) {
            row = sheet.createRow(i + 1);
            // 将内容按顺序赋给对应的列对象
            for (int j = 0; j < list.get(i - 6).size(); j++) {
                Cell cellll = row.createCell(j);
                cellll.setCellValue(list.get(i - 6).get(j));
                cellll.setCellStyle(tableStyle);
            }
        }
        try {
            /* 响应到客户端 */
            setResponseHeader(response, fileName + ".xlsx");
            OutputStream os = response.getOutputStream();
            workbook.write(os);
            os.flush();
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 导出pdf文档
     *
     * @param lists    导出数据,注: 表头需要添加到lists中第一条
     * @param response 响应对象121
     */
    @SneakyThrows
    public static void downloadPdf(HttpServletResponse response, List<List<String>> lists, String fileName,
                                   String userName, String companyName) {
        response.setCharacterEncoding("utf-8");
        fileName = URLEncoder.encode(fileName, "UTF-8");
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/pdf");
        // 下载文件的默认名称
        response.setHeader("Content-disposition",
                "attachment;filename=" + fileName + ".pdf" + ";" + "filename*=utf-8''" + fileName + ".pdf");
        BaseFont baseFont = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", false);
        // 设置日期格式
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        // 自定义字体属性
        com.itextpdf.text.Font font = new com.itextpdf.text.Font(baseFont, 10);
        com.itextpdf.text.Font font2 = new com.itextpdf.text.Font(baseFont, 11);
        com.itextpdf.text.Font titleFont = new com.itextpdf.text.Font(baseFont, 16);
        com.itextpdf.text.Font titleFont2 = new com.itextpdf.text.Font(baseFont, 13);
        Document document = new Document(PageSize.A4, 50, 50, 30, 20);
        PdfWriter.getInstance(document, response.getOutputStream());
        document.open();
        // 段落p0,p1
        Paragraph p0 = new Paragraph();

        Paragraph p1 = new Paragraph(companyName, titleFont);
        // 中间对齐
        p1.setAlignment(Element.ALIGN_CENTER);
        // 设置浮动
        p1.setLeading(40);
        p0.add(p1);
        String str = "打印人：" + userName + "         " + "打印时间:" + df.format(new Date());
        Paragraph p2 = new Paragraph(str, titleFont2);
        // 中间对齐
        p2.setAlignment(Element.ALIGN_RIGHT);
        // 设置浮动
        p2.setLeading(40);
        p2.setSpacingAfter(10f);
        p0.add(p2);
        document.add(p0);
        if (lists.size() > 0) {
            for (int i = 0; i < lists.size(); i++) {
                PdfPTable table = new PdfPTable(lists.get(i).size());
                // 表格在页面中所占的宽度百分比
                table.setWidthPercentage(100);
                // 表格的全部宽度。
                table.setTotalWidth(PageSize.A4.getWidth() - 10);
                table.setLockedWidth(true);
                PdfPCell cell = new PdfPCell();
                if (i == 0) {
                    for (int j = 0; j < lists.get(i).size(); j++) {
                        Paragraph p = new Paragraph(lists.get(i).get(j), font2);
                        p.setFont(titleFont);
                        // 水平居中
                        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        cell.setPhrase(p);
                        cell.setBackgroundColor(new BaseColor(204, 204, 204));
                        // 文档中加入该段落
                        table.addCell(cell);
                        document.add(table);
                    }
                } else {
                    for (int j = 0; j < lists.get(i).size(); j++) {
                        Paragraph p = new Paragraph(lists.get(i).get(j), font);
                        p.setFont(font);
                        // 设置段落居中，其中1为居中对齐，2为右对齐，3为左对齐
                        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        cell.setPhrase(p);
                        table.addCell(cell);
                        document.add(table);
                    }
                }
            }
            document.close();
        }
    }

    private static CellStyle getCellStyle(SXSSFWorkbook workbook) {
        /* 创建标题单元格样式 */
        CellStyle style = workbook.createCellStyle();
        /* 设置单元格填充样式，SOLID_FOREGROUND纯色使用前景颜色填充 */
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        /* 设置填充颜色 */
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        /* 水平居中 */
        style.setAlignment(HorizontalAlignment.CENTER);
        /* 垂直对齐居中 */
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        /* 自动换行 */
        style.setWrapText(true);
        /* 创建字体样式 */
        XSSFFont font = (XSSFFont) workbook.createFont();
        /* 设置字体大小 */
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        return style;
    }

    private static CellStyle getCellStyle1(SXSSFWorkbook workbook) {
        // ----------------一级标题样式---------------------
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFFont ztFont = (XSSFFont) workbook.createFont();
        // 设置字体为斜体字
        ztFont.setItalic(false);
        // 将字体设置为“红色”
        ztFont.setColor(Font.COLOR_NORMAL);
        // 将字体大小设置为18px
        ztFont.setFontHeightInPoints((short) 18);
        // 将“宋体”字体应用到当前单元格上
        ztFont.setFontName("宋体");
        // 加粗
        ztFont.setBold(true);
        titleStyle.setFont(ztFont);
        return titleStyle;
    }

    private static CellStyle getTableCellStyle(SXSSFWorkbook workbook) {
        // 创建内容单元格样式
        CellStyle tableStyle = workbook.createCellStyle();
        /* 水平居中 */
        tableStyle.setAlignment(HorizontalAlignment.CENTER);
        /* 垂直对齐居中 */
        tableStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        /* 自动换行 */
        tableStyle.setWrapText(true);
        return tableStyle;
    }

    private static CellStyle getCellStyle2(SXSSFWorkbook workbook) {
        // 二级标题格样式
        CellStyle titleStyle2 = workbook.createCellStyle();
        titleStyle2.setAlignment(HorizontalAlignment.CENTER);
        titleStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFFont ztFont2 = (XSSFFont) workbook.createFont();
        // 设置字体为斜体字
        ztFont2.setItalic(false);
        // 将字体设置为“红色”
        ztFont2.setColor(Font.COLOR_NORMAL);
        // 将字体大小设置为18px
        ztFont2.setFontHeightInPoints((short) 13);
        // 字体应用到当前单元格上
        ztFont2.setFontName("宋体");
        // 加粗
        ztFont2.setBold(true);
        titleStyle2.setFont(ztFont2);
        return titleStyle2;
    }

    /**
     * 发送响应流方法
     *
     * @param response 响应
     * @param fileName excel文件名263 * @since 2019-05-10
     */
    private static void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            response.setCharacterEncoding("utf-8");
            fileName = URLEncoder.encode(fileName, "UTF-8");
            response.setHeader("content-Type", "application/x-xls");
            response.setHeader("Content-disposition",
                    "attachment;filename=" + fileName + ";" + "filename*=utf-8''" + fileName);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    /**
     * 查阅导出excel，xlsx格式(导出单个查阅的数据)
     *
     * @param sheetName         sheet名称, 参数类型：数组
     * @param title             标题, 参数类型：数组
     * @param list1,list2,list3 需要导出内容数据, 参数类型：List<List<String>>
     *                          需要将所有数据转换成字符串类型
     * @param fileName          excel文件名，参数类型：字符串183
     */
    @SneakyThrows
    public static void getCHSXSSFWorkbook(HttpServletResponse response, String sheetName, String[] title,
                                          String[] title1, String[] title2, List<String> list0, List<List<String>> list1, List<List<String>> list2,
                                          String fileName, String userName, String companyName) {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        // 设置日期格式
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        // 在workbook中添加一个sheet,对应Excel文件中的sheet
        Sheet sheet = workbook.createSheet(sheetName);

        // 在sheet中添加表头第0行
        Row row = null;

        // 声明列对象
        Cell cell = null;
        // ----------------一级标题样式---------------------
        CellStyle titleStyle = getCellStyle1(workbook);

        // ----------------二级标题格样式----------------------------------
        CellStyle titleStyle2 = getCellStyle2(workbook);

        // 创建标题单元格样式
        CellStyle style = getCellStyle(workbook);

        /* 设置内容样式 */
        CellStyle tableStyle = getTableCellStyle(workbook);

        // ----------------------创建第一行---------------
        // 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        row = sheet.createRow(0);
        // 创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
        cell = row.createCell(0);
        // 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(0, 3, 0, title.length - 1));
        // 设置单元格内容
        cell.setCellValue(companyName);
        cell.setCellStyle(titleStyle);

        // ------------------创建第二行(单位、填表日期)---------------------
        int num = title.length - 1;
        row = sheet.createRow(4);
        cell = row.createCell(0);
        cell.setCellStyle(titleStyle2);
        cell.setCellValue("打印人：" + userName);
        // 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(4, 5, 0, (num - 1) / 2));
        cell = row.createCell((num - 1) / 2 + 1);
        sheet.addMergedRegion(new CellRangeAddress(4, 5, (num - 1) / 2 + 1, num));
        cell.setCellValue("打印时间：" + df.format(new Date()));
        cell.setCellStyle(titleStyle2);

        // 创建标题1
        row = sheet.createRow(6);
        cell = row.createCell(0);
        cell.setCellStyle(style);
        for (int i = 0; i < title.length; i++) {
            // 设置列宽
            sheet.setColumnWidth(i, 20 * 256);
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        // 创建内容1
        row = sheet.createRow(7);
        for (int i = 0; i < list0.size(); i++) {
            Cell cel1 = row.createCell(i);
            cel1.setCellValue(list0.get(i));
            cel1.setCellStyle(tableStyle);
        }
        // 创建标题2
        row = sheet.createRow(8);
        cell = row.createCell(0);
        cell.setCellStyle(style);
        for (int i = 0; i < title1.length; i++) {
            // 设置列宽
            sheet.setColumnWidth(i, 20 * 256);
            cell = row.createCell(i);
            cell.setCellValue(title1[i]);
            cell.setCellStyle(style);
        }
        // 创建内容2
        for (int i = 9; i < list1.size() + 9; i++) {
            row = sheet.createRow(i);
            // 将内容按顺序赋给对应的列对象
            for (int j = 0; j < list1.get(i - 9).size(); j++) {
                Cell cellll = row.createCell(j);
                cellll.setCellValue(list1.get(i - 9).get(j));
                cellll.setCellStyle(tableStyle);
            }
        }
        // 创建标题3
        row = sheet.createRow(10 + list1.size());
        cell = row.createCell(0);
        cell.setCellStyle(style);
        for (int i = 0; i < title2.length; i++) {
            // 设置列宽
            sheet.setColumnWidth(i, 20 * 256);
            cell = row.createCell(i);
            cell.setCellValue(title2[i]);
            cell.setCellStyle(style);
        }

        // 创建内容3
        int t = 11 + list1.size();
        for (int i = t; i < list2.size() + t; i++) {
            row = sheet.createRow(i);
            // 将内容按顺序赋给对应的列对象
            for (int j = 0; j < list2.get(i - t).size(); j++) {
                Cell cellll = row.createCell(j);
                cellll.setCellValue(list2.get(i - t).get(j));
                cellll.setCellStyle(tableStyle);
            }
        }
        try {
            /* 响应到客户端 */
            setResponseHeader(response, fileName + ".xlsx");
            OutputStream os = response.getOutputStream();
            workbook.write(os);
            os.flush();
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

 
    /**
     * 查阅导出excel，xlsx格式(导出单个查阅的数据)
     *
     * @param sheetName sheet名称, 参数类型：数组
     * @param fileName  excel文件名，参数类型：字符串183
     */
    @SneakyThrows
    public static void exportExel(HttpServletResponse response, String sheetName, LinkedHashMap<String[], List<List<String>>> map,
                                  String fileName, String userName, String companyName) {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        // 设置日期格式
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        // 在workbook中添加一个sheet,对应Excel文件中的sheet
        Sheet sheet = workbook.createSheet(sheetName);
        // 在sheet中添加表头第0行
        Row row = null;
        // 声明列对象
        Cell cell = null;
        // ----------------一级标题样式---------------------
        CellStyle titleStyle = getCellStyle1(workbook);
        // ----------------二级标题格样式----------------------------------
        CellStyle titleStyle2 = getCellStyle2(workbook);
        // 创建标题单元格样式
        CellStyle style = getCellStyle(workbook);
        /* 设置内容样式 */
        CellStyle tableStyle = getTableCellStyle(workbook);       
         int num=new ArrayList<String[]>(map.keySet()).get(0).length-1;
        // ----------------------创建第一行---------------
        // 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        row = sheet.createRow(0);
        // 创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
        cell = row.createCell(0);
        // 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(0, 3, 0, num));
        // 设置单元格内容
        cell.setCellValue(companyName);
        cell.setCellStyle(titleStyle);

        // ------------------创建第二行(单位、填表日期)---------------------
        
        row = sheet.createRow(4);
        cell = row.createCell(0);
        sheet.addMergedRegion(new CellRangeAddress(4, 5, 0, num));
        cell.setCellStyle(titleStyle2);
        cell.setCellValue("打印人：" + userName+" "+" "+"打印时间：" + df.format(new Date()));
        // 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列                   
        int rowNum = 6;
        if (MapUtil.isNotEmpty(map)) {
            for (Map.Entry<String[], List<List<String>>> entry : map.entrySet()) {
                String[] title = entry.getKey();
                List<List<String>> values = entry.getValue();
                // 创建标题
                row = sheet.createRow(rowNum);
                cell = row.createCell(0);
                cell.setCellStyle(style);
                for (int i = 0; i < title.length; i++) {
                    // 设置列宽
                    sheet.setColumnWidth(i, 20 * 256);
                    cell = row.createCell(i);
                    cell.setCellValue(title[i]);
                    cell.setCellStyle(style);
                }
                rowNum++;
                // 创建内容
                for (int i = rowNum; i < values.size() + rowNum; i++) {
                    row = sheet.createRow(i);
                    // 将内容按顺序赋给对应的列对象
                    for (int j = 0; j < values.get(i - rowNum).size(); j++) {
                        Cell cellll = row.createCell(j);
                        cellll.setCellValue(values.get(i - rowNum).get(j));
                        cellll.setCellStyle(tableStyle);
                    }
                }
                rowNum = rowNum + values.size();
            }
        } else {
            throw new RuntimeException("导出内容不能为空！");
        }

        try {
            /* 响应到客户端 */
            setResponseHeader(response, fileName + ".xlsx");
            OutputStream os = response.getOutputStream();
            workbook.write(os);
            os.flush();
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 查询导出pdf文档(查询单个时使用)
     *
     * @param response 响应对象
     */
    @SneakyThrows
	public static void exportPdf(HttpServletResponse response, String fileName, String userName, String companyName, LinkedHashMap<String[], List<List<String>>> map) {
        response.setCharacterEncoding("utf-8");
        fileName = URLEncoder.encode(fileName, "UTF-8");
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/pdf");
        // 下载文件的默认名称
        response.setHeader("Content-disposition",
                "attachment;filename=" + fileName + ".pdf" + ";" + "filename*=utf-8''" + fileName + ".pdf");
        BaseFont baseFont = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", false);
        // 设置日期格式
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        // 自定义字体属性
        com.itextpdf.text.Font font = new com.itextpdf.text.Font(baseFont, 10);
        com.itextpdf.text.Font font2 = new com.itextpdf.text.Font(baseFont, 11);
        com.itextpdf.text.Font titleFont = new com.itextpdf.text.Font(baseFont, 16);
        com.itextpdf.text.Font titleFont2 = new com.itextpdf.text.Font(baseFont, 13);
        Document document = new Document(PageSize.A4, 50, 50, 30, 20);
        PdfWriter.getInstance(document, response.getOutputStream());
        document.open();
        // 段落p0,p1
        Paragraph p0 = new Paragraph();
        Paragraph p1 = new Paragraph(companyName, titleFont);
        // 中间对齐
        p1.setAlignment(Element.ALIGN_CENTER);
        // 设置浮动
        p1.setLeading(40);
        p0.add(p1);
        String str = "打印人：" + userName + "         " + "打印时间:" + df.format(new Date());
        Paragraph p2 = new Paragraph(str, titleFont2);
        // 中间对齐
        p2.setAlignment(Element.ALIGN_RIGHT);
        // 设置浮动
        p2.setLeading(40);
        p2.setSpacingAfter(10f);
        p0.add(p2);
        document.add(p0);
        if (MapUtil.isNotEmpty(map)) {
            for (Map.Entry<String[], List<List<String>>> entry : map.entrySet()) {
                //列名标题
                String[] title = entry.getKey();
                //内容
                List<List<String>> values = entry.getValue();

                PdfPTable table = new PdfPTable(title.length);
                // 表格在页面中所占的宽度百分比
                table.setWidthPercentage(100);
                // 表格的全部宽度。
                table.setTotalWidth(PageSize.A4.getWidth() - 10);
                table.setLockedWidth(true);
                // 创建标题
                for (int k = 0; k < title.length; k++) {
                    Paragraph p = new Paragraph(title[k], font2);
                    p.setFont(titleFont);
                    PdfPCell cell = new PdfPCell();
                    // 水平居中
                    cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                    cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    cell.setPhrase(p);
                    cell.setBackgroundColor(new BaseColor(204, 204, 204));
                    // 文档中加入该段落
                    table.addCell(cell);
                }
                for (int i = 0; i < values.size(); i++) {
                    // 创建内容
                    for (int j = 0; j < values.get(i).size(); j++) {
                        Paragraph p = new Paragraph(values.get(i).get(j), font);
                        p.setFont(font);
                        PdfPCell cell = new PdfPCell();
                        // 设置段落居中，其中1为居中对齐，2为右对齐，3为左对齐
                        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        cell.setPhrase(p);
                        table.addCell(cell);
                    }
                }
                document.add(table);
            }
        } else {
            throw new RuntimeException("导出内容不能为空！");
        }
        document.close();

    }
}
