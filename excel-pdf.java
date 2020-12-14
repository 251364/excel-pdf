 1 package com.rexen.rest.common.util;
  2 
  3 
  4 import com.itextpdf.text.*;
  5 import com.itextpdf.text.html.simpleparser.HTMLWorker;
  6 import com.itextpdf.text.pdf.*;
  7 import org.apache.poi.hssf.usermodel.HSSFCellStyle;
  8 import org.apache.poi.hssf.usermodel.HSSFFont;
  9 import org.apache.poi.hssf.usermodel.HSSFWorkbook;
 10 import org.apache.poi.ss.usermodel.*;
 11 import org.apache.poi.xssf.streaming.SXSSFWorkbook;
 12 import org.apache.poi.xssf.usermodel.XSSFFont;
 13 
 14 import javax.servlet.http.HttpServletRequest;
 15 import javax.servlet.http.HttpServletResponse;
 16 import java.io.IOException;
 17 import java.io.OutputStream;
 18 import java.io.StringReader;
 19 import java.io.UnsupportedEncodingException;
 20 import java.net.URLEncoder;
 21 import java.util.List;
 22 
 23 /**
 24  * 导出表格
 25  * 27  * @since 2019-05-10
 28  */
 29 public class Export {
 30     /**
 31      * 导出excel，xls格式
 32      *
 33      * @param sheetName sheet名称
 34      * @param title     标题, 参数类型：数组
 35      * @param list      需要导出内容数据, 参数类型：List<List<String>> 需要将所有数据转换成字符串类型
 36      * @param fileName  excel文件名，参数类型：字符串 38      * @since 2019-05-10
 39      */
 40     public static void getHSSFWorkbook(HttpServletResponse response, HttpServletRequest request, String sheetName, String[] title, List<List<String>> list, String fileName) {
 41         /* 获取文件名称后缀*/
 42         final String suffixName = fileName.substring(fileName.lastIndexOf(".") + 1);
 43 
 44         /* 创建HSSFWorkbook，对应一个Excel文件*/
 45         HSSFWorkbook workbook = new HSSFWorkbook();
 46         /* 在workbook中添加一个sheet,对应Excel文件中的sheet*/
 47         Sheet sheet = workbook.createSheet(sheetName);
 48         /* 在sheet中添加表头第0行*/
 49         Row row = sheet.createRow(0);
 50         /* 创建单元格*/
 51         Cell cell = null;
 52 
 53         /* 创建单元格样式*/
 54         CellStyle style = workbook.createCellStyle();
 55         /*设置单元格填充样式，SOLID_FOREGROUND纯色使用前景颜色填充*/
 56         style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
 57         /* 设置填充颜色*/
 58         style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
 59         /* 水平居中*/
 60         style.setAlignment(HorizontalAlignment.CENTER);
 61         /*垂直对齐居中*/
 62         style.setVerticalAlignment(VerticalAlignment.CENTER);
 63         /* 自动换行*/
 64         style.setWrapText(true);
 65         /*创建字体样式*/
 66         HSSFFont font = workbook.createFont();
 67         /*设置字体大小*/
 68         font.setFontHeightInPoints((short) 14);
 69         style.setFont(font);
 70 
 71         /* 设置内容样式*/
 72         HSSFCellStyle tableStyle = workbook.createCellStyle();
 73         /* 水平居中*/
 74         tableStyle.setAlignment(HorizontalAlignment.CENTER);
 75         /* 垂直对齐居中*/
 76         tableStyle.setVerticalAlignment(VerticalAlignment.CENTER);
 77         /* 自动换行*/
 78         tableStyle.setWrapText(true);
 79 
 80         /* 创建表头*/
 81         for (int i = 0; i < title.length; i++) {
 82             // 设置列宽
 83             sheet.setColumnWidth(i, 20 * 256);
 84             cell = row.createCell(i);
 85             cell.setCellValue(title[i]);
 86             cell.setCellStyle(style);
 87         }
 88 
 89         /* 创建内容*/
 90         for (int i = 0; i < list.size(); i++) {
 91             row = sheet.createRow(i + 1);
 92             for (int j = 0; j < list.get(i).size(); j++) {
 93                 /* 将内容按顺序赋给对应的列对象*/
 94                 Cell tableCell = row.createCell(j);
 95                 tableCell.setCellValue(list.get(i).get(j));
 96                 tableCell.setCellStyle(tableStyle);
 97             }
 98         }
 99         try {
100             /* 响应到客户端*/
101             setResponseHeader(response, fileName + ".xls");
102             OutputStream os = response.getOutputStream();
103             workbook.write(os);
104             os.flush();
105             os.close();
106 
107         } catch (IOException e) {
108             e.printStackTrace();
109         }
110 
111     }
112 
113 
114     /**
115      * 导出pdf文档
116      *
117      * @param lists    导出数据,注: 表头需要添加到lists中第一条
118      * @param request  请求对象
119      * @param response 响应对象121      * @since 2019-05-10
122      */
123     // 下载pdf文档
124     public static void download(HttpServletRequest request, HttpServletResponse response, List<List<String>> lists, String fileName) throws Exception {
125         // 告诉浏览器用什么软件可以打开此文件
126         response.setHeader("content-Type", "application/pdf");
127         // 下载文件的默认名称
128         response.setHeader("Content-Disposition","attachment;fileName=" +URLEncoder.encode(fileName + ".pdf", "UTF-8"));
129         BaseFont baseFont = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", false);
130         // 自定义字体属性
131         com.itextpdf.text.Font font = new com.itextpdf.text.Font(baseFont, 11);
132         com.itextpdf.text.Font titleFont = new com.itextpdf.text.Font(baseFont, 14);
133         Document document = new Document(PageSize.A4);
134         PdfWriter.getInstance(document, response.getOutputStream());
135         document.open();
136         if (lists.size() > 0 && lists != null) {
137             for (int i = 0; i < lists.size(); i++) {
138                 PdfPTable table = new PdfPTable(lists.get(i).size());
139                 table.setWidthPercentage(100);
140                 PdfPCell cell = new PdfPCell();
141                 if (i == 0) {
142                     if (lists.get(i).size() > 0 && lists.get(i) != null) {
143                         for (int j = 0; j < lists.get(i).size(); j++) {
144                             System.out.println(lists.get(i).get(j));
145                             Paragraph p=new Paragraph(lists.get(i).get(j), titleFont);
146                             p.setFont(titleFont);
147                             // 水平居中
148                             cell.setHorizontalAlignment(Element.ALIGN_CENTER);
149                             cell.setPhrase(p);
150                             cell.setBackgroundColor(new BaseColor(204, 204, 204));
151                             // 文档中加入该段落
152                             table.addCell(cell);
153                             document.add(table);
154                         }
155                     }
156                 }else {
157                     if (lists.get(i).size() > 0 && lists.get(i) != null) {
158                         for (int j = 0; j < lists.get(i).size(); j++) {
159                             System.out.println(lists.get(i).get(j));
160                             Paragraph p=new Paragraph(lists.get(i).get(j),font);
161                             p.setFont(font);
162                             // 设置段落居中，其中1为居中对齐，2为右对齐，3为左对齐
163                             cell.setHorizontalAlignment(Element.ALIGN_CENTER);
164                             cell.setPhrase(p);
165                             table.addCell(cell);
166                             document.add(table);
167                         }
168                     }
169                 }
170             }
171             document.close();
172         }
173     }
174 
175     /**
176      * 导出excel，xlsx格式
177      *
178      * @param sheetName sheet名称, 参数类型：数组
179      * @param title     标题, 参数类型：数组
180      * @param list      需要导出内容数据, 参数类型：List<List<String>> 需要将所有数据转换成字符串类型
181      * @param fileName  excel文件名，参数类型：字符串183      * @since 2019-05-10
184      */
185     public static void getSXSSFWorkbook(HttpServletResponse response, HttpServletRequest request, String sheetName, String[] title, List<List<String>> list, String fileName) {
186 
187         SXSSFWorkbook workbook = new SXSSFWorkbook();
188 
189         // 在workbook中添加一个sheet,对应Excel文件中的sheet
190         Sheet sheet = workbook.createSheet(sheetName);
191 
192         // 在sheet中添加表头第0行
193         Row row = sheet.createRow(0);
194 
195         // 声明列对象
196         Cell cell = null;
197 
198         /* 创建单元格样式*/
199         CellStyle style = workbook.createCellStyle();
200         /*设置单元格填充样式，SOLID_FOREGROUND纯色使用前景颜色填充*/
201         style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
202         /* 设置填充颜色*/
203         style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
204         /* 水平居中*/
205         style.setAlignment(HorizontalAlignment.CENTER);
206         /*垂直对齐居中*/
207         style.setVerticalAlignment(VerticalAlignment.CENTER);
208         /* 自动换行*/
209         style.setWrapText(true);
210         /*创建字体样式*/
211         XSSFFont font = (XSSFFont) workbook.createFont();
212         /*设置字体大小*/
213         font.setFontHeightInPoints((short) 14);
214         style.setFont(font);
215 
216         /* 设置内容样式*/
217         CellStyle tableStyle = workbook.createCellStyle();
218         /* 水平居中*/
219         tableStyle.setAlignment(HorizontalAlignment.CENTER);
220         /* 垂直对齐居中*/
221         tableStyle.setVerticalAlignment(VerticalAlignment.CENTER);
222         /* 自动换行*/
223         tableStyle.setWrapText(true);
224 
225         // 创建标题
226         for (int i = 0; i < title.length; i++) {
227             // 设置列宽
228             sheet.setColumnWidth(i, 20 * 256);
229             cell = row.createCell(i);
230             cell.setCellValue(title[i]);
231             cell.setCellStyle(style);
232         }
233 
234         // 创建内容
235         for (int i = 0; i < list.size(); i++) {
236             row = sheet.createRow(i + 1);
237             for (int j = 0; j < list.get(i).size(); j++) {
238                 Cell celll = row.createCell(j);
239                 // 将内容按顺序赋给对应的列对象
240                 celll.setCellValue(list.get(i).get(j));
241                 celll.setCellStyle(tableStyle);
242             }
243         }
244         try {
245             /* 响应到客户端*/
246             Export.setResponseHeader(response, fileName + ".xlsx");
247             OutputStream os = response.getOutputStream();
248             workbook.write(os);
249             os.flush();
250             os.close();
251 
252         } catch (IOException e) {
253             e.printStackTrace();
254         }
255     }
256 
257     /**
258      * 发送响应流方法
259      *
260      * @param response 响应
261      * @param fileName excel文件名263      * @since 2019-05-10
264      */
265     public static void setResponseHeader(HttpServletResponse response, String fileName) {
266         try {
267             response.setCharacterEncoding("utf-8");
268             fileName = URLEncoder.encode(fileName, "UTF-8");
269             response.setHeader("content-Type", "application/x-xls");
270             response.setHeader("Content-Disposition", "inline;filename=" + fileName);
271         } catch (Exception ex) {
272             ex.printStackTrace();
273         }
274     }
275 
276 }
//复制代码
//实现导出业务

//复制代码
 1     @Override
 2     public void export(HttpServletRequest request, HttpServletResponse response, List<EventStatusVO> list, String reportType) {
 3         List<List<String>> listList = new ArrayList<>();
 4         String dataFormatStr = "yyyy-MM-dd HH:mm:ss";
 5         SimpleDateFormat dateFormat = new SimpleDateFormat(dataFormatStr);
 6         for (EventStatusVO item: list) {
 7             List<String> lists = new ArrayList<>();
 8             //事件名称
 9             lists.add(item.getName());
10             //事件类型
11             lists.add(item.getEventType());
12             //事件级别
13             lists.add(String.valueOf(item.getEventRate()));
14             //事件场所
15             lists.add(item.getEventPlace());
16             //上报日期
17             lists.add(dateFormat.format(item.getReportDate()));
18             //发生日期 起止
19             lists.add(dateFormat.format(item.getHappenDateStart()) + "至" + dateFormat.format(item.getHappenDateEnd()));
20             //事件进度
21             lists.add(item.getEventStatus());
22             listList.add(lists);
23         }
24         String fileName = "我的上报表" + dateFormat.format(new Date());
25         String sheetName = "我的上报";
26         String []title = {"事件名称", "事件类型", "事件级别", "发生场所", "上报日期", "发生时间", "事件进度"};
27 
28         final String xls = "xls";
29         final String xlsx = "xlsx";
30         final String pdf = "pdf";
31         if (xls.equals(reportType)) {
32             Export.getHSSFWorkbook(response,request, sheetName, title, listList, fileName);
33         }
34         if (xlsx.equals(reportType)) {
35             Export.getSXSSFWorkbook(response,request, sheetName, title, listList, fileName);
36         }
37         if (pdf.equals(reportType)) {
38             try {
39                 List<String> allList = Arrays.asList(title);
40                 listList.add(0, allList);
41                 Export.download(request, response, listList, fileName);
42             } catch (Exception e) {
43                 e.printStackTrace();
44             }
45         }
46     }