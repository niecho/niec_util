package com.niec.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


@SuppressWarnings("deprecation")
public class ExportExcel {
	
	
	/** 每页数据数目 */
	private static final int SHEET_DATA_SIZE = 50000;
	/** 数字正则规则 */
	private static final Pattern NUMBER_PATTERN = Pattern.compile("^[-+]?\\d+(\\.\\d+)?$");
	
	
	
	/**
	 * @see #exportExcel(String, List, String[])
	 *
	 * @param dataList 		数据集合
	 */
	public static <T> InputStream exportExcel(List<T> dataList) {
		return exportExcel(null, dataList, null);
	}
	
	
	/**
	 * @see #exportExcel(String, List, String[])
	 *
	 * @param dataList 		数据集合
	 * @param headerNames 	标题列名
	 */
	public static <T> InputStream exportExcel(List<T> dataList, String[] headerNames) {
		return exportExcel(null, dataList, headerNames);
	}
	
	
	/**
	 * 将集合中的数据写入到输入流中<br/>
	 * 该方法较消耗系统资源,建议数据量较少时使用
	 *
	 * @param sheetName   	sheet名
	 * @param dataList 		数据集合
	 * @param headerNames 	标题列名
	 */
	public static <T> InputStream exportExcel(String sheetName, List<T> dataList, String[] headerNames) {
		// 声明一个工作薄
		SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
		// 生成表头样式
		CellStyle headerStyle = createHeaderStyle(workbook);
		// 生成文字样式
		CellStyle normalStyle = createTextStyle(workbook, false);
		// 生成数字样式
		CellStyle numberStyle = createTextStyle(workbook, true);
		// 数字显示字体
		Font numberFont = workbook.createFont();
		numberFont.setColor(IndexedColors.BLUE.index);
		// 计算数据页数
		int dataPage = dataList.size() / SHEET_DATA_SIZE + (dataList.size() % SHEET_DATA_SIZE > 0 ? 1 : 0);
		// 写入数据
		for (short i = 0; i < dataPage; i++) {
			List<T> pageList = dataList.subList(i * SHEET_DATA_SIZE, i == dataPage - 1 ? dataList.size() : (i + 1) * SHEET_DATA_SIZE);
			createSheet(workbook, sheetName == null ? "第" + i + "页" : sheetName + "_" + i, pageList, headerNames, headerStyle, normalStyle, numberStyle);
		}
		InputStream is = null;
		ByteArrayOutputStream out = null;
		try {
			out = new ByteArrayOutputStream();
			workbook.write(out);
			is = new ByteArrayInputStream(out.toByteArray());
			//刷新缓冲区
			out.flush();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (out != null)
					out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return is;
	}
	
	
	/**
	 * @see #exportExcel(String, List, String[], String)
	 *
	 * @param dataList 		数据集合,集合中元素(T--JavaBean)可以存放byte[]数据(图片)
	 * @param outputPath 	文件输出路径
	 */
	public static<T> void exportExcel(List<T> dataList, String outputPath) {
		exportExcel(null, dataList, null, outputPath);
	}
	
	
	/**
	 * @see #exportExcel(String, List, String[], String)
	 *
	 * @param dataList 		数据集合,集合中元素(T--JavaBean)可以存放byte[]数据(图片)
	 * @param headerNames 	标题行列名
	 * @param outputPath 	文件输出路径
	 */
	public static<T> void exportExcel(List<T> dataList, String [] headerNames, String outputPath) {
		exportExcel(null, dataList, headerNames, outputPath);
	}
	
	
	/**
	 * 将集合中的数据导出到指定的Excel文件中(格式.xlsx)
	 *
	 * @param sheetName   	Sheet标题
	 * @param dataList 		数据集合,集合中元素(T)可以存放byte[]数据(图片)
	 * @param headerNames 	标题行列明
	 * @param outputPath 	文件输出路径
	 */
	public static<T> void exportExcel(String sheetName, List<T> dataList, String [] headerNames, String outputPath) {
		// 声明一个工作薄
		SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
		// 生成首行样式
		CellStyle headerStyle = createHeaderStyle(workbook);
		// 生成文字样式
		CellStyle normalStyle = createTextStyle(workbook, false);
		// 生成数字样式
		CellStyle numberStyle = createTextStyle(workbook, true);
		// 计算数据页数
		int dataPage = dataList.size() / SHEET_DATA_SIZE + (dataList.size() % SHEET_DATA_SIZE > 0 ? 1 : 0);
		// 写入数据
		for (int i = 0; i < dataPage; i++) {
			List<T> pageList = dataList.subList(i * SHEET_DATA_SIZE, i == dataPage - 1 ? dataList.size() : (i + 1) * SHEET_DATA_SIZE);
			createSheet(workbook, sheetName == null ? "第" + (i + 1) + "页" : sheetName + "_" + i, pageList, headerNames, headerStyle, normalStyle, numberStyle);
		}
		FileOutputStream fOut = null;
		try {
			// 生成文件
			fOut = new FileOutputStream(outputPath);
			workbook.write(fOut);
			JOptionPane.showMessageDialog(null, "导出成功!");
			//刷新缓冲区
			fOut.flush();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if(fOut != null)
					fOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	
	/**
	 * 利用JAVA的反射,将集合中的数据以EXCEL的形式输出
	 *
	 * @param workbook   	工作簿
	 * @param sheetName   	Sheet标题
	 * @param dataList 		数据集合,集合中元素(T)可以存放byte[]数据(图片)
	 * @param headerNames 	标题行列名
	 * @param headerStyle 	首行样式
	 * @param normalStyle 	文字样式
	 * @param numberStyle 	数字样式
	 */
	@SuppressWarnings("unchecked")
	private static<T> void createSheet(SXSSFWorkbook workbook, 
									   String sheetName,
									   List<T> dataList, 
									   String [] headerNames, 
									   CellStyle headerStyle, 
									   CellStyle normalStyle,
									   CellStyle numberStyle) {
		// 初始化表格元素
		SXSSFSheet sheet = workbook.createSheet(sheetName);
		Row row = sheet.createRow(0);
		// 初始化表格列宽
		int [] columnMaxWidth = new int[headerNames.length];
		for (short i = 0; i < headerNames.length; i++)
			columnMaxWidth[i] = getTextLength(headerNames[i]);
		// 标题行
		for (short i = 0; i < headerNames.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellStyle(headerStyle);
			cell.setCellValue(new XSSFRichTextString(headerNames[i]));
		}
		// 反射取得元素的get/is方法
		String [] methodNames = getFields(dataList.get(0));
		// 生成一个绘图管理器
		Drawing drawing = sheet.createDrawingPatriarch();
		// 利用反射,动态调用getXxx()/isXxx()方法得到属性值,然后写入表格
		for (int index = 0; index < dataList.size(); index++) {
			row = sheet.createRow(index + 1);
			T t = dataList.get(index);
			for (short column = 0; column < methodNames.length; column++) {
				Cell cell = row.createCell(column);
				try {
					Object value = t.getClass().getMethod(methodNames[column]).invoke(t);
					if (value instanceof Boolean) {
						// 布尔类型
						if ((Boolean) value) {
							cell.setCellValue("是");
						} else {
							cell.setCellValue("否");
						}
					} else if (value instanceof byte[]) {
						// 图片
						sheet.setDefaultRowHeightInPoints(60);
						sheet.setColumnWidth(column, (short) (35.7 * 80));
						XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0,
								1023, 255, (short) 6, index, (short) 6, index);
						drawing.createPicture(anchor, workbook.addPicture((byte[]) value, XSSFWorkbook.PICTURE_TYPE_JPEG));
					} else {
						// 其它类型按照字符串简单处理
						String textValue = value.toString();
						if (textValue != null) {
							if (columnMaxWidth[column] < textValue.length()) {
								columnMaxWidth[column] = textValue.length();
							}
							// 利用正则表达式判断textValue是否全部由数字组成
							Matcher matcher = NUMBER_PATTERN.matcher(textValue);
							if (matcher.matches()) {
								cell.setCellValue(Double.valueOf(matcher.group()));
								cell.setCellStyle(numberStyle);
							} else {
								cell.setCellValue(new XSSFRichTextString(textValue));
								cell.setCellStyle(normalStyle);
							}
						}
					}
				} catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
					e.printStackTrace();
				}
			}
		}
		// 设置表格最终列宽
		for(short column = 0;column < headerNames.length;column++)
			sheet.setColumnWidth(column, 256 * columnMaxWidth[column]);
	}
	
	
	/**
	 * 通过反射取得JavaBean中的所有get/is方法
	 *
	 * @param t 	类实例
	 * @param <T> 	泛型类
	 * @return String[]
	 */
	private static <T> String [] getFields(T t) {
		Field [] fields = t.getClass().getDeclaredFields();
		String [] methodNames = new String[fields.length];
		for (short i = 0; i < fields.length; i++) {
			Field field = fields[i];
			String fieldName = field.getName();
			Class fieldType = field.getType();
			if (fieldType == boolean.class)
				methodNames[i] = "is" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
			else
				methodNames[i] = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
		}
		return methodNames;
	}
	
	
	/**
	 * 生成表头样式
	 *
	 * @param workbook 工作簿
	 * @return CellStyle
	 */
	private static CellStyle createHeaderStyle(SXSSFWorkbook workbook) {
		// 设置样式
		XSSFCellStyle headerStyle = (XSSFCellStyle)workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
		headerStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		headerStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		// 设置字体
		Font titleFont = workbook.createFont();
		titleFont.setColor(IndexedColors.VIOLET.index);
		titleFont.setFontHeightInPoints((short) 12);
		titleFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		headerStyle.setFont(titleFont);
		return headerStyle;
	}
	
	
	/**
	 * 生成正文样式
	 *
	 * @param workbook 工作簿
	 * @param isNumber 是否数字
	 * @return CellStyle
	 */
	private static CellStyle createTextStyle(SXSSFWorkbook workbook, boolean isNumber) {
		// 设置样式
		XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.createCellStyle();
		cellStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		// 设置字体
		Font font = workbook.createFont();
		if(isNumber) {
			font.setColor(IndexedColors.BLUE.index);
			cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		} else {
			cellStyle.setAlignment(HorizontalAlignment.CENTER);
		}
		font.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);
		cellStyle.setFont(font);
		return cellStyle;
	}
	
	
	/**
	 * 定义字符串所占用的长度
	 *
	 * @param text 给定字符串
	 * @return int
	 */
	private static int getTextLength(String text) {
		int result = 0;
		for(char c : text.toCharArray()) {
			Character.UnicodeScript sc = Character.UnicodeScript.of(c);
			if (sc == Character.UnicodeScript.HAN)
				result += 2;
			else
				result++;
		}
		// 加2给两边留出宽度
		return result + 2;
	}
	
	
	/**
	 * 增加汇总行
	 *
	 * @param sheet			表格
	 * @param dataNumber	数据数目
	 * @param normalStyle	表格样式
	 */
	private static void addSummaryRow(SXSSFSheet sheet, int dataNumber, CellStyle normalStyle) {
		if(dataNumber == 0)
			return;
		// 空一行
		Row row = sheet.createRow(dataNumber + 2);
		Cell cell = row.createCell(0);
		cell.setCellStyle(normalStyle);
		cell.setCellValue("总件数：" + dataNumber);
		cell = row.createCell(1);
		cell.setCellStyle(normalStyle);
		// 空一行
		sheet.addMergedRegion(new CellRangeAddress(dataNumber + 2 , dataNumber + 2, 0, 1));
	}
	
	
	/**
	 * 生成注释
	 *
	 * @param drawing			绘图管理器
	 * @param col1				左上表格列
	 * @param row1				左上表格行
	 * @param col2				右下表格列
	 * @param row2				右下表格行
	 * @param commentText		注释内容
	 * @param commentAuthor		注释作者
	 */
	private static void addComment(Drawing drawing,
								   int col1,
								   int row1,
								   int col2,
								   int row2,
								   String commentText,
								   String commentAuthor) {
		// 定义注释的大小和位置,详见文档
		Comment comment = drawing.createCellComment(new XSSFClientAnchor(0,
						0, 0, 0, col1, row1, col2, row2));
		// 设置注释内容
		comment.setString(new XSSFRichTextString(commentText));
		// 设置注释作者,当鼠标移动到单元格上是可以在状态栏中看到该内容.
		comment.setAuthor(commentAuthor);
	}
	
}