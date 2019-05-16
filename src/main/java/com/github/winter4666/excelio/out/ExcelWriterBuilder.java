package com.github.winter4666.excelio.out;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import com.github.winter4666.excelio.out.ExcelWriter.ExcelFormat;

/**
 * 负责构造{@link ExcelWriter}
 * @author wutian
 */
public class ExcelWriterBuilder {
	
	private ExcelWriter excelWriter;
	
	public ExcelWriterBuilder(ExcelFormat excelFormat,String sheetName) {
		excelWriter = new ExcelWriter(excelFormat, sheetName);
	}
	
	public ExcelWriterBuilder() {
		this(ExcelFormat.XSSF, null);
	}
	
	/**
	 * 设置字体名，默认为宋体
	 * @param fontName
	 * @return
	 */
	public ExcelWriterBuilder setFontName(String fontName) {
		excelWriter.fontName = fontName;
		return this;
	}
	
	/**
	 * 设置字体大小
	 * @param fontSize
	 */
	public ExcelWriterBuilder setFontSize(short fontSize) {
		excelWriter.fontSize = fontSize;
		return this;
	}
	
	/**
	 * 设置日期格式，默认yyyy-MM-dd HH:mm:ss
	 * @param dateFormat
	 */
	public ExcelWriterBuilder setDateFormat(String dateFormat) {
		excelWriter.dateFormat = dateFormat;
		return this;
	}
	
	/**
	 * 设置单元格竖直对齐方式
	 * @param horizontalAlignment
	 */
	public ExcelWriterBuilder setCellHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
		excelWriter.cellHorizontalAlignment = horizontalAlignment;
		return this;
	}
	
	/**
	 * 设置单元格水平对齐方式
	 * @param horizontalAlignment
	 */
	public ExcelWriterBuilder setCellVerticalAlignment(VerticalAlignment verticalAlignment) {
		excelWriter.cellVerticalAlignment = verticalAlignment;
		return this;
	}
	
	/**
	 * 设置单元格边界
	 * @param borderStyle
	 * @return
	 */
	public ExcelWriterBuilder setCellBorder(BorderStyle borderStyle) {
		excelWriter.cellBorder = borderStyle;
		return this;
	}
	
	/**
	 * 设置单元格边界颜色
	 */
	public ExcelWriterBuilder setCellBorderColor(IndexedColors color) {
		excelWriter.cellBorderColor = color;
		return this;
	}
	
	/**
	 * 根据内容自动调整excel列宽度
	 * @return
	 */
	public ExcelWriterBuilder enableAutoSizeColumn() {
		excelWriter.autoSizeColumn = true;
		return this;
	}
	
	public ExcelWriter build() {
		return excelWriter;
	}
}
