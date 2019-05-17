package com.github.winter4666.excelio.out;

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
	 * 设置日期格式，默认yyyy-MM-dd HH:mm:ss
	 * @param dateFormat
	 */
	public ExcelWriterBuilder setDateFormat(String dateFormat) {
		excelWriter.dateFormat = dateFormat;
		return this;
	}
	
	/**
	 * 根据内容自动调整excel列宽度，默认关闭
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
