package com.github.winter4666.excelio.out;

import com.github.winter4666.excelio.out.ExcelWriter.ExcelFormat;

/**
 * 负责构造{@link ExcelWriter}
 * @author wutian
 */
public class ExcelWriterBuilder {
	
	private ExcelFormat excelFormat;
	
	private String sheetName;
	
	private Boolean autoSizeColumn;
	
	private Integer rowAccessWindowSize;
	
	public ExcelWriterBuilder() {
		
	}
	
	/**
	 * 设置excel格式，默认XSSF
	 * @param excelFormat
	 * @return
	 */
	public ExcelWriterBuilder setExcelFormat(ExcelFormat excelFormat) {
		this.excelFormat = excelFormat;
		return this;
	}
	
	/**
	 * 生成SXSSF格式的excel时，可以设置该参数，表示内存中最多保存的行数量
	 * @return
	 * @see org.apache.poi.xssf.streaming.SXSSFWorkbook#SXSSFWorkbook(int)
	 */
	public ExcelWriterBuilder setRowAccessWindowSize(int rowAccessWindowSize) {
		this.rowAccessWindowSize = rowAccessWindowSize;
		return this;
	}
	
	/**
	 * 设置sheetName
	 * @param sheetName
	 * @return
	 * @see org.apache.poi.ss.usermodel.Workbook#createSheet(String)
	 */
	public ExcelWriterBuilder setSheetName(String sheetName) {
		this.sheetName = sheetName;
		return this;
	}

	/**
	 * 根据内容自动调整excel列宽度，默认关闭
	 * @return
	 * @see org.apache.poi.ss.usermodel.Sheet#autoSizeColumn(int)
	 */
	public ExcelWriterBuilder enableAutoSizeColumn() {
		autoSizeColumn = true;
		return this;
	}
	
	public ExcelWriter build() {
		ExcelWriter excelWriter = new ExcelWriter(excelFormat, autoSizeColumn,rowAccessWindowSize);
		excelWriter.initSheet(sheetName);
		return excelWriter;
	}
}
