package com.github.winter4666.excelio.out;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.github.winter4666.excelio.common.GridHeader;


/**
 * 对poi导出Excel的过程进行了封装，提供了一系列更方便的写Excel的方法，
 * 可以方便地对单元格进行横向纵向合并，对单元格的展示格式进行控制，导出表格。
 * @author wutian
 */
public class ExcelWriter {
	
	private ExcelFormat excelFormat;
	
	private Workbook workbook;
	
	/**
	 * 当前正在写的sheet
	 */
	private Sheet currentSheet;
	
	/**
	 * 当前正在写的行
	 */
	private Row currentRow;
	
	/**
	 * 当前正在写的行号
	 */
	private int currentRowNum;
	
	/**
	 * 当前正在写的列号
	 */
	private int currentColumnNum;
	
	/**
	 * ExcelWriter的默认CellStyle，通过createCellStyle方法创建的CellStyle会默认clone该样式，也是不指定CellStyle的时候程序默认使用的样式。
	 */
	private CellStyle defaultCellStyle;
	
	/**
	 * 日期格式
	 */
	private String dateFormat;
	
	/**
	 * 需要自动调整列宽度的ColumnIndex
	 */
	private Set<Integer> autoSizeColumnIndexes;
	
	private CreationHelper creationHelper;
	
	private boolean ignoreGridHeader;
	
	private Map<CellStyle, CellStyle> dateCellStyleTemp = new HashMap<>();
	
	/**
	 * 根据内容自动调整excel列宽度
	 */
	private boolean autoSizeColumn;
	
	ExcelWriter(ExcelFormat excelFormat,Boolean autoSizeColumn,Integer rowAccessWindowSize) {
		if(excelFormat == null) excelFormat = ExcelFormat.XSSF;
		this.excelFormat = excelFormat;
		if(excelFormat == ExcelFormat.HSSF) {
			workbook = new HSSFWorkbook();
		} else if(excelFormat == ExcelFormat.SXSSF) {
			if(rowAccessWindowSize == null) {
				workbook = new SXSSFWorkbook();
			} else {
				workbook = new SXSSFWorkbook(rowAccessWindowSize);
			}
		} else {
			workbook = new XSSFWorkbook();
		}
		
		if(autoSizeColumn == null) autoSizeColumn = false;
		this.autoSizeColumn = autoSizeColumn;
		
		creationHelper = workbook.getCreationHelper();
		
		dateFormat = "yyyy-MM-dd HH:mm:ss";
		ignoreGridHeader = false;
		
		defaultCellStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontName("宋体");
		defaultCellStyle.setFont(font);
		defaultCellStyle.setWrapText(true);
	}
	
	void initSheet(String sheetName) {
		if(sheetName != null) {
			sheetName = WorkbookUtil.createSafeSheetName(sheetName);
			currentSheet = workbook.createSheet(sheetName);
		} else {
			currentSheet = workbook.createSheet();
		}
		currentRowNum = 0;
		currentRow = currentSheet.createRow(currentRowNum);
		currentColumnNum = 0;
		
		autoSizeColumnIndexes = new HashSet<>();
		if(autoSizeColumn && excelFormat == ExcelFormat.SXSSF) {
			((SXSSFSheet)currentSheet).trackAllColumnsForAutoSizing();
		}
	}
	
	/**
	 * 得到坐标为(x,y)处的cell的样式
	 * @param x
	 * @param y
	 * @return
	 */
	public CellStyle getCellStyle(int x,int y) {
		return currentSheet.getRow(y).getCell(x).getCellStyle();
	}
	
	/**
	 * 创建一个单元格样式
	 * @return
	 */
	public CellStyle createCellStyle() {
		if(defaultCellStyle != null) {
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.cloneStyleFrom(defaultCellStyle);
			return cellStyle;
		} else {
			return workbook.createCellStyle();
		}
	}
	
	/**
	 * 创建字体样式
	 * @return
	 */
	public Font createFont() {
		return workbook.createFont();
	}
	
	/**
	 * 合并单元格
	 * @param firstRow 起始行
	 * @param lastRow 结束行
	 * @param firstCol 起始列
	 * @param lastCol 结束列
	 */
	public void mergeCells(int firstRow,int lastRow,int firstCol,int lastCol) {
		currentSheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}
	
	/**
	 * 得到当前正在写的sheet
	 * @return
	 */
	public Sheet getCurrentSheet() {
		return currentSheet;
	}
	
	/**
	 * 得到creationHelper
	 * @return
	 */
	public CreationHelper getCreationHelper() {
		return creationHelper;
	}
	
	/**
	 * 设置日期格式，默认yyyy-MM-dd HH:mm:ss
	 * @param dateFormat
	 */
	public ExcelWriter setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
		return this;
	}
	
	/**
	 * 设置ExcelWriter的默认CellStyle，通过createCellStyle方法创建的CellStyle会默认clone该样式，也是不指定CellStyle的时候程序默认使用的样式。
	 * @param cellStyle
	 * @return
	 */
	public ExcelWriter cellStyle(CellStyle cellStyle) {
		defaultCellStyle = workbook.createCellStyle();
		defaultCellStyle.cloneStyleFrom(cellStyle);
		return this;
	}
	
	/**
	 * 设置当前正在写的行的高度
	 * @param height 行高度
	 */
	public ExcelWriter setCurrentRowHeight(float height) {
		currentRow.setHeightInPoints(height);
		return this; 
	}
	
	/**
	 * 设置列宽度
	 * @param columnNum 列号
	 * @param width 列宽度
	 */
	public void setColumnWidth(int columnNum,int width) {
		currentSheet.setColumnWidth(columnNum, width);
	}
	
	/**
	 * 批量设置excel的列宽度（以1/256字符为单位）
	 * @param widths
	 */
	public void setColumnsWidth(int...widths) {
		for(int i = 0;i < widths.length;i++) {
			setColumnWidth(i, widths[i]);
		}
	}
	
	/**
	 * 批量设置excel的列宽度（以字符为单位）
	 * @param widths
	 */
	public void setColumnsWidth(float...widths) {
		for(int i = 0;i < widths.length;i++) {
			setColumnWidth(i, (int)(widths[i] * 256));
		}
	}
	
	/**
	 * 设置当前正在写的列宽度
	 * @param width 列宽度
	 * @return
	 */
	public ExcelWriter setCurrentColumnWidth(int width) {
		setColumnWidth(currentColumnNum, width);
		return this;
	}
	
	/**
	 * 写表格的时候忽略表头
	 * @param ignoreGridHeader
	 */
	public ExcelWriter ignoreGridHeader(boolean ignoreGridHeader) {
		this.ignoreGridHeader = ignoreGridHeader;
		return this;
	}
	
	private CellStyle getDateCellStyle(CellStyle cellStyle) {
		CellStyle dateCellStyle = dateCellStyleTemp.get(cellStyle);
		if(dateCellStyle == null) {
			dateCellStyle = workbook.createCellStyle();
			dateCellStyle.cloneStyleFrom(cellStyle);
			dateCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(dateFormat));
			dateCellStyleTemp.put(cellStyle, dateCellStyle);
		}
		return dateCellStyle;
	}

	/**
	 * 在Excel中写一条数据，自定义样式
	 * @param data 要写的数据
	 * @param horizontalCellNum 要写的数据横向所占的单元格数
	 * @param verticalCellNum 要写的数据纵向所占的单元格数
	 * @param cellStyle 样式
	 * @return
	 */
	public ExcelWriter write(Object data,int horizontalCellNum,int verticalCellNum,CellStyle cellStyle) {
		//处理日期格式
		if(cellStyle != null && data instanceof Date && cellStyle.getDataFormat() == 0) {
			cellStyle = getDateCellStyle(cellStyle);
		}
		
		//添加横向单元格
		boolean setValue = false;
		for(int i = 0;i < horizontalCellNum;i++) {
			Cell cell = currentRow.createCell(currentColumnNum);
			
			//设置cell值
			if(!setValue) {
				if(data != null) {
					if(data instanceof String) {
						cell.setCellValue((String)data);
					} else if(data instanceof Number) {
						cell.setCellValue(Double.valueOf(data.toString()));
					} else if(data instanceof Date) {
						cell.setCellValue((Date)data);
					} else {
						cell.setCellValue(data.toString());
					}
				} else {
					cell.setBlank();
				}
				setValue = true;
			}
			cell.setCellStyle(cellStyle);
			
			currentColumnNum++;
			detectAutoSizeColumnIndexes();
		}
		
		//添加纵向单元格
		for(int i = 1;i < verticalCellNum;i++) {
			Row row = null;
			if(currentSheet.getRow(currentRowNum + i) != null) {
				row = currentSheet.getRow(currentRowNum + i);
			} else {
				row = currentSheet.createRow(currentRowNum + i);
			}
			for(int j = 0;j < horizontalCellNum;j++) {
				Cell cell = row.createCell(currentColumnNum - horizontalCellNum + j);
				cell.setCellStyle(cellStyle);
			}
		}
		
		//合并
		if(horizontalCellNum > 1 || verticalCellNum > 1) {
			currentSheet.addMergedRegion(new CellRangeAddress(currentRowNum, currentRowNum + verticalCellNum - 1, currentColumnNum - horizontalCellNum, currentColumnNum -1));
		}
		return this;
	}
	
	/**
	 * 在Excel中写一条数据，默认样式
	 * @param data 要写的数据
	 * @param horizontalCellNum 要写的数据横向所占的单元格数
	 * @param verticalCellNum 要写的数据纵向所占的单元格数
	 * @return
	 */
	public ExcelWriter write(Object data,int horizontalCellNum,int verticalCellNum) {
		write(data, horizontalCellNum,verticalCellNum, defaultCellStyle);
		return this;
	}
	
	/**
	 * 在Excel中写一条数据，自定义样式
	 * @param data 要写的数据
	 * @param cellNum 要写的数据所占的单元格数
	 * @param cellStyle 样式
	 * @return
	 */
	public ExcelWriter write(Object data,int cellNum,CellStyle cellStyle) {
		return write(data, cellNum, 1, cellStyle);
	}
	
	/**
	 * 在Excel中写一条数据，设置单元格数
	 * @param data 要写的数据
	 * @param cellNum 要写的数据所占的单元格数
	 * @return
	 */
	public ExcelWriter write(Object data,int cellNum) {
		return write(data, cellNum, 1);
	}
	
	/**
	 * 在Excel中写一条数据，默认样式
	 * @param data 要写的数据
	 * @return
	 */
	public ExcelWriter write(Object data) {
		return write(data, 1);
	}
	
	/**
	 * 在Excel中写一个表格，默认样式
	 * @param headers
	 * @param data
	 * @return
	 */
	public ExcelWriter writeGrid(List<GridHeader> headers, List<?> data) {
		return writeGrid(headers, data, new GridCellStyle() {
			
			@Override
			public CellStyle getHeaderCellStyle(ExcelWriter excelWriter, String fieldName) {
				return defaultCellStyle;
			}
			
			@Override
			public CellStyle getDataCellStyle(ExcelWriter excelWriter, String fieldName, int gridRowNum, Object fieldValue) {
				return defaultCellStyle;
			}

		});
	}
	
	
	/**
	 * 在Excel中写一个表格，自定义样式
	 * @param headers 表头
	 * @param data 数据
	 * @param gridCellStyle 表格样式
	 */
	@SuppressWarnings("unchecked")
	public ExcelWriter writeGrid(List<GridHeader> headers, List<?> data,GridCellStyle gridCellStyle) {
		int offset = currentColumnNum - 0;
		
		//写表头
		if(!ignoreGridHeader) {
			for(GridHeader header : headers) {
				write(header.getLabel(), header.getCellNum(), gridCellStyle.getHeaderCellStyle(this, header.getFieldName()));
			}
			nextLine();
		}
		if(data == null) {
			return this;
		}
		//写表格数据
		for(int i = 0;i < data.size();i++) {
			Object rowData = data.get(i);
			if(currentColumnNum == 0 && offset != 0) {
				skip(offset);
			}
			for(GridHeader header : headers) {
				Object dataCellValue;
				if(rowData instanceof Map) {
					dataCellValue = ((Map<String,Object>)rowData).get(header.getFieldName());
				} else {
					try {
						Method method = rowData.getClass().getMethod("get" 
								+ header.getFieldName().substring(0, 1).toUpperCase() + header.getFieldName().substring(1));
						dataCellValue = method.invoke(rowData);
					} catch (NoSuchMethodException | SecurityException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
						throw new RuntimeException("get value of " + header.getFieldName() + "error", e);
					}
				}
				if(header.getFieldValueConverter() != null) {
					dataCellValue = header.getFieldValueConverter().convert(dataCellValue,rowData);
				}
				write(dataCellValue, header.getCellNum(), gridCellStyle.getDataCellStyle(this, header.getFieldName(), i, dataCellValue));
			}
			if(i != data.size() - 1) {
				nextLine();
			}
		}
		return this;
	}
	
	private void detectAutoSizeColumnIndexes() {
		if(autoSizeColumn) {
			int columnIndex = currentColumnNum - 1;
			if(!autoSizeColumnIndexes.contains(columnIndex)) {
				autoSizeColumnIndexes.add(columnIndex);
			}
		}
	}
	
	private void autoSizeColumns() {
		if(autoSizeColumn) {
			for(int columnIndex : autoSizeColumnIndexes) {
				currentSheet.autoSizeColumn(columnIndex);
			}
		}
	}
	
	/**
	 * 使写Excel的光标跳过若干单元格
	 * @param cellNum 跳过的单元格数
	 * @return
	 */
	public ExcelWriter skip(int cellNum) {
		currentColumnNum = currentColumnNum + cellNum;
		return this;
	}
	
	/**
	 * 让写Excel的光标换一行
	 * @param height 下一行的高度 
	 * @return
	 */
	public ExcelWriter nextLine(Float height) {
		currentRowNum++;
		if(currentSheet.getRow(currentRowNum) == null) {
			currentRow = currentSheet.createRow(currentRowNum);
		} else {
			currentRow = currentSheet.getRow(currentRowNum);
		}
		currentColumnNum = 0;
		if(height != null) {
			setCurrentRowHeight(height);
		}
		return this;
	}
	
	/**
	 * 让写Excel的光标换一行
	 * @return
	 */
	public ExcelWriter nextLine() {
		nextLine(null);
		return this;
	}
	
	/**
	 * 切换到下一个sheet
	 */
	public ExcelWriter nextSheet() {
		nextSheet(null);
		return this;
	}
	
	/**
	 * 切换到下一个sheet
	 * @param name sheet名
	 */
	public ExcelWriter nextSheet(String name) {
		autoSizeColumns();
		initSheet(name);
		return this;
	}
	
	/**
	 * 导出excel
	 * @param outputStream
	 * @throws IOException
	 */
	public void export(OutputStream outputStream) throws IOException {
		autoSizeColumns();
		try {
			workbook.write(outputStream);
		} finally {
			workbook.close();
			if(excelFormat == ExcelFormat.SXSSF) {
				SXSSFWorkbook wb = (SXSSFWorkbook)workbook;
				wb.dispose();
			}
		}
	}
	
	/**
	 * 导出excel到文件
	 * @param file
	 * @throws IOException
	 */
	public void exportToFile(File file) throws IOException {
		FileOutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(file);
			export(outputStream);
		} finally {
			if(outputStream != null) {
				outputStream.close();
			}
		}
	}
	
	/**
	 * 导出excel到内存
	 * @return
	 * @throws IOException
	 */
	public byte[] exportToByteArray() throws IOException {
		ByteArrayOutputStream outputStream = null;
		try {
			outputStream = new ByteArrayOutputStream();
			export(outputStream);
		} finally {
			if (outputStream != null) {
				outputStream.close();
			}
		}
		return outputStream.toByteArray();
	}
	
	/**
	 * Excel格式
	 * @author wutian
	 */
	public enum ExcelFormat {
		HSSF,
		XSSF,
		SXSSF
	}
	
	/**
	 * 表格样式
	 * @author wutian
	 */
	public interface GridCellStyle {
		
		/**
		 * 得到表头样式
		 * @param excelWriter excelWriter
		 * @param fieldName 字段名称
		 * @return
		 */
		CellStyle getHeaderCellStyle(ExcelWriter excelWriter,String fieldName);
		
		/**
		 * 得到表格数据样式
		 * @param excelWriter excelWriter
		 * @param fieldName 字段名称
		 * @param currentRowNum 所在表格行数，从0开始
		 * @param fieldValue 字段值
		 * @return
		 */
		CellStyle getDataCellStyle(ExcelWriter excelWriter,String fieldName,int gridRowNum,Object fieldValue);
	}

}
