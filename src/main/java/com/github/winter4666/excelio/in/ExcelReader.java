package com.github.winter4666.excelio.in;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.github.winter4666.excelio.common.GridColumn;

/**
 * 封装使用poi读取excel的过程
 * @author wutian
 */
public class ExcelReader {
	
	private Workbook workbook;
	
	/**
	 * 当前正在读的sheet
	 */
	private Sheet currentSheet;
	
	/**
	 * 当前正在读的行号，从0开始
	 */
	private int currentRownum;
	
	/**
	 * 当前正在读的列号，从0开始
	 */
	private int currentColumn;
	
	public ExcelReader(File file) {
		try {
			workbook = WorkbookFactory.create(file);
			init();
		} catch (EncryptedDocumentException | IOException e) {
			throw new RuntimeException(e);
		}
	}
	
	public ExcelReader(InputStream inputStream) {
		try {
			workbook = WorkbookFactory.create(inputStream);
			init();
		} catch (EncryptedDocumentException | IOException e) {
			throw new RuntimeException(e);
		}
	}
	
	private void init() {
		currentSheet = workbook.getSheetAt(workbook.getActiveSheetIndex()); 
		currentRownum = 0;
		currentColumn = 0;
	}
	
	/**
	 * 让读Excel的光标换一行
	 * @return
	 */
	public ExcelReader nextLine() {
		currentRownum++;
		currentColumn = 0;
		return this;
	}
	
	/**
	 * 切换到下一个sheet
	 */
	public ExcelReader selectSheet(String sheetName) {
		currentSheet = workbook.getSheet(sheetName);
		currentRownum = 0;
		currentColumn = 0;
		return this;
	}
	
	/**
	 * 读取当前单元格数据
	 * @return
	 */
	public Object read() {
		Object value = null;
		Cell cell = currentSheet.getRow(currentRownum).getCell(currentColumn);
        switch (cell.getCellType()) {
        case STRING:
        	value = cell.getRichStringCellValue().getString();
        	break;
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
            	value = cell.getDateCellValue();
            } else {
            	value = cell.getNumericCellValue();
            }
            break;
        case BOOLEAN:
        	value = cell.getBooleanCellValue();
        	break;
        case FORMULA:
        	value = cell.getCellFormula();
        	break;
        case BLANK:
        	break;
        default:
        	break;
        }
        currentColumn++;
        return value;
	}
	
	/**
	 * 读取表格数据
	 * @param gridColumns
	 * @return
	 */
	public List<Map<String,Object>> readGrid(List<GridColumn> gridColumns) {
		//读表格头
		List<String> headers = new ArrayList<String>();
		for(int i = 1;i <= currentSheet.getRow(currentRownum).getLastCellNum();i++) {
			headers.add((String)read());
		}
		if(headers.size() <= 0) {
			throw new RuntimeException("grid header not found");
		}
		nextLine();
		//读表格数据
		List<Map<String,Object>> gridData = new ArrayList<>();
		if(currentSheet.getLastRowNum() - currentRownum < 0) {
			throw new RuntimeException("no data in grid");
		}
		try {
			for(int i = currentRownum;i <= currentSheet.getLastRowNum();i++) {
				Map<String,Object> rowData = new HashMap<>();
				for(String header : headers) {
					Object cellValue = read();
					GridColumn gridColumn = null; 
					for(GridColumn column : gridColumns) {
						if(column.getLabel().equals(header)) {
							gridColumn = column;
						}
					}
					if(gridColumn == null) {
						continue;
					}
					
					rowData.put(gridColumn.getFieldName(), cellValue);
				}
				gridData.add(rowData);
				nextLine();
			}
		} catch ( IllegalArgumentException | SecurityException e) {
			throw new RuntimeException(e);
		}
		return gridData;
	}
	
	/**
	 * 读取表格数据
	 * @param <T>
	 * @param gridColumns
	 * @param clazz
	 * @return
	 */
	public <T> List<T> readGrid(List<GridColumn> gridColumns, Class<T> clazz) {
		//读表格头
		List<String> headers = new ArrayList<String>();
		for(int i = 1;i <= currentSheet.getRow(currentRownum).getLastCellNum();i++) {
			headers.add((String)read());
		}
		if(headers.size() <= 0) {
			throw new RuntimeException("grid header not found");
		}
		nextLine();
		//读表格数据
		List<T> gridData = new ArrayList<T>();
		if(currentSheet.getLastRowNum() - currentRownum < 0) {
			throw new RuntimeException("no data in grid");
		}
		try {
			for(int i = currentRownum;i <= currentSheet.getLastRowNum();i++) {
				T rowData = clazz.getDeclaredConstructor().newInstance();
				for(String header : headers) {
					Object cellValue = read();
					GridColumn gridColumn = null; 
					for(GridColumn column : gridColumns) {
						if(column.getLabel().equals(header)) {
							gridColumn = column;
						}
					}
					if(gridColumn == null) {
						continue;
					}
					
					Field field = clazz.getDeclaredField(gridColumn.getFieldName());
					field.setAccessible(true);
					if(cellValue != null) {
						if("Long".equals(field.getType().getSimpleName())) {
							long newCellValue = (long)(Double.valueOf(cellValue.toString()).doubleValue());
							field.set(rowData, newCellValue);
						} else if("Double".equals(field.getType().getSimpleName())) {
							Double newCellValue = Double.valueOf(cellValue.toString());
							field.set(rowData, newCellValue);
						} else {
							field.set(rowData, cellValue);
						}
					} else {
						field.set(rowData, cellValue);
					}
				}
				gridData.add(rowData);
				nextLine();
			}
		} catch (InstantiationException | IllegalAccessException | IllegalArgumentException | InvocationTargetException
				| NoSuchMethodException | SecurityException | NoSuchFieldException e) {
			throw new RuntimeException(e);
		}
		return gridData;
	}
	
	public void close() {
		try {
			workbook.close();
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

}
