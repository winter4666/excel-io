package com.github.winter4666.excelio.common;

import java.util.HashMap;
import java.util.Map;

/**
 * 表格头
 * @author wutian
 */
public class GridHeader {
	
	/**
	 * 字段名
	 */
	private String fieldName;
	
	/**
	 * 字段在表头显示的名称
	 */
	private String label;
	
	/**
	 * 字典
	 */
	private Map<Object, String> dictionary;
	
	/**
	 *  转换字段的值
	 */
	private FieldValueConverter fieldValueConverter;
	
	/**
	 * 所占单元格的格数
	 */
	private int cellNum;
	
	private GridHeader() {
		
	}
	
	public static GridHeader newInstance(String fieldName,String label) {
		GridHeader gridHeader = new GridHeader();
		gridHeader.fieldName = fieldName;
		gridHeader.label = label;
		gridHeader.cellNum = 1;
		return gridHeader;
	}
	
	public GridHeader cellNum(int cellNum) {
		this.cellNum = cellNum;
		return this;
	}
	
	public GridHeader dictionary(Map<Object, String> dictionary) {
		this.dictionary = dictionary;
		return this;
	}
	
	public GridHeader entry(Object key,String value) {
		if(dictionary == null) {
			dictionary = new HashMap<>();
		}
		dictionary.put(key, value);
		return this;
	}
	
	public GridHeader fieldValueConverter(FieldValueConverter fieldValueConverter) {
		this.fieldValueConverter = fieldValueConverter;
		return this;
	}
	
	public String getFieldName() {
		return fieldName;
	}

	public String getLabel() {
		return label;
	}

	public Map<Object, String> getDictionary() {
		return dictionary;
	}

	public int getCellNum() {
		return cellNum;
	}
	
	public FieldValueConverter getFieldValueConverter() {
		return fieldValueConverter;
	}

	/**
	 * 转换字段的值
	 * @author wutian
	 * @param <T>
	 */
	public interface FieldValueConverter {
		
		/**
		 * 传入字段原来的值，返回转换后的值作为字段新的值
		 * @param fieldValue
		 * @return
		 */
		Object convert(Object fieldValue);
	}

}
