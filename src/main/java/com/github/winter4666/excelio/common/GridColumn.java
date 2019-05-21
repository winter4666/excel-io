package com.github.winter4666.excelio.common;

/**
 * 表格列
 * @author wutian
 */
public class GridColumn {
	
	/**
	 * 字段名
	 */
	private String fieldName;
	
	/**
	 * 字段在表头显示的名称
	 */
	private String label;
	
	/**
	 *  转换字段的值
	 */
	private FieldValueConverter fieldValueConverter;
	
	/**
	 * 所占单元格的格数
	 */
	private int cellNum;
	
	private GridColumn() {
		
	}
	
	public static GridColumn newInstance(String fieldName) {
		GridColumn gridColumn = new GridColumn();
		gridColumn.fieldName = fieldName;
		gridColumn.cellNum = 1;
		return gridColumn;
	}
	
	public static GridColumn newInstance(String fieldName,String label) {
		GridColumn gridColumn = new GridColumn();
		gridColumn.fieldName = fieldName;
		gridColumn.label = label;
		gridColumn.cellNum = 1;
		return gridColumn;
	}
	
	public GridColumn label(String label) {
		this.label = label;
		return this;
	}
	
	public GridColumn cellNum(int cellNum) {
		this.cellNum = cellNum;
		return this;
	}
	
	public GridColumn fieldValueConverter(FieldValueConverter fieldValueConverter) {
		this.fieldValueConverter = fieldValueConverter;
		return this;
	}
	
	public String getFieldName() {
		return fieldName;
	}

	public String getLabel() {
		return label;
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
		 * 传入字段原来的值，返回处理过的值作为写到excel的值
		 * @param fieldValue 字段初始值
		 * @param rowData 整个数据对象
		 * @return 处理过
		 */
		Object convert(Object fieldValue,Object rowData);
	}

}
