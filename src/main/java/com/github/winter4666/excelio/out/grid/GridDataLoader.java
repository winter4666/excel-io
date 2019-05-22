package com.github.winter4666.excelio.out.grid;

/**
 * excel导出表格时加载数据的接口
 * @author wutian
 */
public interface GridDataLoader {
	
	/**
	 * 加载数据
	 */
	void loadData();
	
	/**
	 * 获取每一行数据
	 * @param listener
	 */
	void getRowData(GridDataLoaderListener listener);
	
	interface GridDataLoaderListener {
		void onReadRowData(int gridRowNum,Object rowData);
	}
	
}
