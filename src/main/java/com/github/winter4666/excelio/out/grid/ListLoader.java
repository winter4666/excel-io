package com.github.winter4666.excelio.out.grid;

import java.util.List;

/**
 * 列表数据加载器
 * @author wutian
 */
public class ListLoader implements GridDataLoader{
	
	private List<?> list;
	
	public ListLoader(List<?> list) {
		this.list = list;
	}

	@Override
	public void loadData() {
		
	}

	@Override
	public void getRowData(GridDataLoaderListener listener) {
		for(int i = 0;i < list.size();i++) {
			Object rowData = list.get(i);
			listener.onReadRowData(i, rowData);
		}
	}

}
