package com.github.winter4666.excelio.out.grid;

import java.util.List;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 异步分页加载器
 * @author wutian
 */
public class AsyncPagingLoader implements GridDataLoader {
	
	private final static Logger logger = LoggerFactory.getLogger(AsyncPagingLoader.class);
	
	private BlockingQueue<Object> gridRowDataQuene;
	
	private int pageSize;
	
	private GridDataSource<?> gridDataSource;
	
	private static final int OFFER_TIMEOUT = 1;
	
	private static final int POLL_TIMEOUT = 5;
	
	private Thread thread;
	
	/**
	 * 构造异步分页加载器，默认一次加载1000条数据
	 * @param gridDataSource 数据源
	 */
	public AsyncPagingLoader(GridDataSource<?> gridDataSource) {
		this(gridDataSource,1000);
	}
	
	/**
	 * 构造异步分页加载器
	 * @param gridDataSource 数据源
	 * @param pageSize 一次加载数据的条数
	 */
	public AsyncPagingLoader(GridDataSource<?> gridDataSource,int pageSize) {
		this.gridDataSource = gridDataSource;
		this.pageSize = pageSize;
	}
	
	@Override
	public void loadData() {
		gridRowDataQuene = new LinkedBlockingQueue<>(pageSize*2);
		thread = new Thread(new Runnable() {
			
			@Override
			public void run() {
				int pageNo = 1;
				try {
					while(true) {
						List<?> rowDataList = gridDataSource.getGridData(pageNo, pageSize);
						if(rowDataList != null) {
							for(Object rowData : rowDataList) {
								if(!gridRowDataQuene.offer(rowData, OFFER_TIMEOUT, TimeUnit.MINUTES)) {
									throw new RuntimeException("offer timeout");
								}
							}
						}
						if(rowDataList == null || rowDataList.size() < pageSize) {
							if(!gridRowDataQuene.offer(new PoisonPill(),OFFER_TIMEOUT, TimeUnit.MINUTES)) {
								throw new RuntimeException("offer timeout");
							}
							break;
						} else {
							pageNo++;
						}
					}
				} catch (Throwable t) {
					logger.error("error occours while get grid data from dataSource,pageNo=" + pageNo + ",pageSize=" + pageSize,t);
					try {
						gridRowDataQuene.clear();
						gridRowDataQuene.put(new PoisonPill(t));
					} catch (InterruptedException e) {
						logger.error(e.getMessage(),e);
					}
				}
				
			}
		});
		thread.start();
	}
	
	@Override
	public void getRowData(GridDataLoaderListener listener) {
		int i = 0;
		while(true) {
			try {
				Object rowData = gridRowDataQuene.poll(POLL_TIMEOUT, TimeUnit.MINUTES);
				if(rowData == null) {
					thread.interrupt();
					throw new RuntimeException("poll timeout");
				} else if(rowData instanceof PoisonPill) {
					PoisonPill poisonPill = (PoisonPill)rowData;
					if(poisonPill.getT() != null) {
						throw new RuntimeException("error occours while get grid data from dataSource", poisonPill.getT());
					} else {
						break;
					}
				} else {
					listener.onReadRowData(i, rowData);
					i++;
				}
			} catch (InterruptedException e) {
				logger.error(e.getMessage(),e);
			}
		}
	}
	
	/**
	 * 致命药丸
	 * @author wutian
	 */
	private static class PoisonPill {
		
		private Throwable t;
		
		public PoisonPill() {
			
		}

		public PoisonPill(Throwable t) {
			super();
			this.t = t;
		}

		public Throwable getT() {
			return t;
		}
		
	}
	
	/**
	 * 表格数据源
	 * @param <T>
	 */
	public interface GridDataSource<T> {
		
		/**
		 * 获取表格数据
		 * @param pageNo 页码，从1开始
		 * @param pageSize 一页里面的记录数
		 * @return
		 */
		List<T> getGridData(int pageNo,int pageSize) throws InterruptedException;
		
	}

}
