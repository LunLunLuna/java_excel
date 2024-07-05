package com.zenithst.common.util.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;

public interface ExcelRowCreateHandler<T> {	
	//void excelRowCreate(Object obj,SXSSFSheet sheet,int rowno);
	void excelRowCreate(T t,SXSSFSheet sheet,int rowno);
}
