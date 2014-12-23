package com.fis.report.build;

import com.smartxls.ChartShape;
import com.smartxls.WorkBook;

/**
 * 
 * @author THINHNV
 * 
 */
public class BuildChartReport {
	/**
	 * 
	 * @param templateUrl
	 * @param numOfSheet
	 * @param locationRowData
	 * @param locationColData
	 * @throws Exception
	 */
	public void createLineChart(String templateUrl, int numOfSheet, int locationRowData, int locationColData) throws Exception {
		WorkBook workBook = new WorkBook();
		workBook.read(getClass().getResourceAsStream(templateUrl));
		workBook.setSheet(numOfSheet);
		ChartShape chart = workBook.getChart(0);
		chart.addSeries();
	}
}
