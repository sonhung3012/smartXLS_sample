package com.fis.report;

import java.awt.Color;
import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Vector;

import com.fis.report.common.SmartUtil;
import com.fis.report.common.Util;
import com.fis.report.data.ReportConfig;
import com.fss.sql.Database;
import com.smartxls.ChartFormat;
import com.smartxls.ChartShape;
import com.smartxls.PictureShape;
import com.smartxls.RangeStyle;
import com.smartxls.ShapeFormat;
import com.smartxls.WorkBook;

/**
 * 
 * @author THINHNV
 * 
 */
public class Pms3GWeeklyReport {

	protected String mstrSysDate;
	protected String mstrUserName;
	protected String strDate;
	protected String strWeek;
	protected Vector<Vector<String>> vtDataPerHour = new Vector<Vector<String>>();
	protected Vector<Vector<String>> vtDataStatic = new Vector<Vector<String>>();
	protected Vector<Vector<String>> vtDataWeek = new Vector<Vector<String>>();
	protected Vector<Vector<String>> vtDataArea = new Vector<Vector<String>>();
	protected Vector<Vector<String>> vtDataAreaCells = new Vector<Vector<String>>();
	protected Vector<Vector<String>> vtDataAreaNodeb = new Vector<Vector<String>>();
	protected Map<String, Vector<String>> kpiMap = new HashMap<String, Vector<String>>();

	private int colEnd;
	private int rowStartSecondTable;
	private int rowStartThirdTable;
	private String[] arrCriteria1 = new String[] { "Number of Cells", "Target", "KPI average", "Critical : (<=90%)", "Major   : (90-95%)", "Minor   : (95-98%)", "Normal : (98-100%)" };
	private String[] arrCriteria2 = new String[] { "Number of Cells", "Target", "KPI average", "Critical : (>=15%)", "Major   : (10-15%)", "Minor   : (5-10%)", "Normal : (0-5%)" };
	private String[] arrCriteria3 = new String[] { "Number of Cells", "KPI average" };
	private String[] arrCriteria4 = new String[] { "Number of Node B", "Target", "KPI average", "Critical : (>=100%)", "Major   : (85-100%)", "Minor   : (30-85%)", "Normal : (0-30%)" };

	/**
	 * 
	 * @param mcnMain
	 * @param fileName
	 * @param strWeek
	 * @param strDate
	 * @throws Exception
	 */
	public void toExcel(Connection mcnMain, String strWeek, String strDate) throws Exception {

		this.strWeek = strWeek;
		this.strDate = strDate;
		// String strTempfile = "/com/fis/report/Template/lineChart.xls";
		WorkBook workBook = new WorkBook();
		// workBook.read(this.getClass().getResourceAsStream(strTempfile));
		try {
			initHeaderXls(workBook);
			// Get Data from DB
			pmsWeeklyReport(mcnMain, strWeek, strDate);
			initSheet(workBook);
			createSheet(workBook, Util.QOS_KPIS_TYPE, Util.QOS_KPIS_NUM);
			createSheet(workBook, Util.TRAFFIC_TYPE, Util.TRAFFIC_NUM);
			createSheet(workBook, Util.UTILIZATION_TYPE, Util.UTILIZATION_NUM);
		} finally {
			FileOutputStream fileOut = new FileOutputStream("Output/pms3gweeklyreport.xls");
			workBook.write(fileOut);
			Desktop.getDesktop().open(new File("Output/pms3gweeklyreport.xls"));
		}

	}

	private void initHeaderXls(WorkBook workBook) throws Exception {

		PictureShape pictureShape = workBook.addPicture(1, 0, 2.5, 6, "images/LT_icon.jpg");
		ShapeFormat shapeFormat = pictureShape.getFormat();
		shapeFormat.setPlacementStyle(ShapeFormat.PlacementFreeFloating);
		pictureShape.setFormat();

		workBook.setColWidth(1, 10000);
		for (int i = 2; i < 15; i++) {
			workBook.setColWidth(i, 5000);
		}
		workBook.setRowHeight(0, 500);
		workBook.setRowHeight(1, 400);
		workBook.setRowHeight(8, 500);

		workBook.setText(0, 3, "NETWORK OPERATION CENTER");
		workBook.setText(1, 3, "MOBILE DEPARTMENT");
		workBook.setText(2, 3, "----------o0o----------");
		workBook.setText(4, 7, "Date:");
		workBook.setText(4, 8, new SimpleDateFormat("dd/MM/yyyy hh:mm:ss").format(new Date()));
		workBook.setText(5, 7, "Reporter:");
		workBook.setText(5, 8, "FPT Report Schedule");
		workBook.setText(6, 7, "Source:");
		workBook.setText(6, 8, "FPT Network Management System");
		workBook.setText(8, 0, "WEEKLY 3G RAN PERFORMANCE REPORT");
		workBook.setText(9, 5, "Week " + strWeek + " - Month:");
		workBook.setText(9, 6, Util.getMonthNames(strDate));
		workBook.setText(11, 0, "I. Weekly average summary");
		workBook.setText(73, 0, "II. Average Network KPI per week");

		RangeStyle rangeStyle = workBook.getRangeStyle(0, 3, 0, 8);
		rangeStyle.setFontName("Times New Roman");
		rangeStyle.setFontSize(300);
		rangeStyle.setFontBold(true);
		rangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentCenter);
		rangeStyle.setVerticalAlignment(RangeStyle.VerticalAlignmentCenter);
		rangeStyle.setMergeCells(true);
		for (int i = 0; i < 3; i++) {

			workBook.setRangeStyle(rangeStyle, i, 3, i, 8);
		}

		RangeStyle rangeStyle2 = workBook.getRangeStyle();
		rangeStyle2.setFontBold(true);
		workBook.setRangeStyle(rangeStyle2, 4, 7, 6, 8);
		workBook.setRangeStyle(rangeStyle2, 9, 5, 9, 6);
		workBook.setRangeStyle(rangeStyle2, 11, 0, 11, 0);
		workBook.setRangeStyle(rangeStyle2, 73, 0, 73, 0);

		RangeStyle rangeStyle3 = workBook.getRangeStyle(8, 0, 8, 11);
		rangeStyle3.setFontName("Times New Roman");
		rangeStyle3.setFontSize(350);
		rangeStyle3.setFontBold(true);
		rangeStyle3.setHorizontalAlignment(RangeStyle.HorizontalAlignmentCenter);
		rangeStyle3.setVerticalAlignment(RangeStyle.VerticalAlignmentCenter);
		rangeStyle3.setMergeCells(true);
		workBook.setRangeStyle(rangeStyle3, 8, 0, 8, 11);

	}

	/**
	 * 
	 * @param workBook
	 * @throws Exception
	 */
	private void initSheet(WorkBook workBook) throws Exception {

		workBook.copySheet(Util.QOS_KPIS_NUM);
		workBook.copySheet(Util.TRAFFIC_NUM);
		workBook.setSheetName(Util.QOS_KPIS_NUM, "QoS_KPIs");
		workBook.setSheetName(Util.TRAFFIC_NUM, "Traffic_KPIs");
		workBook.setSheetName(Util.UTILIZATION_NUM, "Utilization_KPIs");

	}

	/**
	 * 
	 * @param workBook
	 * @param sheetType
	 * @throws Exception
	 */
	private void createSheet(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

		workBook.setSheet(sheetNumber);
		initFirstTable(workBook, sheetType, sheetNumber);
		initSecondTable(workBook, sheetType, sheetNumber);
		initThirdTable(workBook, sheetType, sheetNumber);
		workBook.setColHidden(colEnd + 2, true);
		workBook.setColHidden(colEnd + 3, true);

	}

	/**
	 * 
	 * @param workBook
	 * @param sheetType
	 * @throws Exception
	 */
	private void initFirstTable(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

		Util.putPosition(ReportConfig.lstReportKpi, sheetType, Util.ROW_AVG_START, Util.ROW_AVG_NEXT);
		SmartUtil.mergeCells(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, 2, true);
		workBook.setText(Util.ROW_AVG_START - 1, 1, "KPI Name");
		colEnd = getColHeaderEnd(workBook, Util.ROW_AVG_START - 1, 4, strWeek, strDate);
		workBook.setText(Util.ROW_AVG_START - 1, 3, "Target");
		workBook.setText(Util.ROW_AVG_START - 1, colEnd + 1, "Weekly average");
		SmartUtil.setDefaultStyle(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, colEnd + 1);
		SmartUtil.fillColorPatten(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, colEnd + 1, Color.LIGHT_GRAY.getRGB());
		fillDataForFirstTable(workBook, sheetType, sheetNumber);
		SmartUtil.setFullBorder(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START + 3 * ReportConfig.getNumOfKPI(sheetType) - 1, colEnd + 1, RangeStyle.BorderThin);
		workBook.setSelection(Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, colEnd + 1);
		workBook.autoFilter();
	}

	/**
	 * 
	 * @param workBook
	 * @param sheetType
	 * @throws Exception
	 */

	private void fillDataForFirstTable(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

		int allNumberSeries = 0;
		int hwNumberSeries = 0;
		int alcatelNumberSeries = 0;

		// ChartShape allChart = workBook.getChart(Util.ALL_NETWORK_CHART);
		ChartShape allChart = workBook.addChart(1, 13, colEnd + 1, 32);
		allChart.setChartType(ChartShape.Line);
		allChart.setLinkRange(SmartUtil.getFormula(workBook.getSheetName(sheetNumber), Util.ROW_AVG_START, 4, Util.ROW_AVG_START + 1, colEnd), true);
		allChart.setTitle("All Network Weekly Average");

		ChartFormat chartFormat = allChart.getPlotFormat();
		chartFormat.setSolid();
		chartFormat.setForeColor(Color.WHITE.getRGB());

		ChartFormat legendFormat = allChart.getLegendFormat();
		legendFormat.setSolid();
		legendFormat.setForeColor(Color.LIGHT_GRAY.getRGB());

		ChartFormat titleformat = allChart.getTitleFormat();
		titleformat.setFontSize(300);
		titleformat.setFontBold(true);

		allChart.setTitleFormat(titleformat);
		allChart.setPlotFormat(chartFormat);
		allChart.setLegendFormat(legendFormat);
		allChart.setLegendPosition(ChartFormat.LegendPlacementBottom);

		// ChartShape hwChart = workBook.getChart(Util.HUAWEI_CHART);
		ChartShape hwChart = workBook.addChart(1, 33, colEnd + 1, 52);
		hwChart.setChartType(ChartShape.Line);
		hwChart.setLinkRange(SmartUtil.getFormula(workBook.getSheetName(sheetNumber), Util.ROW_AVG_START, 4, Util.ROW_AVG_START + 1, colEnd), true);
		hwChart.setTitle("Huawei Weekly Average");

		hwChart.setTitleFormat(titleformat);
		hwChart.setPlotFormat(chartFormat);
		hwChart.setLegendFormat(legendFormat);
		hwChart.setLegendPosition(ChartFormat.LegendPlacementBottom);

		// ChartShape alcatelChart = workBook.getChart(Util.ALCATEL_CHART);
		ChartShape alcatelChart = workBook.addChart(1, 53, colEnd + 1, 72);
		alcatelChart.setChartType(ChartShape.Line);
		alcatelChart.setLinkRange(SmartUtil.getFormula(workBook.getSheetName(sheetNumber), Util.ROW_AVG_START, 4, Util.ROW_AVG_START + 1, colEnd), true);
		alcatelChart.setTitle("Alcatel Weekly Average");

		alcatelChart.setTitleFormat(titleformat);
		alcatelChart.setPlotFormat(chartFormat);
		alcatelChart.setLegendFormat(legendFormat);
		alcatelChart.setLegendPosition(ChartFormat.LegendPlacementBottom);

		for (int i = 0; i < vtDataPerHour.size(); i++) {
			Vector<String> vtRow = (Vector<String>) vtDataPerHour.get(i);
			String kpiCode = (String) vtRow.get(vtRow.size() - 2);
			Integer row = Util.getPosition(kpiCode, sheetType);
			if (row != null) {
				workBook.setText(row, 1, ReportConfig.mapKpiName.get(kpiCode));
				workBook.setText(row, 2, "All Network");
				SmartUtil.setStyleString(workBook, row, 1, row, 2);
				Vector<String> kpiValue = new Vector<String>();
				for (int j = 0; j < vtRow.size() - 2; j++) {
					SmartUtil.setNumber(workBook, row, 4 + j, (String) vtRow.get(j));
					SmartUtil.setStyleNumber(workBook, row, 4, row, 4 + j);
					kpiValue.add(vtRow.get(j));
				}
				kpiMap.put(kpiCode, kpiValue);
				SmartUtil.setNumber(workBook, row, colEnd + 1, Util.getAverageStringNumber(vtRow));
				SmartUtil.setStyleNumber(workBook, row, colEnd + 1, row, colEnd + 1);
				SmartUtil.addSeries(allChart, ReportConfig.mapKpiName.get(kpiCode), SmartUtil.getFormula(workBook.getSheetName(sheetNumber), row + 1, 4, row + 1, colEnd), allNumberSeries++);

			}
		}

		for (int i = 0; i < vtDataStatic.size(); i++) {
			Vector<String> vtRow = (Vector<String>) vtDataStatic.get(i);
			String supplierId = (String) vtRow.get(vtRow.size() - 1);
			String kpiCode = (String) vtRow.get(vtRow.size() - 2);
			int typeCode = Integer.parseInt(vtRow.get(vtRow.size() - 3));
			if (typeCode == Util.TYPE_CODE_AVG) {
				Integer row = Util.getPosition(kpiCode, sheetType);
				if (row != null) {
					String supplierName = supplierId.equals(Util.HUAWEI_TYPE) ? "Huawei" : "Alcatel";
					for (int j = 0; j < vtRow.size() - 2; j++) {
						if (supplierId.equals(Util.ALCATEL_TYPE)) {
							workBook.setText(row + 1, 2, supplierName);
							SmartUtil.setStyleString(workBook, row + 1, 2, row + 1, 2);
							SmartUtil.setNumber(workBook, row + 1, 4 + j, (String) vtRow.get(j));
							SmartUtil.setStyleNumber(workBook, row + 1, 4, row + 1, 4 + j);

						} else {
							workBook.setText(row + 2, 2, supplierName);
							SmartUtil.setStyleString(workBook, row + 2, 2, row + 2, 2);
							SmartUtil.setNumber(workBook, row + 2, 4 + j, (String) vtRow.get(j));
							SmartUtil.setStyleNumber(workBook, row + 2, 4, row + 2, 4 + j);
						}
					}
					if (supplierId.equals(Util.ALCATEL_TYPE)) {

						SmartUtil.setNumber(workBook, row + 1, colEnd + 1, Util.getAverageStringNumber(vtRow));
						SmartUtil.setStyleNumber(workBook, row + 1, colEnd + 1, row + 1, colEnd + 1);
						SmartUtil.addSeries(alcatelChart, ReportConfig.mapKpiName.get(kpiCode), SmartUtil.getFormula(workBook.getSheetName(sheetNumber), row + 2, 4, row + 2, colEnd),
						        alcatelNumberSeries++);

					} else {

						SmartUtil.setNumber(workBook, row + 2, colEnd + 1, Util.getAverageStringNumber(vtRow));
						SmartUtil.setStyleNumber(workBook, row + 2, colEnd + 1, row + 2, colEnd + 1);
						SmartUtil.addSeries(hwChart, ReportConfig.mapKpiName.get(kpiCode), SmartUtil.getFormula(workBook.getSheetName(sheetNumber), row + 3, 4, row + 3, colEnd), hwNumberSeries++);

					}
				}
			}
		}

		rowStartSecondTable = allChart.getSeriesCount() * 3 + Util.ROW_AVG_START;

	}

	/**
	 * 
	 * @param workBook
	 * @param sheetType
	 * @throws Exception
	 */
	private void initSecondTable(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

		workBook.setText(rowStartSecondTable - 2, 0, "III. Weekly KPI statistics");
		SmartUtil.adjustFont(workBook, 0, true, false, false, rowStartSecondTable - 2, 0, rowStartSecondTable - 2, 0);

		Util.putPosition(ReportConfig.lstReportKpi, sheetType, rowStartSecondTable + 1, sheetType == 2 ? 6 : 19);
		// SmartUtil.mergeCells(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, 2, true);
		workBook.setText(rowStartSecondTable, 1, "KPI Name");
		workBook.setText(rowStartSecondTable, 2, "Vendor");
		workBook.setText(rowStartSecondTable, 3, "Criteria");
		getColHeaderEnd(workBook, rowStartSecondTable, 4, strWeek, strDate);
		workBook.setText(rowStartSecondTable, colEnd + 1, "Weekly average");
		SmartUtil.setDefaultStyle(workBook, rowStartSecondTable, 1, rowStartSecondTable, colEnd + 1);
		SmartUtil.fillColorPatten(workBook, rowStartSecondTable, 1, rowStartSecondTable, colEnd + 1, Color.LIGHT_GRAY.getRGB());
		fillDataForSecondTable(workBook, sheetType, sheetNumber);

		SmartUtil.setFullBorder(workBook, rowStartSecondTable, 1, rowStartSecondTable + (sheetType == 2 ? 6 : 19) * ReportConfig.getNumOfKPI(sheetType), colEnd + 1, RangeStyle.BorderThin);
		workBook.setSelection(rowStartSecondTable, 1, rowStartSecondTable, colEnd + 1);
		workBook.autoFilter();
	}

	private void fillDataForSecondTable(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

		for (int i = 0; i < vtDataPerHour.size(); i++) {
			Vector<String> vtRow = (Vector<String>) vtDataPerHour.get(i);
			String kpiCode = (String) vtRow.get(vtRow.size() - 2);
			Integer row = Util.getPosition(kpiCode, sheetType);
			if (row != null) {

				workBook.setText(row, 1, ReportConfig.mapKpiName.get(kpiCode));

				SmartUtil.mergeCells(workBook, row, 1, row + (sheetType == 2 ? 5 : 18), 1, true);
				SmartUtil.setStyleString(workBook, row, 1, row + (sheetType == 2 ? 5 : 18), 1);
				workBook.setText(row, 2, "All Network");
				SmartUtil.mergeCells(workBook, row, 2, row + (sheetType == 2 ? 1 : 6), 2, true);
				SmartUtil.setStyleString(workBook, row, 2, row + (sheetType == 2 ? 1 : 6), 2);

				if (sheetType != 2) {
					workBook.setRowOutlineLevel(row + 3, row + 18, 1, false);
				}
				if (sheetType == 1 && Float.parseFloat(vtRow.get(0)) > 90) {

					for (int j = 0; j < arrCriteria1.length; j++) {

						workBook.setText(row + j, 3, arrCriteria1[j]);
						SmartUtil.setStyleString(workBook, row + j, 3, row + j, 3);

						if (j != 1 && j != 2) {
							for (int k = 4; k <= colEnd + 1; k++) {

								workBook.setFormula(
								        row + j,
								        k,
								        "SUMIF(" + workBook.formatRCNr(row + 7, 3, true) + ":" + workBook.formatRCNr(row + 18, 3, true) + "," + workBook.formatRCNr(row + j, 3, true) + ","
								                + workBook.formatRCNr(row + 7, k, false) + ":" + workBook.formatRCNr(row + 18, k, false) + ")");
							}
						}
					}
				} else if (sheetType == 1 && Float.parseFloat(vtRow.get(0)) < 5) {

					for (int j = 0; j < arrCriteria2.length; j++) {

						workBook.setText(row + j, 3, arrCriteria2[j]);
						SmartUtil.setStyleString(workBook, row + j, 3, row + j, 3);

						if (j != 1 && j != 2) {
							for (int k = 4; k <= colEnd + 1; k++) {

								workBook.setFormula(
								        row + j,
								        k,
								        "SUMIF(" + workBook.formatRCNr(row + 7, 3, true) + ":" + workBook.formatRCNr(row + 18, 3, true) + "," + workBook.formatRCNr(row + j, 3, true) + ","
								                + workBook.formatRCNr(row + 7, k, false) + ":" + workBook.formatRCNr(row + 18, k, false) + ")");
							}
						}

					}

				} else if (sheetType == 2) {

					for (int j = 0; j < arrCriteria3.length; j++) {

						workBook.setText(row + j, 3, arrCriteria3[j]);
						SmartUtil.setStyleString(workBook, row + j, 3, row + j, 3);

						if (j != 1) {
							for (int k = 4; k <= colEnd + 1; k++) {

								workBook.setFormula(
								        row + j,
								        k,
								        "SUMIF(" + workBook.formatRCNr(row + 2, 3, true) + ":" + workBook.formatRCNr(row + 5, 3, true) + "," + workBook.formatRCNr(row + j, 3, true) + ","
								                + workBook.formatRCNr(row + 2, k, false) + ":" + workBook.formatRCNr(row + 5, k, false) + ")");
							}

						}
					}

				} else if (sheetType == 3) {

					for (int j = 0; j < arrCriteria4.length; j++) {

						workBook.setText(row + j, 3, arrCriteria4[j]);
						SmartUtil.setStyleString(workBook, row + j, 3, row + j, 3);

						if (j != 1 && j != 2) {
							for (int k = 4; k <= colEnd + 1; k++) {

								workBook.setFormula(
								        row + j,
								        k,
								        "SUMIF(" + workBook.formatRCNr(row + 7, 3, true) + ":" + workBook.formatRCNr(row + 18, 3, true) + "," + workBook.formatRCNr(row + j, 3, true) + ","
								                + workBook.formatRCNr(row + 7, k, false) + ":" + workBook.formatRCNr(row + 18, k, false) + ")");
							}

						}
					}

				}

				for (int j = 0; j < vtRow.size() - 2; j++) {

					SmartUtil.setNumber(workBook, row + (sheetType == 2 ? 1 : 2), 4 + j, (String) vtRow.get(j));
					SmartUtil.setStyleNumber(workBook, row + (sheetType == 2 ? 1 : 2), 4, row + (sheetType == 2 ? 1 : 2), 4 + j);
				}
				SmartUtil.setNumber(workBook, row + (sheetType == 2 ? 1 : 2), colEnd + 1, Util.getAverageStringNumber(vtRow));
				SmartUtil.setStyleNumber(workBook, row + (sheetType == 2 ? 1 : 2), colEnd + 1, row + (sheetType == 2 ? 1 : 2), colEnd + 1);

			}

		}

		for (int i = 0; i < vtDataStatic.size(); i++) {

			Vector<String> vtRow = (Vector<String>) vtDataStatic.get(i);
			String supplierId = (String) vtRow.get(vtRow.size() - 1);
			String kpiCode = (String) vtRow.get(vtRow.size() - 2);
			int typeCode = Integer.parseInt(vtRow.get(vtRow.size() - 3));

			if (typeCode == Util.TYPE_CODE_AVG) {
				Integer row = Util.getPosition(kpiCode, sheetType);

				if (row != null) {

					String supplierName = supplierId.equals(Util.HUAWEI_TYPE) ? "Huawei" : "Alcatel";
					for (int j = 0; j < vtRow.size() - 2; j++) {

						if (supplierId.equals(Util.ALCATEL_TYPE)) {

							workBook.setText(row + (sheetType == 2 ? 2 : 7), 2, supplierName);
							SmartUtil.mergeCells(workBook, row + (sheetType == 2 ? 2 : 7), 2, row + (sheetType == 2 ? 3 : 12), 2, true);
							SmartUtil.setStyleString(workBook, row + (sheetType == 2 ? 2 : 7), 2, row + (sheetType == 2 ? 3 : 12), 2);

							if (sheetType == 1 && Float.parseFloat(vtRow.get(0)) > 90) {

								for (int k = 0; k < arrCriteria1.length; k++) {
									if (k != 1) {
										int rowAlcatel = row + 7 + (k == 0 ? k : k - 1);
										workBook.setText(rowAlcatel, 3, arrCriteria1[k]);
										SmartUtil.setStyleString(workBook, rowAlcatel, 3, rowAlcatel, 3);
									}
								}
							} else if (sheetType == 1 && Float.parseFloat(vtRow.get(0)) < 5) {

								for (int k = 0; k < arrCriteria2.length; k++) {
									if (k != 1) {
										int rowAlcatel = row + 7 + (k == 0 ? k : k - 1);
										workBook.setText(rowAlcatel, 3, arrCriteria2[k]);
										SmartUtil.setStyleString(workBook, rowAlcatel, 3, rowAlcatel, 3);
									}
								}

							} else if (sheetType == 2) {

								for (int k = 0; k < arrCriteria3.length; k++) {

									int rowAlcatel = row + 2 + k;
									workBook.setText(rowAlcatel, 3, arrCriteria3[k]);
									SmartUtil.setStyleString(workBook, rowAlcatel, 3, rowAlcatel, 3);
								}

							} else if (sheetType == 3) {

								for (int k = 0; k < arrCriteria4.length; k++) {
									if (k != 1) {
										int rowAlcatel = row + 7 + (k == 0 ? k : k - 1);
										workBook.setText(rowAlcatel, 3, arrCriteria4[k]);
										SmartUtil.setStyleString(workBook, rowAlcatel, 3, rowAlcatel, 3);
									}
								}

							}

							SmartUtil.setNumber(workBook, row + (sheetType == 2 ? 3 : 8), 4 + j, (String) vtRow.get(j));
							SmartUtil.setStyleNumber(workBook, row + (sheetType == 2 ? 3 : 8), 4, row + (sheetType == 2 ? 3 : 8), 4 + j);

						} else {

							workBook.setText(row + (sheetType == 2 ? 4 : 13), 2, supplierName);
							SmartUtil.mergeCells(workBook, row + (sheetType == 2 ? 4 : 13), 2, row + (sheetType == 2 ? 5 : 18), 2, true);
							SmartUtil.setStyleString(workBook, row + (sheetType == 2 ? 4 : 13), 2, row + (sheetType == 2 ? 5 : 18), 2);

							if (sheetType == 1 && Float.parseFloat(vtRow.get(0)) > 90) {

								for (int k = 0; k < arrCriteria1.length; k++) {

									if (k != 1) {
										int rowHw = row + 13 + (k == 0 ? k : k - 1);
										workBook.setText(rowHw, 3, arrCriteria1[k]);
										SmartUtil.setStyleString(workBook, rowHw, 3, rowHw, 3);
									}
								}
							} else if (sheetType == 1 && Float.parseFloat(vtRow.get(0)) < 5) {

								for (int k = 0; k < arrCriteria2.length; k++) {

									if (k != 1) {
										int rowHw = row + 13 + (k == 0 ? k : k - 1);
										workBook.setText(rowHw, 3, arrCriteria2[k]);
										SmartUtil.setStyleString(workBook, rowHw, 3, rowHw, 3);
									}
								}

							} else if (sheetType == 2) {

								for (int k = 0; k < arrCriteria3.length; k++) {

									int rowHw = row + 4 + k;
									workBook.setText(rowHw, 3, arrCriteria3[k]);
									SmartUtil.setStyleString(workBook, rowHw, 3, rowHw, 3);
								}

							} else if (sheetType == 3) {

								for (int k = 0; k < arrCriteria4.length; k++) {

									if (k != 1) {
										int rowHw = row + 13 + (k == 0 ? k : k - 1);
										workBook.setText(rowHw, 3, arrCriteria4[k]);
										SmartUtil.setStyleString(workBook, rowHw, 3, rowHw, 3);
									}
								}

							}

							SmartUtil.setNumber(workBook, row + (sheetType == 2 ? 5 : 14), 4 + j, (String) vtRow.get(j));
							SmartUtil.setStyleNumber(workBook, row + (sheetType == 2 ? 5 : 14), 4, row + (sheetType == 2 ? 5 : 14), 4 + j);
						}

					}
					if (supplierId.equals(Util.ALCATEL_TYPE)) {

						SmartUtil.setNumber(workBook, row + (sheetType == 2 ? 3 : 8), colEnd + 1, Util.getAverageStringNumber(vtRow));
						SmartUtil.setStyleNumber(workBook, row + (sheetType == 2 ? 3 : 8), colEnd + 1, row + (sheetType == 2 ? 3 : 8), colEnd + 1);

					} else {

						SmartUtil.setNumber(workBook, row + (sheetType == 2 ? 5 : 14), colEnd + 1, Util.getAverageStringNumber(vtRow));
						SmartUtil.setStyleNumber(workBook, row + (sheetType == 2 ? 5 : 14), colEnd + 1, row + (sheetType == 2 ? 5 : 14), colEnd + 1);

					}

					rowStartThirdTable = row + (sheetType == 2 ? 9 : 22);
				}
			}
		}
	}

	/**
	 * 
	 * @param workBook
	 * @param sheetType
	 * @throws Exception
	 */
	private void initThirdTable(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

		workBook.setText(rowStartThirdTable - 2, 0, "IV. Weekly average KPI per zone/area");
		SmartUtil.adjustFont(workBook, 0, true, false, false, rowStartThirdTable - 2, 0, rowStartThirdTable - 2, 0);

		Util.putPosition(ReportConfig.lstReportKpi, sheetType, rowStartThirdTable + 24, 23);
		// SmartUtil.mergeCells(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, 2, true);
		workBook.setText(rowStartThirdTable, 1, "KPI Name");
		workBook.setText(rowStartThirdTable, 2, "Center");
		workBook.setText(rowStartThirdTable, 3, "Province");
		getColHeaderEnd(workBook, rowStartThirdTable, 4, strWeek, strDate);
		workBook.setText(rowStartThirdTable, colEnd + 1, "Weekly average");
		SmartUtil.setDefaultStyle(workBook, rowStartThirdTable, 1, rowStartThirdTable, colEnd + 1);
		SmartUtil.fillColorPatten(workBook, rowStartThirdTable, 1, rowStartThirdTable, colEnd + 1, Color.LIGHT_GRAY.getRGB());
		fillDataForThirdTable(workBook, sheetType, sheetNumber);

		SmartUtil.setFullBorder(workBook, rowStartThirdTable, 1, rowStartThirdTable + 23 * (ReportConfig.getNumOfKPI(sheetType) + 1), colEnd + 1, RangeStyle.BorderThin);
		workBook.setSelection(rowStartThirdTable, 1, rowStartThirdTable, colEnd + 1);
		workBook.autoFilter();
	}

	private void fillDataForThirdTable(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

		Vector<Vector<String>> dataAreaHeader = new Vector<Vector<String>>();
		dataAreaHeader.addAll(sheetType == 3 ? vtDataAreaNodeb : vtDataAreaCells);
		Integer row = rowStartThirdTable + 1;
		workBook.setRowOutlineLevel(row + 1, row + 22, 1, false);
		SmartUtil.setStyleString(workBook, row, 1, row + 22, 3);
		SmartUtil.alignCenter(workBook, row, 2, row + 22, 3);
		SmartUtil.setStyleNumber(workBook, row, 4, row + 22, colEnd + 1);
		for (int i = 0; i < dataAreaHeader.size(); i++) {

			Vector<String> vtRow = (Vector<String>) dataAreaHeader.get(i);

			workBook.setText(row, 1, vtRow.get(vtRow.size() - 5));
			SmartUtil.mergeCells(workBook, row, 1, row + 22, 1, true);

			workBook.setText(row, 2, "All Network");
			SmartUtil.mergeCells(workBook, row, 2, row, 3, true);

			if (vtRow.get(vtRow.size() - 1).equals("2")) {

				workBook.setText(row + i + 1, 2, vtRow.get(vtRow.size() - 2));
				SmartUtil.mergeCells(workBook, row + i + 1, 2, row + i + 1, 3, vtRow.get(vtRow.size() - 1).equals("2"));
			} else {

				workBook.setText(row + i + 1, 2, vtRow.get(vtRow.size() - 4));
				workBook.setText(row + i + 1, 3, vtRow.get(vtRow.size() - 3));
			}

			for (int j = 0; j < vtRow.size() - 5; j++) {

				workBook.setFormula(row, 4 + j,
				        "SUMIF(" + workBook.formatRCNr(row + 1, colEnd + 2, true) + ":" + workBook.formatRCNr(row + 22, colEnd + 2, true) + ",\"2\"," + workBook.formatRCNr(row + 1, 4 + j, false)
				                + ":" + workBook.formatRCNr(row + 22, 4 + j, false) + ")");
				SmartUtil.setNumber(workBook, row + i + 1, 4 + j, (String) vtRow.get(j));
			}

			workBook.setFormula(row, colEnd + 1, "MAX(" + workBook.formatRCNr(row, 4, false) + ":" + workBook.formatRCNr(row, colEnd, false) + ")");
			workBook.setFormula(row + i + 1, colEnd + 1, "MAX(" + workBook.formatRCNr(row + i + 1, 4, false) + ":" + workBook.formatRCNr(row + i + 1, colEnd, false) + ")");
			SmartUtil.setNumber(workBook, row + i + 1, colEnd + 2, vtRow.get(vtRow.size() - 1));

		}

		int count = 0;
		for (int i = 0; i < vtDataArea.size(); i++) {

			count = count == 22 ? 0 : count;
			Vector<String> vtRow = (Vector<String>) vtDataArea.get(i);
			String kpiCode = (String) vtRow.get(vtRow.size() - 5);
			row = Util.getPosition(kpiCode, sheetType);
			if (row != null) {

				workBook.setRowOutlineLevel(row + 1, row + 22, 1, false);

				SmartUtil.setStyleString(workBook, row, 1, row + 22, 3);
				SmartUtil.alignCenter(workBook, row, 2, row + 22, 3);
				SmartUtil.setStyleNumber(workBook, row, 4, row + 22, colEnd + 1);

				workBook.setText(row, 1, ReportConfig.mapKpiName.get(kpiCode));
				SmartUtil.mergeCells(workBook, row, 1, row + 22, 1, true);

				workBook.setText(row, 2, "All Network");
				SmartUtil.mergeCells(workBook, row, 2, row, 3, true);
				workBook.setFormula(row, colEnd + 1, "MAX(" + workBook.formatRCNr(row, 4, false) + ":" + workBook.formatRCNr(row, colEnd, false) + ")");

				for (int j = 0; j < vtRow.size() - 5; j++) {
					SmartUtil.setNumber(workBook, row, 4 + j, kpiMap.get(kpiCode).get(j));
					SmartUtil.setNumber(workBook, row + count + 1, 4 + j, (String) vtRow.get(j));
				}

				if (vtRow.get(vtRow.size() - 1).equals("2")) {

					workBook.setText(row + count + 1, 2, vtRow.get(vtRow.size() - 2));
					SmartUtil.mergeCells(workBook, row + count + 1, 2, row + count + 1, 3, vtRow.get(vtRow.size() - 1).equals("2"));
				} else {

					workBook.setText(row + count + 1, 2, vtRow.get(vtRow.size() - 4));
					workBook.setText(row + count + 1, 3, vtRow.get(vtRow.size() - 3));
				}
				workBook.setFormula(row + count + 1, colEnd + 1, "MAX(" + workBook.formatRCNr(row + count + 1, 4, false) + ":" + workBook.formatRCNr(row + count + 1, colEnd, false) + ")");
				SmartUtil.setNumber(workBook, row + count + 1, colEnd + 2, vtRow.get(vtRow.size() - 1));

				count++;
			}
		}

	}

	/**
	 * @param workBook
	 * @param row
	 * @param col
	 * @param vstrWeek
	 * @param strDate
	 * @return
	 * @throws Exception
	 */
	private int getColHeaderEnd(WorkBook workBook, int row, int col, String vstrWeek, String strDate) throws Exception {

		int nextCol = col;
		switch (vstrWeek) {
		case Util.WEEK_1:
			for (int i = 1; i <= 7; i++) {
				workBook.setText(row, nextCol, "0" + i + "-" + Util.getMonthNames(strDate));
				nextCol++;
			}
			break;
		case Util.WEEK_2:
			for (int i = 8; i <= 14; i++) {
				if (i < 10)
					workBook.setText(row, nextCol, "0" + i + "-" + Util.getMonthNames(strDate));
				else
					workBook.setText(row, nextCol, i + "-" + Util.getMonthNames(strDate));
				nextCol++;
			}
			break;
		case Util.WEEK_3:
			for (int i = 15; i <= 21; i++) {
				workBook.setText(row, nextCol, i + "-" + Util.getMonthNames(strDate));
				nextCol++;
			}
			break;
		case Util.WEEK_4:
			for (int i = 22; i <= Util.getDaysInMonth(strDate); i++) {
				workBook.setText(row, nextCol, i + "-" + Util.getMonthNames(strDate));
				nextCol++;
			}
		}
		return nextCol - 1;
	}

	/**
	 * @param vstrWeek
	 * @param vstrStartDate
	 * @throws Exception
	 */
	public void pmsWeeklyReport(Connection mcnMain, String vstrWeek, String vstrStartDate) throws Exception {

		PreparedStatement psmtPerHour = null;
		ResultSet rsPerHour = null;
		PreparedStatement psmtStatic = null;
		ResultSet rsStatic = null;
		PreparedStatement psmtAvgWeek = null;
		ResultSet rsAvgWeek = null;
		PreparedStatement psmtArea = null;
		ResultSet rsArea = null;

		try {
			// ===================================================================
			// Get data I.1: Per hour
			String strSQLDataPerHour = " select * from (select to_char (event_time, 'dd') date_time, network_id, a.kpi_code, sheet_id, kpi_order, round(ave,2) ave "
			        + " from (select to_char (event_time, 'dd') date_time, network_id, event_time, kpi_code, avg(value) ave "
			        + "	from pms_average_by_cell pms  where event_time >= to_date (?, 'dd/MM/yyyy') ";
			if (vstrWeek.equals(Util.WEEK_4))
				strSQLDataPerHour += " and event_time < last_day(to_date(?,'dd/MM/YYYY')) + 1";
			else
				strSQLDataPerHour += " and event_time < to_date(?, 'dd/MM/yyyy') + 7 ";

			strSQLDataPerHour += " and network_id = 9 and period = 2 and value not in ('NaN','Infinity','-Infinity') " + " group by network_id,event_time, kpi_code )  a, report_kpi b "
			        + " where a.kpi_code = b.kpi_code) pivot (sum(ave) for (date_time) in (" + getListDays(vstrWeek) + " )) order by sheet_id,kpi_order";

			psmtPerHour = mcnMain.prepareStatement(strSQLDataPerHour);
			psmtPerHour.setString(1, vstrStartDate);
			psmtPerHour.setString(2, vstrStartDate);
			rsPerHour = psmtPerHour.executeQuery();
			while (rsPerHour.next()) {
				Vector<String> vt = new Vector<String>();
				addElement(vt, rsPerHour, vstrWeek);
				vt.addElement(rsPerHour.getString("kpi_code"));
				vt.addElement(rsPerHour.getString("sheet_id"));
				vtDataPerHour.add(vt);
			}
			// get weekly average by hour
			String strSQLAVGWeek = " select network_id, supplier_id, alarm_level, kpi_code, type type_code, value from pms_daily_by_week pms "
			        + " where event_time = to_date (?, 'dd/MM/yyyy') and network_id = 9";

			psmtAvgWeek = mcnMain.prepareStatement(strSQLAVGWeek);
			psmtAvgWeek.setString(1, vstrStartDate);
			rsAvgWeek = psmtPerHour.executeQuery();
			while (rsAvgWeek.next()) {

			}

			// Get Date I.2: Static by Supplier and Alarm level
			String sqlDataStatic = " select b.sheet_id, a.* from (select network_id,supplier_id, " + " (select name from supplier where supplier_id = pms.supplier_id) supplier_name,alarm_level_id, "
			        + " (select code from alarm_level where alarm_level_id = pms.alarm_level_id) alarm_name, "
			        + " to_char (event_time, 'dd') date_time,kpi_code, type_code, round(value,2) value from pms_daily_by_supplier pms " + " where event_time >= to_date (?, 'dd/mm/yyyy') ";
			if (vstrWeek.equals(Util.WEEK_4))
				sqlDataStatic += " and event_time < last_day(to_date(?,'dd/MM/yyyy')) + 1";
			else
				sqlDataStatic += " and event_time < to_date(?,'dd/MM/yyyy') + 7 ";

			sqlDataStatic += " and network_id = 9 and type_code < 4 ) pivot(sum (value) for (date_time) in (" + getListDays(vstrWeek)
			        + " )) a, report_kpi b where a.kpi_code = b.kpi_code order by b.sheet_id, b.kpi_order  ";

			psmtStatic = mcnMain.prepareStatement(sqlDataStatic);
			psmtStatic.setString(1, vstrStartDate);
			psmtStatic.setString(2, vstrStartDate);
			rsStatic = psmtStatic.executeQuery();
			while (rsStatic.next()) {
				Vector<String> vt = new Vector<String>();
				addElement(vt, rsStatic, vstrWeek);
				vt.addElement(rsStatic.getString("alarm_name"));
				vt.addElement(rsStatic.getString("type_code"));
				vt.addElement(rsStatic.getString("kpi_code"));
				vt.addElement(rsStatic.getString("supplier_id"));
				vtDataStatic.add(vt);
			}

			// Get Data I.3: tinh theo Area code

			String sqlDataArea = "SELECT * FROM (SELECT network_id, area_code, (select name from area where area_code = pms.area_code) area_name, kpi_code, avg(value) ave, center_code, center_type, to_char(event_time, 'dd') date_time "
			        + " FROM pms_daily_by_area pms WHERE event_time >= to_date (?, 'dd/MM/yyyy') ";

			if (vstrWeek.equals(Util.WEEK_4))
				sqlDataArea += "and event_time < last_day(to_date(?,'dd/MM/yyyy')) + 1";
			else
				sqlDataArea += "and event_time < to_date(?,'dd/MM/yyyy') + 7";

			sqlDataArea += " AND network_id = 9 AND (area_code <> 'UNIDENTIFIED' OR area_code is null) GROUP BY pms.network_id, pms.event_time, pms.area_code, pms.kpi_code, pms.center_code, pms.center_type) "
			        + "PIVOT (SUM (ave) FOR (date_time) IN (" + getListDays(vstrWeek) + ")) " + " ORDER BY kpi_code, center_code, area_name ";

			psmtArea = mcnMain.prepareStatement(sqlDataArea);
			psmtArea.setString(1, vstrStartDate);
			psmtArea.setString(2, vstrStartDate);
			rsArea = psmtArea.executeQuery();

			while (rsArea.next()) {

				Vector<String> vt = new Vector<String>();
				Vector<String> vtCells = new Vector<String>();
				Vector<String> vtNodeb = new Vector<String>();

				if (rsArea.getString("kpi_code").equalsIgnoreCase("CELLS")) {

					addElement(vtCells, rsArea, vstrWeek);
					vtCells.addElement("Total Cells");
					vtCells.addElement(rsArea.getString("area_code"));
					vtCells.addElement(rsArea.getString("area_name"));
					vtCells.addElement(rsArea.getString("center_code"));
					vtCells.addElement(rsArea.getString("center_type"));
					vtDataAreaCells.add(vtCells);

				} else if (rsArea.getString("kpi_code").equalsIgnoreCase("NODE B")) {

					addElement(vtNodeb, rsArea, vstrWeek);
					vtNodeb.addElement("Number of Node B");
					vtNodeb.addElement(rsArea.getString("area_code"));
					vtNodeb.addElement(rsArea.getString("area_name"));
					vtNodeb.addElement(rsArea.getString("center_code"));
					vtNodeb.addElement(rsArea.getString("center_type"));
					vtDataAreaNodeb.add(vtNodeb);

				} else {

					addElement(vt, rsArea, vstrWeek);
					vt.addElement(rsArea.getString("kpi_code"));
					vt.addElement(rsArea.getString("area_code"));
					vt.addElement(rsArea.getString("area_name"));
					vt.addElement(rsArea.getString("center_code"));
					vt.addElement(rsArea.getString("center_type"));
					vtDataArea.add(vt);
				}
			}

		} finally {
			Database.closeObject(rsPerHour);
			Database.closeObject(psmtPerHour);
			Database.closeObject(rsStatic);
			Database.closeObject(psmtStatic);
			Database.closeObject(rsAvgWeek);
			Database.closeObject(rsAvgWeek);
		}

	}

	/**
	 * @param vt
	 * @param rs
	 * @param vstrWeek
	 * @throws SQLException
	 */
	public static void addElement(Vector<String> vt, ResultSet rs, String vstrWeek) throws SQLException {

		switch (vstrWeek) {
		case Util.WEEK_1:
			for (int i = 1; i <= 7; i++) {
				vt.addElement(rs.getString("0" + i + ""));
			}
			break;
		case Util.WEEK_2:
			for (int i = 8; i <= 14; i++) {
				if (i < 10)
					vt.addElement(rs.getString("0" + i + ""));
				else
					vt.addElement(rs.getString("" + i + ""));
			}
			break;
		case Util.WEEK_3:
			for (int i = 15; i <= 21; i++) {
				vt.addElement(rs.getString("" + i + ""));
			}
			break;
		case Util.WEEK_4:
			for (int i = 22; i <= 31; i++) {
				vt.addElement(rs.getString("" + i + ""));
			}
		}
	}

	/**
	 * 
	 * @param strWeek
	 * @return list days of week
	 */
	public static String getListDays(String strWeek) {

		String days = "";
		switch (strWeek) {
		case Util.WEEK_1:
			for (int i = 1; i <= 7; i++) {
				days += "'0" + i + "' AS \"0" + i + "\",";
			}
			break;
		case Util.WEEK_2:
			for (int i = 8; i <= 14; i++) {
				days += (i < 10) ? ("'0" + i + "' AS \"0" + i + "\",") : ("'" + i + "' AS \"" + i + "\",");
			}
			break;
		case Util.WEEK_3:
			for (int i = 15; i <= 21; i++) {
				days += "'" + i + "' AS \"" + i + "\",";
			}
			break;
		case Util.WEEK_4:
			for (int i = 22; i <= 31; i++) {
				days += "'" + i + "' AS \"" + i + "\",";
			}
		}
		return days.substring(0, days.length() - 1);
	}

	/**
	 * Test
	 * 
	 * @param args
	 * @throws Exception
	 */
	public static void main(String args[]) throws Exception {

		Connection conn = null;
		try {
			com.fis.database.oracle.OracleUniversalConnectionFactory connectionFactory = new com.fis.database.oracle.OracleUniversalConnectionFactory("jdbc:oracle:thin:@10.30.3.12:1521/nms",
			        "NMS_OWNER_LTC", "nms", 1);
			conn = connectionFactory.getConnection();
			ReportConfig config = new ReportConfig(conn);
			config.configLoader();

			Pms3GWeeklyReport pms = new Pms3GWeeklyReport();
			pms.toExcel(conn, "4", "22/10/2014");
		} finally {
			Database.closeObject(conn);
		}
	}
}
