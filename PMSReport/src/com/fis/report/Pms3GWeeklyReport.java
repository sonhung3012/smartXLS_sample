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
	private int colEnd;

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
		fillAverageData(workBook, sheetType, sheetNumber);
		workBook.setColHidden(colEnd + 2, true);
		workBook.setColHidden(colEnd + 3, true);

	}

	/**
	 * 
	 * @param workBook
	 * @param sheetType
	 * @throws Exception
	 */
	private void fillAverageData(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

		Util.putPosition(ReportConfig.lstReportKpi, sheetType, Util.ROW_AVG_START, Util.ROW_AVG_NEXT);
		SmartUtil.mergeCells(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, 2, true);
		workBook.setText(Util.ROW_AVG_START - 1, 1, "KPI Name");
		colEnd = getColHeaderEnd(workBook, Util.ROW_AVG_START - 1, 4, strWeek, strDate);
		workBook.setText(Util.ROW_AVG_START - 1, 3, "Target");
		workBook.setText(Util.ROW_AVG_START - 1, colEnd + 1, "Weekly average");
		SmartUtil.setDefaultStyle(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, colEnd + 1);
		SmartUtil.fillColorPatten(workBook, Util.ROW_AVG_START - 1, 1, Util.ROW_AVG_START - 1, colEnd + 1, Color.LIGHT_GRAY.getRGB());
		fillData(workBook, sheetType, sheetNumber);
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

	private void fillData(WorkBook workBook, int sheetType, int sheetNumber) throws Exception {

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
				for (int j = 0; j < vtRow.size() - 2; j++) {
					SmartUtil.setNumber(workBook, row, 4 + j, (String) vtRow.get(j));
					SmartUtil.setStyleNumber(workBook, row, 4, row, 4 + j);
				}
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

		// fill data for table Weekly KPI statistics
		int rowHeader3 = allChart.getSeriesCount() * 3 + Util.ROW_AVG_START;
		workBook.setText(rowHeader3, 0, "III. Weekly KPI statistics");
		RangeStyle rangeStyle = workBook.getRangeStyle(rowHeader3, 0, rowHeader3, 0);
		SmartUtil.adjustFont(rangeStyle, 0, true, false, false);
		workBook.setRangeStyle(rangeStyle, rowHeader3, 0, rowHeader3, 0);

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
	 * 
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
	 * 
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
