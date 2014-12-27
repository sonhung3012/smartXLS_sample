package com.fis.report.data;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.fis.report.entity.ReportKPI;
import com.fss.sql.Database;

/**
 * 
 * @author THINHNV
 * 
 */
public class ReportConfig {

	private Connection mcnMain;
	private PreparedStatement pstmt = null;
	private ResultSet rs = null;

	public static List<ReportKPI> lstReportKpi = new ArrayList<ReportKPI>();
	public static Map<String, String> mapKpiName = new HashMap<String, String>();

	public ReportConfig(Connection mcnMain) throws SQLException {

		this.mcnMain = mcnMain;
	}

	public void configLoader() throws SQLException {

		getReportKPI();
	}

	private void getReportKPI() throws SQLException {

		try {
			pstmt = mcnMain.prepareStatement("select id,sheet_id,kpi_code,kpi_name,kpi_order from report_kpi");
			rs = pstmt.executeQuery();
			while (rs.next()) {
				int sheetId = rs.getInt("sheet_id");
				String kpiCode = rs.getString("kpi_code");
				String kpiName = rs.getString("kpi_name");
				int kpiOrder = rs.getInt("kpi_order");
				ReportKPI rpKPI = new ReportKPI(sheetId, kpiCode, kpiName, kpiOrder);
				lstReportKpi.add(rpKPI);
				mapKpiName.put(kpiCode, kpiName);
			}
		} finally {
			Database.closeObject(pstmt);
			Database.closeObject(rs);
		}
	}

	/**
	 * 
	 * @param sheetId
	 * @return
	 */
	public static int getNumOfKPI(int sheetId) {

		int count = 0;
		for (ReportKPI rp : lstReportKpi) {
			if (sheetId == rp.getSheetId()) {
				count++;
			}
		}
		return count;
	}
	// public Threshold getThresholdNomal(String kpiCode){
	// Resource resource = getResource(kpiCode);
	// ThresholdGroup[] thresholdGroup = resource.getThresholdGroup();
	// for (ThresholdGroup tg : thresholdGroup) {
	// List<Threshold> thresholdList = tg.getThresholds();
	// for (Threshold th : thresholdList) {
	// if(th.getThresholdType().equals("4")){
	// return th;
	// }
	// }
	// }
	// return null;
	// }
	//
	// private Resource getResource(String kpiCode){
	// for(KPI kpi : lstKPI){
	// if(kpiCode.equals(kpi.getKpiCode())){
	// return kpi.getResource();
	// }
	// }
	// return null;
	// }
	// private void getResource() throws SQLException {
	// try {
	// pstmt = mcnMain.prepareStatement("select id,sheet_id,kpi_code,kpi_name,kpi_order from report_kpi");
	// rs = pstmt.executeQuery();
	// while (rs.next()) {
	// int sheetId = rs.getInt("sheet_id");
	// String kpiCode = rs.getString("kpi_code");
	// String kpiName = rs.getString("kpi_name");
	// int kpiOrder = rs.getInt("kpi_order");
	// ReportKPI rpKPI = new ReportKPI(sheetId, kpiCode, kpiName, kpiOrder);
	// lstReportKpi.add(rpKPI);
	// }
	// } finally {
	// Database.closeObject(pstmt);
	// Database.closeObject(rs);
	// }
	// }
	//
	// private void getKpi() throws SQLException {
	// try {
	// pstmt = mcnMain.prepareStatement("select id,sheet_id,kpi_code,kpi_name,kpi_order from report_kpi");
	// rs = pstmt.executeQuery();
	// while (rs.next()) {
	// int sheetId = rs.getInt("sheet_id");
	// String kpiCode = rs.getString("kpi_code");
	// String kpiName = rs.getString("kpi_name");
	// int kpiOrder = rs.getInt("kpi_order");
	// ReportKPI rpKPI = new ReportKPI(sheetId, kpiCode, kpiName, kpiOrder);
	// lstReportKpi.add(rpKPI);
	// }
	// } finally {
	// Database.closeObject(pstmt);
	// Database.closeObject(rs);
	// }
	// }
	//
	// private void getThreshold() throws SQLException {
	// try {
	// pstmt = mcnMain.prepareStatement("select id,sheet_id,kpi_code,kpi_name,kpi_order from report_kpi");
	// rs = pstmt.executeQuery();
	// while (rs.next()) {
	// int sheetId = rs.getInt("sheet_id");
	// String kpiCode = rs.getString("kpi_code");
	// String kpiName = rs.getString("kpi_name");
	// int kpiOrder = rs.getInt("kpi_order");
	// ReportKPI rpKPI = new ReportKPI(sheetId, kpiCode, kpiName, kpiOrder);
	// lstReportKpi.add(rpKPI);
	// }
	// } finally {
	// Database.closeObject(pstmt);
	// Database.closeObject(rs);
	// }
	// }
	//
	// private void getThresholdGroup() throws SQLException {
	// try {
	// pstmt = mcnMain.prepareStatement("select id,sheet_id,kpi_code,kpi_name,kpi_order from report_kpi");
	// rs = pstmt.executeQuery();
	// while (rs.next()) {
	// int sheetId = rs.getInt("sheet_id");
	// String kpiCode = rs.getString("kpi_code");
	// String kpiName = rs.getString("kpi_name");
	// int kpiOrder = rs.getInt("kpi_order");
	// ReportKPI rpKPI = new ReportKPI(sheetId, kpiCode, kpiName, kpiOrder);
	// lstReportKpi.add(rpKPI);
	// }
	// } finally {
	// Database.closeObject(pstmt);
	// Database.closeObject(rs);
	// }
	// }

}