package com.fis.report.common;

import java.util.Calendar;
import java.util.Collections;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Vector;

import com.fis.report.entity.ReportKPI;

/**
 * 
 * @author THINHNV
 * 
 */

public class Util {

	public final static String WEEK_1 = "1";
	public final static String WEEK_2 = "2";
	public final static String WEEK_3 = "3";
	public final static String WEEK_4 = "4";

	public final static int QOS_KPIS_NUM = 0;
	public final static int TRAFFIC_NUM = 1;
	public final static int UTILIZATION_NUM = 2;
	public final static int QOS_KPIS_TYPE = 1;
	public final static int TRAFFIC_TYPE = 2;
	public final static int UTILIZATION_TYPE = 3;

	public final static int TYPE_CODE_NUM_OF_CELL = 1;
	public final static int TYPE_CODE_BY_ALARM = 2;
	public final static int TYPE_CODE_AVG = 3;

	public final static int ROW_AVG_START = 76;
	public final static int ROW_AVG_NEXT = 3;
	public final static int ROW_HOUR_START = 148;
	public final static int ROW_KPI_NEXT = 19;
	public final static int ROW_STATIC_START = 153;
	public final static int ROW_STATIC_NEXT = 6;
	public final static int ROW_AREA_START = 569;
	public final static int ROW_AREA_NEXT = 23;

	public final static int COLUMN_HOUR_FROM = 4;
	public final static int COLUMN_STATIC_FROM = 4;
	public final static int COLUMN_AREA_FROM = 2;

	public final static String HUAWEI_TYPE = "1";
	public final static String ALCATEL_TYPE = "2";

	public final static int ALL_NETWORK_CHART = 0;
	public final static int ALCATEL_CHART = 1;
	public final static int HUAWEI_CHART = 2;

	public static Map<String, Integer> mapPosition = null;

	private static String[] monthNames = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };

	/**
	 * 
	 * @param lstKPI
	 * @param sheetType
	 * @param startRow
	 * @param nextRow
	 * @return
	 */
	public static void putPosition(List<ReportKPI> lstKPI, int sheetType, int startRow, int nextRow) {

		mapPosition = new HashMap<String, Integer>();
		Collections.sort(lstKPI);
		int rPosition = startRow;
		for (ReportKPI rpKPI : lstKPI) {
			if (rpKPI.getSheetId() == sheetType) {
				mapPosition.put(rpKPI.getKpiCode() + sheetType, rPosition);
				rPosition += nextRow;
			}
		}
	}

	/**
	 * 
	 * @param mapPosition
	 * @param kpiCode
	 * @return
	 * @throws Exception
	 */
	public static Integer getPosition(String kpiCode, int sheetType) {

		if (mapPosition.get(kpiCode + sheetType) != null)
			return mapPosition.get(kpiCode + sheetType);
		return null;
	}

	public static String getAverageStringNumber(Vector<String> list) {

		Double sum = 0d;
		int count = 0;
		for (int i = 0; i < 7; i++) {

			if (list.get(i) != null) {
				Double eleDouble = Double.parseDouble(list.get(i));
				sum += eleDouble;
				count++;
			}
		}

		Double average = sum / count;

		return String.valueOf(average);
	}

	public static String getMonthNames(String strDate) {

		return monthNames[Integer.parseInt(strDate.substring(strDate.indexOf("/") + 1, strDate.lastIndexOf("/"))) - 1];
	}

	public static int getDaysInMonth(String strDate) {

		int day = Integer.parseInt(strDate.substring(0, strDate.indexOf("/")));
		int month = Integer.parseInt(strDate.substring(strDate.indexOf("/") + 1, strDate.lastIndexOf("/"))) - 1;
		int year = Integer.parseInt(strDate.substring(strDate.lastIndexOf("/") + 1));
		GregorianCalendar calendar = new GregorianCalendar(year, month, day);
		int daysInMonth = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);

		return daysInMonth;
	}
}