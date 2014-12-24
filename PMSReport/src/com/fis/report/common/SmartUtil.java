package com.fis.report.common;

import java.awt.Color;

import com.smartxls.ChartFormat;
import com.smartxls.ChartShape;
import com.smartxls.RangeStyle;
import com.smartxls.WorkBook;

/**
 * 
 * @author THINHNV
 * 
 */
public class SmartUtil {

	/**
	 * 
	 * @param columnNumber
	 * @return name of column in excel
	 */
	public static String getColumnName(int columnNumber) {

		String columnName = "";
		int dividend = columnNumber + 1;
		int modulus;
		while (dividend > 0) {
			modulus = (dividend - 1) % 26;
			columnName = (char) (65 + modulus) + columnName;
			dividend = (int) ((dividend - modulus) / 26);
		}
		return columnName;
	}

	/**
	 * 
	 * @param sheetName
	 * @param row
	 * @param fromCol
	 * @param toCol
	 * @return
	 */
	public static String getFormula(String sheetName, int fromRow, int fromCol, int toRow, int toCol) {

		// workBook.getSheetName(0)+ "!$E$"+ (row) +":$K$" + (row);
		return sheetName + "!$" + getColumnName(fromCol) + "$" + fromRow + ":$" + getColumnName(toCol) + "$" + toRow;
	}

	/**
	 * 
	 * @param workBook
	 * @param start_x
	 * @param start_y
	 * @param end_x
	 * @param end_y
	 * @param borderStyle
	 *            : RangeStyle.BorderThick ....
	 * @throws Exception
	 */
	public static void setFullBorder(WorkBook workBook, int start_x, int start_y, int end_x, int end_y, short borderStyle) throws Exception {

		RangeStyle rangeStyle = workBook.getRangeStyle(start_x, start_y, end_x, end_y);
		rangeStyle.setLeftBorder(borderStyle);
		rangeStyle.setTopBorder(borderStyle);
		rangeStyle.setRightBorder(borderStyle);
		rangeStyle.setBottomBorder(borderStyle);
		// apply to the horizon border of the whole range
		rangeStyle.setHorizontalInsideBorder(borderStyle);
		// apply to the vertical border of the whole range
		rangeStyle.setVerticalInsideBorder(borderStyle);
		workBook.setRangeStyle(rangeStyle, start_x, start_y, end_x, end_y);
	}

	/**
	 * 
	 * @param workBook
	 * @param start_x
	 * @param start_y
	 * @param end_x
	 * @param end_y
	 * @param isMerge
	 *            = true -> merge: = false -> UnMerge
	 * @throws Exception
	 */
	public static void mergeCells(WorkBook workBook, int start_x, int start_y, int end_x, int end_y, Boolean isMerge) throws Exception {

		RangeStyle rangeStyle = workBook.getRangeStyle(start_x, start_y, end_x, end_y);
		rangeStyle.setMergeCells(isMerge);
		workBook.setRangeStyle(rangeStyle, start_x, start_y, end_x, end_y);
	}

	/**
	 * 
	 * @param workBook
	 * @param start_x
	 * @param start_y
	 * @param end_x
	 * @param end_y
	 * @param color
	 * @throws Exception
	 */
	public static void fillColorPatten(WorkBook workBook, int start_x, int start_y, int end_x, int end_y, int color) throws Exception {

		RangeStyle rangeStyle = workBook.getRangeStyle(start_x, start_y, end_x, end_y);
		rangeStyle.setPattern(RangeStyle.PatternSolid);
		rangeStyle.setPatternFG(color);
		workBook.setRangeStyle(rangeStyle, start_x, start_y, end_x, end_y);
	}

	/**
	 * 
	 * @param workBook
	 * @param start_x
	 * @param start_y
	 * @param end_x
	 * @param end_y
	 * @param fontName
	 * @param fontSize
	 * @throws Exception
	 */
	public static void setFonts(WorkBook workBook, int start_x, int start_y, int end_x, int end_y, String fontName, int fontSize) throws Exception {

		RangeStyle rangeStyle = workBook.getRangeStyle(start_x, start_y, end_x, end_y);// get format from range B2:C3
		rangeStyle.setFontName(fontName);
		rangeStyle.setFontSize(fontSize);
		workBook.setRangeStyle(rangeStyle, start_x, start_y, end_x, end_y);// set format for range B2:C3
	}

	public static void adjustFont(RangeStyle rangeStyle, int color, boolean bold, boolean italic, boolean underline) {

		rangeStyle.setFontBold(bold);
		rangeStyle.setFontItalic(italic);
		if (underline) {
			rangeStyle.setFontUnderline(RangeStyle.UnderlineSingle);
		}
		rangeStyle.setFontColor(color);
	}

	public static void paintBorder(WorkBook workBook, int start_x, int start_y, int end_x, int end_y) throws Exception {

		RangeStyle rangeStyle = workBook.getRangeStyle(start_x, start_y, end_x, end_y);
		rangeStyle.setTopBorder(RangeStyle.BorderThin);
		rangeStyle.setBottomBorder(RangeStyle.BorderThin);
		rangeStyle.setLeftBorder(RangeStyle.BorderThin);
		rangeStyle.setRightBorder(RangeStyle.BorderThin);
		workBook.setRangeStyle(rangeStyle, start_x, start_y, end_x, end_y);

	}

	/**
	 * 
	 * @param workBook
	 * @param start_x
	 * @param start_y
	 * @param end_x
	 * @param end_y
	 * @throws Exception
	 */
	public static void setDefaultStyle(WorkBook workBook, int start_x, int start_y, int end_x, int end_y) throws Exception {

		RangeStyle rangeStyle = workBook.getRangeStyle(start_x, start_y, end_x, end_y);// get format from range B2:C3
		rangeStyle.setFontName("Times New Roman");
		rangeStyle.setFontSize(9 * 20);
		rangeStyle.setFontBold(true);
		rangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentCenter);
		workBook.setRangeStyle(rangeStyle, start_x, start_y, end_x, end_y);// set format for range B2:C3
	}

	/**
	 * 
	 * @param workBook
	 * @param start_x
	 * @param start_y
	 * @param end_x
	 * @param end_y
	 * @throws Exception
	 */
	public static void setStyleString(WorkBook workBook, int start_x, int start_y, int end_x, int end_y) throws Exception {

		RangeStyle rangeStyle = workBook.getRangeStyle(start_x, start_y, end_x, end_y);// get format from range B2:C3
		rangeStyle.setFontName("Times New Roman");
		rangeStyle.setFontSize(9 * 20);
		rangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentLeft);
		rangeStyle.setVerticalAlignment(RangeStyle.VerticalAlignmentCenter);
		workBook.setRangeStyle(rangeStyle, start_x, start_y, end_x, end_y);// set format for range B2:C3
	}

	/**
	 * 
	 * @param workBook
	 * @param start_x
	 * @param start_y
	 * @param end_x
	 * @param end_y
	 * @throws Exception
	 */
	public static void setStyleNumber(WorkBook workBook, int start_x, int start_y, int end_x, int end_y) throws Exception {

		RangeStyle rangeStyle = workBook.getRangeStyle(start_x, start_y, end_x, end_y);// get format from range B2:C3
		rangeStyle.setFontName("Times New Roman");
		rangeStyle.setFontSize(9 * 20);
		rangeStyle.setHorizontalAlignment(RangeStyle.HorizontalAlignmentRight);
		rangeStyle.setCustomFormat("0.00");
		workBook.setRangeStyle(rangeStyle, start_x, start_y, end_x, end_y);// set format for range B2:C3
	}

	/**
	 * 
	 * @param workBook
	 * @param x
	 * @param y
	 * @param value
	 * @throws NumberFormatException
	 * @throws Exception
	 */
	public static void setNumber(WorkBook workBook, int x, int y, String value) throws NumberFormatException, Exception {

		if (value == null || value.equals("")) {
			return;
		}
		workBook.setNumber(x, y, Double.parseDouble(value));
	}

	/**
	 * 
	 * @param chart
	 * @param name
	 * @param value
	 * @param lineNumber
	 * @throws Exception
	 */
	public static void addSeries(ChartShape chart, String name, String formula, int lineNumber) throws Exception {

		chart.addSeries();
		setLineFormat(chart, lineNumber);
		chart.setSeriesName(lineNumber, name);
		chart.setSeriesYValueFormula(lineNumber, formula);
		chart.setSeriesSmoothedLine(lineNumber, true);
	}

	/**
	 * 
	 * @param chart
	 * @throws Exception
	 */
	public static void setLegendFormat(ChartShape chart) throws Exception {

		ChartFormat format = chart.getLegendFormat();
		format.setForeColor(Color.RED.getRGB());
		chart.setLegendFormat(format);
	}

	/**
	 * 
	 * @param chart
	 * @throws Exception
	 */
	public static void setPilotFormat(ChartShape chart) throws Exception {

		ChartFormat format = chart.getPlotFormat();
		format.setBackColor(Color.BLUE.getRGB());
		chart.setPlotFormat(format);
	}

	/**
	 * 
	 * @param chart
	 * @param lineNumber
	 * @throws Exception
	 */
	public static void setLineFormat(ChartShape chart, int lineNumber) throws Exception {

		ChartFormat format = chart.getSeriesFormat(lineNumber);
		format.setLineWeight(30);
		chart.setSeriesFormat(lineNumber, format);
	}

	/**
	 * Test
	 * 
	 * @param args
	 */
	public static void main(String[] args) {

		System.out.println("Get column name: " + getFormula("abc", 76, 1, 76, 7));
	}
}
