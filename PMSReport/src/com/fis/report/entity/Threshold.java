package com.fis.report.entity;

import java.io.Serializable;
import java.math.BigDecimal;


/**
 * The persistent class for the THRESHOLD database table.
 * 
 */

public class Threshold implements Serializable {
	private static final long serialVersionUID = 1L;
	private long thresholdId;
	private BigDecimal alarmLevel;
	private BigDecimal maxValue;
	private BigDecimal minValue;
	private ThresholdGroup thresholdGroup;
	private ThresholdType thresholdType;

	public Threshold() {
	}

	public long getThresholdId() {
		return this.thresholdId;
	}

	public void setThresholdId(long thresholdId) {
		this.thresholdId = thresholdId;
	}

	public BigDecimal getAlarmLevel() {
		return this.alarmLevel;
	}

	public void setAlarmLevel(BigDecimal alarmLevel) {
		this.alarmLevel = alarmLevel;
	}

	public BigDecimal getMaxValue() {
		return this.maxValue;
	}

	public void setMaxValue(BigDecimal maxValue) {
		this.maxValue = maxValue;
	}

	public BigDecimal getMinValue() {
		return this.minValue;
	}

	public void setMinValue(BigDecimal minValue) {
		this.minValue = minValue;
	}

	public ThresholdGroup getThresholdGroup() {
		return this.thresholdGroup;
	}

	public void setThresholdGroup(ThresholdGroup thresholdGroup) {
		this.thresholdGroup = thresholdGroup;
	}

	public ThresholdType getThresholdType() {
		return this.thresholdType;
	}

	public void setThresholdType(ThresholdType thresholdType) {
		this.thresholdType = thresholdType;
	}

}