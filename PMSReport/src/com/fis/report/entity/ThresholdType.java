package com.fis.report.entity;

import java.io.Serializable;
import java.util.List;


/**
 * The persistent class for the THRESHOLD_TYPE database table.
 * 
 */
public class ThresholdType implements Serializable {
	private static final long serialVersionUID = 1L;
	private long thresholdTypeId;
	private String thresholdClass;
	private String thresholdTypeName;
	private List<Threshold> thresholds;
	public ThresholdType() {
	}
	public long getThresholdTypeId() {
		return this.thresholdTypeId;
	}

	public void setThresholdTypeId(long thresholdTypeId) {
		this.thresholdTypeId = thresholdTypeId;
	}

	public String getThresholdClass() {
		return this.thresholdClass;
	}

	public void setThresholdClass(String thresholdClass) {
		this.thresholdClass = thresholdClass;
	}

	public String getThresholdTypeName() {
		return this.thresholdTypeName;
	}

	public void setThresholdTypeName(String thresholdTypeName) {
		this.thresholdTypeName = thresholdTypeName;
	}

	public List<Threshold> getThresholds() {
		return this.thresholds;
	}

	public void setThresholds(List<Threshold> thresholds) {
		this.thresholds = thresholds;
	}

	public Threshold addThreshold(Threshold threshold) {
		getThresholds().add(threshold);
		threshold.setThresholdType(this);

		return threshold;
	}

	public Threshold removeThreshold(Threshold threshold) {
		getThresholds().remove(threshold);
		threshold.setThresholdType(null);

		return threshold;
	}

}