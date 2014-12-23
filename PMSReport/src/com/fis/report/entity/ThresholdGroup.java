package com.fis.report.entity;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.List;


/**
 * The persistent class for the THRESHOLD_GROUP database table.
 * 
 */
public class ThresholdGroup implements Serializable {
	private static final long serialVersionUID = 1L;
	private long thresholdGroupId;
	private BigDecimal supplierId;
	private String thresholdGroupName;
	private List<Threshold> thresholds;
	private Resource resource;

	public ThresholdGroup() {
	}

	public long getThresholdGroupId() {
		return this.thresholdGroupId;
	}

	public void setThresholdGroupId(long thresholdGroupId) {
		this.thresholdGroupId = thresholdGroupId;
	}

	public BigDecimal getSupplierId() {
		return this.supplierId;
	}

	public void setSupplierId(BigDecimal supplierId) {
		this.supplierId = supplierId;
	}

	public String getThresholdGroupName() {
		return this.thresholdGroupName;
	}

	public void setThresholdGroupName(String thresholdGroupName) {
		this.thresholdGroupName = thresholdGroupName;
	}

	public List<Threshold> getThresholds() {
		return this.thresholds;
	}

	public void setThresholds(List<Threshold> thresholds) {
		this.thresholds = thresholds;
	}

	public Threshold addThreshold(Threshold threshold) {
		getThresholds().add(threshold);
		threshold.setThresholdGroup(this);
		return threshold;
	}

	public Threshold removeThreshold(Threshold threshold) {
		getThresholds().remove(threshold);
		threshold.setThresholdGroup(null);
		return threshold;
	}

	public Resource getResource() {
		return this.resource;
	}

	public void setResource(Resource resource) {
		this.resource = resource;
	}

}