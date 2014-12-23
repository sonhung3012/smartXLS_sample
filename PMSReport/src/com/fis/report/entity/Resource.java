package com.fis.report.entity;
import java.io.Serializable;

/**
 * <p>
 * Title:
 * </p>
 * 
 * <p>
 * Description:
 * </p>
 * 
 * <p>
 * Copyright: Copyright (c) 2004
 * </p>
 * 
 * <p>
 * Company:
 * </p>
 * 
 * @author not attributable
 * @version 1.0
 */
public class Resource implements Serializable {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private int id;
	private String code;
	private String name;
	private String description;
	private ThresholdGroup[] thresholdGroup;
	private KPI kpi;

	public Resource(int id,String code, String name, String description) {
		this.id = id;
		this.code = code;
		this.name = name;
		this.description = description;
	}
	
	public Resource(int id, String name, String description) {
		this.id = id;
		this.name = name;
		this.description = description;
	}
	public Resource(int id, String name) {
		this.id = id;
		this.name = name;
	}

	public String getdescription() {
		return description;
	}

	public void setdescription(String description) {
		this.description = description;
	}

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}
	
	public String getCode() {
		return code;
	}

	public void setCode(String code) {
		this.code = code;
	}
	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
	
	public ThresholdGroup[] getThresholdGroup() {
		return thresholdGroup;
	}

	public void setThresholdGroup(ThresholdGroup[] thresholdGroup) {
		this.thresholdGroup = thresholdGroup;
	}

	public KPI getKpi() {
		return kpi;
	}

	public void setKpi(KPI kpi) {
		this.kpi = kpi;
	}

	public Resource duplicate(String description) {
		return new Resource(getId(),getCode(), getName(), description);
	}

	public String toString() {
		return "id=" + id + ",name=" + name;
	}
}
