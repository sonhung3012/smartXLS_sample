package com.fis.report.entity;

import java.io.Serializable;

public class KPI implements Serializable{
	
	private static final long serialVersionUID = 4182628193326483518L;
	private int id;
	private String kpiCode;
	private String kpiName;
	private Resource resource;
	
	public KPI(int id, String kpiCode, String kpiName) {
		this.id = id;
		this.kpiCode = kpiCode.trim();
		this.kpiName = kpiName;

	}
	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getKpiCode() {
		return kpiCode;
	}

	public void setKpiCode(String kpi_code) {
		this.kpiCode = kpi_code;
	}

	public String getKpiName() {
		return kpiName;
	}

	public void setKpiName(String kpiName) {
		this.kpiName = kpiName;
	}

	public Resource getResource() {
		return resource;
	}

	public void setResource(Resource resource) {
		this.resource = resource;
	}



}
