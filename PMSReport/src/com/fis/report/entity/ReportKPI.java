package com.fis.report.entity;

public class ReportKPI implements Comparable<ReportKPI> {
    private String id;
    private int sheetId;
    private String kpiCode;
    private String kpiName;
    private int kpiOrder;
    
    public ReportKPI(int sheetId,String kpiCode, String kpiName,int kpiOrder){
	this.sheetId = sheetId;
	this.kpiCode = kpiCode;
	this.kpiName = kpiName;
	this.kpiOrder = kpiOrder;
    }
    
    public String getId() {
        return id;
    }
    public void setId(String id) {
        this.id = id;
    }
    public int getSheetId() {
        return sheetId;
    }
    public void setSheetId(int sheetId) {
        this.sheetId = sheetId;
    }
    public String getKpiCode() {
        return kpiCode;
    }
    public void setKpiCode(String kpiCode) {
        this.kpiCode = kpiCode;
    }
    public String getKpiName() {
        return kpiName;
    }
    public void setKpiName(String kpiName) {
        this.kpiName = kpiName;
    }
    public int getKpiOrder() {
        return kpiOrder;
    }
    public void setKpiOrder(int kpiOrder) {
        this.kpiOrder = kpiOrder;
    }

	@Override
	public int compareTo(ReportKPI report) {
		if(this.sheetId == report.sheetId && this.kpiOrder >= report.kpiOrder)
			return 1;
		return -1;
	}
}
