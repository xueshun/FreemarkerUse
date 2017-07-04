package com.word.document.entity;

import java.io.Serializable;

public class GpsType implements Serializable{

	private static final long serialVersionUID = 1L;
	private String gpsTypeName;
	private Double gpsTypeCount;
	public String getGpsTypeName() {
		return gpsTypeName;
	}
	public void setGpsTypeName(String gpsTypeName) {
		this.gpsTypeName = gpsTypeName;
	}
	public Double getGpsTypeCount() {
		return gpsTypeCount;
	}
	public void setGpsTypeCount(Double gpsTypeCount) {
		this.gpsTypeCount = gpsTypeCount;
	}
	@Override
	public String toString() {
		return "GpsType [gpsTypeName=" + gpsTypeName + ", gpsTypeCount=" + gpsTypeCount + "]";
	}
	
}
