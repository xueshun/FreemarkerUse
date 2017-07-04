package com.word.document.entity;

import java.io.Serializable;
import java.util.Date;
import java.util.List;

public class Order implements Serializable{

	private static final long serialVersionUID = 1L;
    private String companyName;
    private Date applyTime;
    private String branchName;
    private String wyminis;
    private String wy900s;
    private String wy900sStrong;
    private String wy900sOrdinary;
    private String wy500;
    private String wy600;
    private String goodsAccept;
    private String tel;
    private String address;
    private List<GoodsCount> goodsList;
    
    
    
	public String getTel() {
		return tel;
	}
	public void setTel(String tel) {
		this.tel = tel;
	}
	public List<GoodsCount> getGoodsList() {
		return goodsList;
	}
	public void setGoodsList(List<GoodsCount> goodsList) {
		this.goodsList = goodsList;
	}
	public String getCompanyName() {
		return companyName;
	}
	public void setCompanyName(String companyName) {
		this.companyName = companyName;
	}
	
	public Date getApplyTime() {
		return applyTime;
	}
	public void setApplyTime(Date applyTime) {
		this.applyTime = applyTime;
	}
	public String getBranchName() {
		return branchName;
	}
	public void setBranchName(String branchName) {
		this.branchName = branchName;
	}
	public String getWyminis() {
		return wyminis;
	}
	public void setWyminis(String wyminis) {
		this.wyminis = wyminis;
	}
	public String getWy900s() {
		return wy900s;
	}
	public void setWy900s(String wy900s) {
		this.wy900s = wy900s;
	}
	public String getWy900sStrong() {
		return wy900sStrong;
	}
	public void setWy900sStrong(String wy900sStrong) {
		this.wy900sStrong = wy900sStrong;
	}
	public String getWy900sOrdinary() {
		return wy900sOrdinary;
	}
	public void setWy900sOrdinary(String wy900sOrdinary) {
		this.wy900sOrdinary = wy900sOrdinary;
	}
	public String getWy500() {
		return wy500;
	}
	public void setWy500(String wy500) {
		this.wy500 = wy500;
	}
	public String getWy600() {
		return wy600;
	}
	public void setWy600(String wy600) {
		this.wy600 = wy600;
	}
	public String getGoodsAccept() {
		return goodsAccept;
	}
	public void setGoodsAccept(String goodsAccept) {
		this.goodsAccept = goodsAccept;
	}
	
	
	
	
	public String getAddress() {
		return address;
	}
	public void setAddress(String address) {
		this.address = address;
	}
	@Override
	public String toString() {
		return "Order [companyName=" + companyName + ", applyTime=" + applyTime
				+ ", branchName=" + branchName + ", wyminis=" + wyminis
				+ ", wy900s=" + wy900s + ", wy900sStrong=" + wy900sStrong
				+ ", wy900sOrdinary=" + wy900sOrdinary + ", wy500=" + wy500
				+ ", wy600=" + wy600 + ", goodsAccept=" + goodsAccept
				+ ", tel=" + tel + ", address=" + address + ", goodsList="
				+ goodsList + "]";
	}
}
