package com.word.document.entity;

import java.io.Serializable;

public class GoodsCount implements Serializable{

	private static final long serialVersionUID = 1L;
	
	private String goodsName;
	private String goodsCount;
	public String getGoodsName() {
		return goodsName;
	}
	public void setGoodsName(String goodsName) {
		this.goodsName = goodsName;
	}
	public String getGoodsCount() {
		return goodsCount;
	}
	public void setGoodsCount(String goodsCount) {
		this.goodsCount = goodsCount;
	}
	@Override
	public String toString() {
		return "GoodsCount [goodsName=" + goodsName + ", goodsCount="
				+ goodsCount + "]";
	}
}
	
