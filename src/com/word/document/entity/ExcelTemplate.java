package com.word.document.entity;


import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelTemplate {

	/*public ExcelTemplate() throws Exception {
		
		
	}*/

	public List<Order> inputStream(InputStream inputStream) throws Exception{
		Workbook workBook = ExcelUtils.createWorkBook(inputStream);
		Order order = null;
		Sheet sheet = workBook.getSheetAt(0);
		List<Order> orderList = new ArrayList<Order>();
		if(sheet != null){
			for (int i = 4; i < sheet.getPhysicalNumberOfRows(); i++) {
				GoodsCount good = null;
				List<GoodsCount> goodsList = new ArrayList<GoodsCount>();
				Row row = sheet.getRow(i);
				if(row.getCell(0)==null ||row.getCell(0).toString()==""|| row.getCell(1)==null ||row.getCell(1).toString()==""
						||row.getCell(2).toString()==""){

					continue;
				}
				order = new Order();
				order.setCompanyName(sheet.getRow(0).getCell(0).toString().trim());//总公司名称
				order.setApplyTime(DateUtils.StringToDate(row.getCell(1).toString().trim(),"yyyy.MM.dd"));//申请时间
				order.setBranchName(row.getCell(2).toString().trim());//分公司名称

				if(row.getCell(3).toString()==null||row.getCell(3).toString()==""){
					order.setWyminis("");
				}else{
					order.setWyminis(row.getCell(3).toString().trim());//WYMINIS
					good = new GoodsCount();
					good.setGoodsName("WYMINIS");
					good.setGoodsCount(row.getCell(3).toString().trim());
					goodsList.add(good);
				}


				if(row.getCell(4).toString()==null||row.getCell(4).toString()==""){
					order.setWy900s("");
				}else{
					order.setWy900s(row.getCell(4).toString().trim());//WY900s
					good = new GoodsCount();
					good.setGoodsName("WY900S");
					good.setGoodsCount(row.getCell(4).toString().trim());
					goodsList.add(good);
				}


				if(row.getCell(5).toString()==null||row.getCell(5).toString()==""){
					order.setWy900sStrong("");
				}else{
					order.setWy900sStrong(row.getCell(5).toString().trim());//WY900S强磁
					good = new GoodsCount();
					good.setGoodsName("WY900S强磁");
					good.setGoodsCount(row.getCell(5).toString().trim());
					goodsList.add(good);
				}


				if(row.getCell(6).toString()==null||row.getCell(6).toString()==""){
					order.setWy900sOrdinary("");
				}else{
					order.setWy900sOrdinary(row.getCell(6).toString().trim());//WY900s普通
					good = new GoodsCount();
					good.setGoodsName("WY900S普通");
					good.setGoodsCount(row.getCell(6).toString().trim());
					goodsList.add(good);
				}


				if(row.getCell(7).toString()==null||row.getCell(7).toString()==""){
					order.setWy500("");
				}else{
					order.setWy500(row.getCell(7).toString().trim());//WY500
					good = new GoodsCount();
					good.setGoodsName("WY500");
					good.setGoodsCount(row.getCell(7).toString().trim());
					goodsList.add(good);
				}


				if(row.getCell(8).toString()==null||row.getCell(8).toString()==""){
					order.setWy600("");
				}else{
					order.setWy600(row.getCell(8).toString().trim());//WY600
					good = new GoodsCount();
					good.setGoodsName("WY600");
					good.setGoodsCount(row.getCell(8).toString().trim());
					goodsList.add(good);
				}
				Cell cell = row.getCell(10);
				DecimalFormat df = new DecimalFormat("#");
				order.setGoodsList(goodsList);
				order.setGoodsAccept(row.getCell(9).toString().trim());//收货人
				order.setTel(df.format(cell.getNumericCellValue()));//电话
				order.setAddress(row.getCell(11).toString().trim());//收货地址
				System.out.println(order);
				orderList.add(order);
			}
		}
		return orderList;
	}
}
