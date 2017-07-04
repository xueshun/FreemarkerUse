package com.word.document;

import java.io.File;
import java.io.FileInputStream;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.sun.xml.internal.ws.util.StringUtils;
import com.word.document.entity.DateUtils;
import com.word.document.entity.ExcelUtils;
import com.word.document.entity.GoodsCount;
import com.word.document.entity.Order;



public class Main {

	private static final String Excel_Path="F:\\1111.xlsx";
	private static final List<String> gpsTypeList= new ArrayList<String>();

	public static void main(String[] args) throws Exception {
		//数据存储
		Map<String,Object> dataMap = new HashMap<String, Object>(); 
		List<Order> orderList = orderList();
		/*for (Order order : orderList) {
			System.out.println(order.getGoodsList().get(0).getGoodsCount().substring(0, 2));
		}*/
		List<Map<String,Object>> list1 = new ArrayList<Map<String,Object>>();
		List<Integer> list2 = new ArrayList<Integer>();
		dataMap.put("xytitle", "试卷");
		for (Order order : orderList) {
			Map<String,Object> map = new HashMap<String, Object>();
			map.put("companyName", order.getCompanyName()+"-"+order.getBranchName());
			map.put("acceptName", order.getGoodsAccept());
			map.put("acceptTel", order.getTel());
			map.put("goodCount", order.getGoodsList());
			map.put("acceptAddress", order.getAddress());
			map.put("applyTime", DateUtils.DateToString(order.getApplyTime(),"yyyy-MM-dd"));
			list1.add(map);
			list2.add(order.getGoodsList().size());
		
		}
		dataMap.put("table1", list1);
		dataMap.put("table2", list2);
		MDoc mdoc = new MDoc();
		mdoc.createDoc(dataMap, "F:\\test.doc");
	}

	/**
	 * 
	 * @return
	 * @throws Exception 
	 */
	@SuppressWarnings("unused")
	public static List<Order> orderList() throws Exception{
		Order order = null;
		FileInputStream input = new FileInputStream(new File("F:\\2017-6-30订单格式-模板.xlsx"));
		Workbook workBook = ExcelUtils.createWorkBook(input);
		Sheet sheet = workBook.getSheetAt(0);
		gpsTypeList.add("WYMINIS");
		gpsTypeList.add("WY900S");
		gpsTypeList.add("WY900S强磁");
		gpsTypeList.add("WY900S普通");
		gpsTypeList.add("WY500");
		gpsTypeList.add("WY600");
		//{"WYMINIS","WY900S","WY900S强磁","WY900S普通","WY500","WY600"};
		int gpsTypeCount=3;
		List<Order> orderList = new ArrayList<Order>();
		if(sheet != null){
			for (int i = 3; i <=8 ; i++) {
				if (sheet.getRow(3).getCell(i)==null) {
					break;
				}else{
					if(gpsTypeList.contains(sheet.getRow(3).getCell(i).toString().trim())){
						gpsTypeCount=gpsTypeCount+1;
					}
				}
			}
			
			
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

				for (int j = 3; j < gpsTypeCount; j++) {
					good = new GoodsCount();
					good.setGoodsName(sheet.getRow(3).getCell(j).toString().trim());
					String goodCount = sheet.getRow(i).getCell(j).toString().trim();
					if(org.apache.commons.lang3.StringUtils.isNotEmpty(goodCount)){
						System.out.println(Double.valueOf(goodCount));
						good.setGoodsCount(goodCount.substring(0, goodCount.indexOf(".")));
						goodsList.add(good);
					}else{
						good.setGoodsCount("");
					}
					
				}
				
				Cell cell = row.getCell(gpsTypeCount+1);
				Cell cell1 = row.getCell(gpsTypeCount);
				System.out.println();
				DecimalFormat df = new DecimalFormat("#");
				order.setGoodsList(goodsList);
				order.setGoodsAccept(cell1.getRichStringCellValue().toString());//收货人
				order.setTel(df.format(cell.getNumericCellValue()));//电话
				order.setAddress(row.getCell(gpsTypeCount+2).toString().trim());//收货地址
				orderList.add(order);
			}
		}
		return orderList;
	}
}
