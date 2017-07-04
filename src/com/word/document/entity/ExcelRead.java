package com.word.document.entity;

import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;


public class ExcelRead {

	
	private static final List<String> gpsTypeList= new ArrayList<String>();
	@SuppressWarnings("null")
	public void excelRead() throws Exception{
		Order order = null;
		FileInputStream input = new FileInputStream(new File("F:\\2017-6-30订单格式-模板.xlsx"));
		Workbook workBook = ExcelUtils.createWorkBook(input);
		Sheet sheet = workBook.getSheetAt(0);
		List<Order> orderList = new ArrayList<Order>();
		gpsTypeList.add("WYMINIS");
		gpsTypeList.add("WY900S");
		gpsTypeList.add("WY900S强磁");
		gpsTypeList.add("WY900S普通");
		gpsTypeList.add("WY500");
		gpsTypeList.add("WY600");
		int gpsTypeCount=3;
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
				
				List<GpsType> gpsTypeList = new ArrayList<GpsType>();
				GpsType gpsType = new GpsType();
				for (int j = 3; j < gpsTypeCount; j++) {
					if(row.getCell(j).toString()==null||row.getCell(j).toString()==""){
						continue;
					}else{
						good = new GoodsCount();
						String gpsTypeName = sheet.getRow(3).getCell(j).toString().trim();
						gpsType.setGpsTypeName(gpsTypeName);//设备名称
						gpsType.setGpsTypeCount(Double.valueOf(sheet.getRow(i).getCell(j).toString().trim()));//设备数量
						good.setGoodsName(gpsTypeName);
						good.setGoodsCount(sheet.getRow(i).getCell(j).toString().trim());
					}
					gpsTypeList.add(gpsType);
					goodsList.add(good);
				}
				/*
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
				}*/
				Cell cell = row.getCell(gpsTypeCount+1);
				DecimalFormat df = new DecimalFormat("#");
				order.setGoodsList(goodsList);
				order.setGoodsAccept(row.getCell(gpsTypeCount).toString().trim());//收货人
				order.setTel(df.format(cell.getNumericCellValue()));//电话
				order.setAddress(row.getCell(gpsTypeCount+2).toString().trim());//收货地址
				orderList.add(order);
				WriteWord(orderList);
			}
		}

	}

	static void WriteWord(List<Order> orderList) throws IOException{

		for (Order order : orderList) {
			System.out.println(order);
			/*Map<String, Object> params = new HashMap<String, Object>();
			params.put("companyName", order.getCompanyName());
			params.put("acceptName", order.getGoodsAccept());
			params.put("acceptTel", order.getTel());
			params.put("acceptAddress", order.getAddress());
			for (GoodsCount good : order.getGoodsList()) {
				params.put("goodsName1",good.getGoodsName() );
				params.put("goodsName2", good.getGoodsName());
				params.put("count1", good.getGoodsCount());
				params.put("count2",good.getGoodsCount());
			}
			params.put("applyTime",order.getApplyTime());
			String filePath = "F:\\order_import.docx";  
			InputStream is = new FileInputStream(filePath);  
			XWPFDocument doc = new XWPFDocument(is);  
			//替换段落里面的变量  
			ExcelRead.replaceInPara(doc, params);  
			//替换表格里面的变量  
			ExcelRead.replaceInTable(doc, params);
			OutputStream os = new FileOutputStream("F:\\write"+order.getBranchName()+".docx");  
			doc.write(os); 
			ExcelRead.close(is);
			ExcelRead.close(os);*/
		}

	}

	public static void main(String[] args) throws Exception {
		ExcelRead excelRead = new ExcelRead();
		excelRead.excelRead();
	}

	/** 
	 * 替换段落里面的变量 
	 * @param doc 要替换的文档 
	 * @param params 参数 
	 */  
	private static void replaceInPara(XWPFDocument doc, Map<String, Object> params) {  
		Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();  
		XWPFParagraph para;  
		while (iterator.hasNext()) {  
			para = iterator.next();  
			ExcelRead.replaceInPara(para, params);  
		}  
	}  

	/** 
	 * 替换段落里面的变量 
	 * @param para 要替换的段落 
	 * @param params 参数 
	 */  
	private static void replaceInPara(XWPFParagraph para, Map<String, Object> params) {  
		List<XWPFRun> runs;  
		Matcher matcher;  
		if (ExcelRead.matcher(para.getParagraphText()).find()) {  
			runs = para.getRuns();  

			int start = -1;  
			int end = -1;  
			String str = "";  
			for (int i = 0; i < runs.size(); i++) {  
				XWPFRun run = runs.get(i);  
				String runText = run.toString();  
				System.out.println("------>>>>>>>>>" + runText);  
				if ('$' == runText.charAt(0)&&'{' == runText.charAt(1)) {  
					start = i;  
				}  
				if ((start != -1)) {  
					str += runText;  
				}  
				if ('}' == runText.charAt(runText.length() - 1)) {  
					if (start != -1) {  
						end = i;  
						break;  
					}  
				}  
			}  
			System.out.println("start--->"+start);  
			System.out.println("end--->"+end);  

			System.out.println("str---->>>" + str);  

			for (int i = start; i <= end; i++) {  
				para.removeRun(i);  
				i--;  
				end--;  
				System.out.println("remove i="+i);  
			}  

			for (String key : params.keySet()) {  
				if (str.equals(key)) { 
					System.out.println(params.get(key));
					para.createRun().setText((String) params.get(key));  
					break;  
				}  
			}  


		}  
	}  

	/** 
	 * 替换表格里面的变量 
	 * @param doc 要替换的文档 
	 * @param params 参数 
	 */  
	private static void replaceInTable(XWPFDocument doc, Map<String, Object> params) {  
		Iterator<XWPFTable> iterator = doc.getTablesIterator();  
		XWPFTable table;  
		List<XWPFTableRow> rows;  
		List<XWPFTableCell> cells;  
		List<XWPFParagraph> paras;  
		while (iterator.hasNext()) {  
			table = iterator.next();  
			rows = table.getRows();  
			for (XWPFTableRow row : rows) {  
				cells = row.getTableCells();  
				for (XWPFTableCell cell : cells) {  
					paras = cell.getParagraphs();  
					for (XWPFParagraph para : paras) {  
						ExcelRead.replaceInPara(para, params);  
					}  
				}  
			}  
		}  
	}  

	/** 
	 * 正则匹配字符串 
	 * @param str 
	 * @return 
	 */  
	private static Matcher matcher(String str) {  
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);  
		Matcher matcher = pattern.matcher(str);  
		return matcher;  
	}  

	/** 
	 * 关闭输入流 
	 * @param is 
	 */  
	private static void close(InputStream is) {  
		if (is != null) {  
			try {  
				is.close();  
			} catch (IOException e) {  
				e.printStackTrace();  
			}  
		}  
	}  

	/** 
	 * 关闭输出流 
	 * @param os 
	 */  
	private static void close(OutputStream os) {  
		if (os != null) {  
			try {  
				os.close();  
			} catch (IOException e) {  
				e.printStackTrace();  
			}  
		}  
	}
}
