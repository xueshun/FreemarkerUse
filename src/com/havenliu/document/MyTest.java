package com.havenliu.document;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class MyTest {

	private Configuration configuration = null;

	public MyTest() {
		configuration = new Configuration();
		configuration.setDefaultEncoding("utf-8");
	}

	@Test
	public void createDoc() {
		//要填入模本的数据文件
		Map<String,Object> dataMap=new HashMap<String,Object>();
		getData(dataMap);
		//设置模本装置方法和路径,FreeMarker支持多种模板装载方法。可以重servlet，classpath，数据库装载，
		//这里我们的模板是放在com.havenliu.document.template包下面
		configuration.setClassForTemplateLoading(this.getClass(), "/com/havenliu/document/template");
		Template t=null;
		try {
			//test.ftl为要装载的模板
			t = configuration.getTemplate("fctestpaper.ftl");
		} catch (IOException e) {
			e.printStackTrace();
		}
		//输出文档路径及名称
		//File outFile = new File("D:/outFilessa"+Math.random()*10000+".doc");
		File outFile = new File("F:/outFilessa.doc");
		Writer out = null;
		try {
			out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile)));
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}
		 
        try {
			t.process(dataMap, out);
		} catch (TemplateException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 注意dataMap里存放的数据Key值要与模板中的参数相对应
	 * @param dataMap
	 */
	private void getData(Map<String, Object> dataMap) {  
//        dataMap.put("title", "标题");  
//        dataMap.put("nian", "2012");  
//        dataMap.put("yue", "2");  
//        dataMap.put("ri", "13");  
//        dataMap.put("bianzhi", "唐鑫");  
//        dataMap.put("dianhua", "13020265912");  
//        dataMap.put("shenheren", "占文涛");  
//      dataMap.put("number", 1);  
//      dataMap.put("content", "内容"+2);  
        
		dataMap.put("xytitle", "期末考试"); 
		
        List<Map<String,Object>> list1 = new ArrayList<Map<String,Object>>();  
        for (int i = 0; i < 10; i++) {  
            Map<String,Object> map = new HashMap<String,Object>();  
            map.put("xzn", i);  
            map.put("xztest", "(   )操作系统允许在一台主机上同时连接多台终端，多个用户可以通过各自的终端同时交互地使用计算机。"+i);
            map.put("ans1", "A.网络"+i);
            map.put("ans2", "B.分布式"+i);
            map.put("ans3", "C.分时"+i);
            map.put("ans4", "D.实时"+i);
            list1.add(map);  
        }  
          
        dataMap.put("table1", list1); 

        List<Map<String,Object>> list2 = new ArrayList<Map<String,Object>>();  
        for (int i = 0; i < 10; i++) {  
            Map<String,Object> map = new HashMap<String,Object>();  
            map.put("tkn", i);  
            map.put("tktest", "操作系统是计算机系统中的一个___系统软件_______，它管理和控制计算机系统中的___资源_________."+i);  
            list2.add(map);  
        }  
          
        dataMap.put("table2", list2);
        
        List<Map<String,Object>> list3 = new ArrayList<Map<String,Object>>();  
        for (int i = 0; i < 10; i++) {  
            Map<String,Object> map = new HashMap<String,Object>();  
            map.put("pdn", i);  
            map.put("pdtest", "复合型防火墙防火墙是内部网与外部网的隔离点，起着监视和隔绝应用层通信流的作用，同时也常结合过滤器的功能。"+i);  
            list3.add(map);  
        }  
          
        dataMap.put("table3", list3);
        
        List<Map<String,Object>> list4 = new ArrayList<Map<String,Object>>();  
        for (int i = 0; i < 10; i++) {  
            Map<String,Object> map = new HashMap<String,Object>();  
            map.put("jdn", i);  
            map.put("jdtest", "说明作业调度，中级调度和进程调度的区别，并分析下述问题应由哪一级调度程序负责。"+i);  
            list4.add(map);  
        }  
          
        dataMap.put("table4", list4);
        
        List<Map<String,Object>> list11 = new ArrayList<Map<String,Object>>();  
        for (int i = 0; i < 10; i++) {  
            Map<String,Object> map = new HashMap<String,Object>();  
            map.put("fuck", i);  
            map.put("abc", "A"+i);  
            list11.add(map);  
        }  
        dataMap.put("table11", list11);
        
        List<Map<String,Object>> list12 = new ArrayList<Map<String,Object>>();  
        for (int i = 0; i < 10; i++) {  
            Map<String,Object> map = new HashMap<String,Object>();  
            map.put("fill", i);  
            map.put("def", "中级调度"+i);  
            list12.add(map);  
        }  
        dataMap.put("table12", list12);
        
        List<Map<String,Object>> list13 = new ArrayList<Map<String,Object>>();  
        for (int i = 0; i < 10; i++) {  
            Map<String,Object> map = new HashMap<String,Object>();  
            map.put("judge", i);  
            map.put("hij", "对"+i);  
            list13.add(map);  
        }
        dataMap.put("table13", list13);
        
        List<Map<String,Object>> list14 = new ArrayList<Map<String,Object>>();  
        for (int i = 0; i < 10; i++) {  
            Map<String,Object> map = new HashMap<String,Object>();  
            map.put("answer", i);  
            map.put("xyz", "说明作业调度，中级调度和进程调度的区别，并分析下述问题应由哪一级调度程序负责。"+i);  
            list14.add(map);  
        }  
          
        dataMap.put("table14", list14);
        //dataMap.put("list", list);  
    } 
}
