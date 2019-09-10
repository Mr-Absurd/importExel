package com.utils.Absurd.com.readExel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import javax.swing.filechooser.FileSystemView;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import javassist.ClassPool;
import javassist.CtClass;
import javassist.CtMethod;
import javassist.CtNewMethod;


public class ReadExel {
	
	/**
	 * 使用本程序注意事项:
	 * 请保证xls或者xlsx没有为空或者空格的数据.
	 * 请保证语句书写格式,新增语句格式为 [INSERT INTO TABLENAME],更改语句格式为 [UPDATE PSCS.TABLENAME].
	 * main方法调用 readExcel("文件地址","语句格式").
	 * 生成文件如果会在桌面生成名字"ace.sql"的文件,错误数据会在桌面生成"错误数据.xls"的文件
	 * xls/xlsx 文件第一行为字段名,请务必核对该列数据为插入数据库的数据,当为update时.默认最后一个列的数据为更改的标识符.
	 * 未知错误请自行debug寻找原因,后期代码会放在git上面更新,添加其他功能.
	 * 该程序支持打成jar包,配置文件已加,前端页面已写,YFileChooser为前端.
	 * */
	public static void main(String[] args) {
		
		readExcel("C:\\Users\\Absurd\\Desktop\\客商信息\\TEST.xlsx", "UPDATE PSCS.TPGLV10");
		System.out.println("111111111111111111111111");
	}
	 /**
	   * 根据fileType不同读取excel文件
	   *
	   * @param path
	   * @param path
	   * @throws IOException
	   */
	  @SuppressWarnings({ "resource", "deprecation" })
	public static String readExcel(String path,String sql) {
	    String fileType = path.substring(path.lastIndexOf(".") + 1);
	    // return a list contains many list
	    List<String> listSql = new ArrayList<String>();
	    List<List<String>> lists = new ArrayList<List<String>>();
	    List<List<String>> listBad = new ArrayList<List<String>>();
	    //读取excel文件
	    InputStream is = null;
	    try {
	      is = new FileInputStream(path);
	      //获取工作薄
	      Workbook wb = null;
	      if (fileType.equals("xls")) {
	        wb = new HSSFWorkbook(is);
	      } else if (fileType.equals("xlsx")) {
	        wb = new XSSFWorkbook(is);
	      } else {
	        return null;
	      }
	      	   	      
	      for(int i = 0;i<wb.getNumberOfSheets();i++) {
	    	  //读取第一个工作页sheet
		      Sheet sheet = wb.getSheetAt(i);
		      //第一行为标题
		      for (Row row : sheet) {
		        ArrayList<String> list = new ArrayList<String>();
		        for (Cell cell : row) {
		          //根据不同类型转化成字符串       	
		          cell.setCellType(CellType.STRING);
		          list.add(cell.getStringCellValue());
		        }
		        if(list.toString().indexOf("'")>0) {
		        	  listBad.add(list);
		        }else {
		        	  lists.add(list);
		        }        
		      
		      }
	      }
	      
	      listSql = lists.get(0);
	      
	      lists.remove(0);
	      
	      if(listBad.size()>0) {
	    	  ExeclUtil.createExcel(listBad);
	      } 
	      
	      //调用生成SQL语句
	      analysisSql(sql, lists, listSql);
	      
	    } catch (IOException e) {
	      e.printStackTrace();
	    } finally {
	      try {
	        if (is != null) is.close();
	      } catch (IOException e) {
	        e.printStackTrace();
	      }
	    }
	    return "成功";
	  }
	  
	@SuppressWarnings("rawtypes")
	public static String analysisSql(String sql,List<List<String>> list,List<String> listsql) {
		
		try {
			
			File desktopDir = FileSystemView.getFileSystemView() .getHomeDirectory();
			String desktopPath = desktopDir.getAbsolutePath();
			
			@SuppressWarnings("unchecked")
			List<String> sonList = new ArrayList();
			@SuppressWarnings("unchecked")
			List<String> resultList = new ArrayList();
			String sql1;
			if(sql.startsWith("INSERT")) {
				sql1=behindz(sql,listsql);
			}else {
				sql1=behind2z(sql,listsql);
			}
			//String sql1 = "\""+front+"VALUES\"+"+behind(behind);
			//String sql1 = behind2("UPDATE EQMS.TQSEP03  SET BACK_COL_1 = ITEM_DESC=");
			sql1 = "public String test(String[] flag) { return "+sql1 +";  }";
			
			createMethod cm = new createMethod();
			String className = "com.utils.Absurd.com.readExel.createMethod";
			 
			ClassPool pool = ClassPool.getDefault();
			CtClass cc = pool.get(className);
			CtMethod mthd = CtNewMethod.make(sql1, cc);
			cc.addMethod(mthd);

			AppClassLoader appClassLoader = AppClassLoader.getInstance();
			Class<?> clazz = appClassLoader.findClassByBytes(className, cc.toBytecode());

			Object obj = appClassLoader.getObj(clazz,cm);
			Method method = obj.getClass().getDeclaredMethod("test",String[].class );
			
			for (int i = 0; i < list.size(); i++) {
				sonList=list.get(i);
				String[] s = new String[listsql.size()];
				for(int j=0;j<listsql.size();j++) {
					s[j] = sonList.get(j);								
				}
				System.out.println(i);
				resultList.add((String) method.invoke(obj, new Object[] {s}));
			}
			
			
			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(desktopPath+"\\ace.sql"), "UTF-8"));
				
			//遍历集合
			for (String s : resultList) {
				bw.write(s);
				bw.newLine();
				bw.flush();
			}
			bw.close();
		} catch (Exception e) {
			e.printStackTrace();
			return "失败";
		} 
		
		return "成功";
	}
	
	//停用这个方法;
	private static String behind(String cont ) {
		String[] line = cont.split("VALUES");
		String front = line[0];
		String behind = line[1];
		
		String[] line1 = behind.split(",");
		
		String field = "\"(\'\"";
		String mark2 = "\"\');\"";
		String mark3 = "\"\',\'\"";
		
		for(int i=0;i<line1.length;i++) {
			if(i==line1.length-1) {
				field = field +"+flag["+i+"]+"+mark2;
			}else{
				field = field+"+flag["+i+"]+"+mark3;
			}
		}
		
		field= "\""+front+"VALUES\"+"+field;
		return field;
	}
	
	//insert 生成
	private static String behindz(String cont,List<String> list) {
		
		int num = list.size();
		
		
		String field = "";
		String mark2 = "\"\');\"";
		String mark3 = "\"\',\'\"";
		String mark4 = ",";
		String mark5 = ") VALUES (\'\"";
		String mark6 = " (";
		
		cont = cont + mark6;
		//先合成前面的数据.
		for(int i=0;i<num;i++) {
			if(i==num-1) {
				cont = cont +list.get(i)+mark5;
			}else{
				cont = cont+list.get(i)+mark4;
			}
		}
		
		//合成后面的语句
		for(int i=0;i<num;i++) {
			if(i==num-1) {
				field = field+"+flag["+i+"]+"+mark2;
			}else{
				field = field+"+flag["+i+"]+"+mark3;
			}
		}
		
		field= "\""+cont+field;
		return field;
	}
	
	
	//停用这个update生成
	private static String behind2(String cont) {
		
		String[] line = cont.split("=");
		String front = line[0];
		String mark = "=\'\"";
		String mark2 = "=\'\"";
		String mark3 = "\"\';\"";
		String mark4 = "\"\',";
		
	
		front = "\""+front+mark;
		
		for(int i=0;i<line.length;i++) {
			if(i==0) {
				front = front+"+flag["+i+"]+"+mark4;
			}else if(i==line.length-1) {
				front = front+"where "+line[i]+mark2+"+flag["+i+"]+";
			}else{
				front = front+line[i]+mark2+"+flag["+i+"]+"+mark4;
			}
		}
		
		int indx = front.lastIndexOf(",");
		if(indx!=-1){
			front = front.substring(0,indx)+front.substring(indx+1,front.length());
		}
		return front+mark3;
	}
	
	
	//新的修改数据
	private static String behind2z(String cont,List<String> list) {
		
		
	
		String mark = "=\'\"";
		String mark2 = "=\'\"";
		String mark3 = "\"\';\"";
		String mark4 = "\"\',";
		
	
		cont = "\""+cont+" set ";
		
		for(int i=0;i<list.size();i++) {
			if(i==list.size()-1) {
				cont = cont+"where "+list.get(i)+mark2+"+flag["+i+"]+";
			}else{
				cont = cont+list.get(i)+mark2+"+flag["+i+"]+"+mark4;
			}
		}
		
		int indx = cont.lastIndexOf(",");
		if(indx!=-1){
			cont = cont.substring(0,indx)+cont.substring(indx+1,cont.length());
		}
		return cont+mark3;
	}
}
