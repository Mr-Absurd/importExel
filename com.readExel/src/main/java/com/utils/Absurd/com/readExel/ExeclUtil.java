package com.utils.Absurd.com.readExel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;


import javax.swing.filechooser.FileSystemView;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExeclUtil {
	static File desktopDir = FileSystemView.getFileSystemView() .getHomeDirectory();
	static String desktopPath = desktopDir.getAbsolutePath();
    /**
     * @功能：手工构建一个简单格式的Excel
     */
    @SuppressWarnings("resource")
	public static void createExcel(List<List<String>> listBad) {
        // 第一步，创建一个webbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet("sheet1");
        sheet.setDefaultColumnWidth(20);// 默认列宽
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
        HSSFRow row = sheet.createRow((int) 0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        // 创建一个居中格式
        //style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        int num = listBad.get(0).size();
        // 添加excel title
        HSSFCell cell = null;
        for (int i = 0; i < num; i++) {
            cell = row.createCell((short) i);   
            cell.setCellStyle(style);
        }

        // 第五步，写入实体数据 实际应用中这些数据从数据库得到,list中字符串的顺序必须和数组strArray中的顺序一致
        int i = 0;
       
        for (List<String> list : listBad) {
            row = sheet.createRow(i);
            // 第四步，创建单元格，并设置值
            for (int j = 0; j < num; j++) {
                row.createCell((short) j).setCellValue(list.get(j));
            }

            // 第六步，将文件存到指定位置
            try {
                FileOutputStream fout = new FileOutputStream(desktopPath+"\\错误数据.xls");
                wb.write(fout);
                fout.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            i++;
        }
    }
}
