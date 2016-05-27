package com.david.lastModify;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.TrueFileFilter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


/**
 * Hello world!
 *
 */
public class App{
	
	 public static void main( String[] args ) throws ParseException, IOException{
		Workbook wb =  new HSSFWorkbook();
		//Create style
		CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setColor(HSSFColor.BLUE.index);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND); 
		
		Sheet noteSheet = wb.createSheet("Summary");
    	String time="2015-12-01 00:00:00";
    	int noteSheetRow=0;
    	fillRow(noteSheet,noteSheetRow++,style,0,"Checking Time for Date after ",time);
    	
    	SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd HH:mm:SS");
    	Date targetDate=sdf.parse(time);
    	System.out.println(targetDate);
    	String root=args[0];
    	File file=new File(root);
    	for (String folderName:file.list()){
    		File folder=new File(root+File.separatorChar+folderName.trim());
    		if (!folder.isDirectory() || !folder.getName().endsWith(".ear")) continue;
    		System.out.println(folderName);
    		Sheet earSheet = wb.createSheet(folderName.trim());
    		int row=0;
    		fillRow(earSheet,row++,style,null,"File","Creation Time","Last Access Time","Last Modified Date");
    		fillRow(noteSheet,noteSheetRow++,style,0,"Project folder scaned:",folder.getAbsolutePath());
    		if (folder.exists()){
    			for (File f:FileUtils.listFilesAndDirs(folder, TrueFileFilter.INSTANCE, TrueFileFilter.INSTANCE)){
    				Path path=Paths.get(f.getAbsolutePath());
    				BasicFileAttributes attr = Files.readAttributes(path, BasicFileAttributes.class);
 
    	        	Date lastModify=new Date(f.lastModified());
    	        	if (lastModify.after(targetDate) && f.isFile()){

    	        		fillRow(earSheet,row++,null,null,f.getAbsolutePath(),attr.creationTime().toString(),attr.lastAccessTime().toString(),attr.lastModifiedTime().toString());
    	        	}
    	        }
    		} else {
    			System.err.println("Sorry Foldre not exists "+folder.getAbsolutePath());
    		}
    		earSheet.autoSizeColumn(0);
    		earSheet.autoSizeColumn(1);
    		earSheet.autoSizeColumn(2);
    		earSheet.autoSizeColumn(3);
    	}
        FileOutputStream out = new FileOutputStream("result.xls");
		noteSheet.autoSizeColumn(0);
		noteSheet.autoSizeColumn(1);
        wb.write(out);
        out.close();
        wb.close();
        System.out.println("Done!");
        
    }
    

	private static void fillRow(Sheet noteSheet, int row, CellStyle style, Integer index,String...strings ) {
		Row r = noteSheet.createRow(row);
		int c=0;
		for (String value:strings){
			Cell cell=r.createCell(c);
			cell.setCellValue(value);
			if (style!=null && index==null){
				cell.setCellStyle(style);
			} else if (index!=null && c==index.intValue()){
				cell.setCellStyle(style);
			}
			c++;	
		}		
	}
}
