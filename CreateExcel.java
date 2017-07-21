package getfile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jxl.CellView;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class CreateExcel {
	
	public static void typeXls(List<Map<String,String>> rows) throws WriteException, IOException{
		 //新增excel,只能是xls格式
		 WritableWorkbook  workBook = Workbook.createWorkbook(new File("aaa.xls"));
		 
		 //新增頁籤,0為第幾張
		 WritableSheet sheet = workBook.createSheet("first sheet", 0);
		 
		 //設定字型
		 WritableFont myFont = new WritableFont(WritableFont.createFont("標楷體"), 14);
		 myFont.setColour(Colour.BLACK);
		 
		 //對儲存格格式設定
		 WritableCellFormat cellFormat = new WritableCellFormat ();
		 cellFormat.setFont(myFont); // 指定字型
		 cellFormat.setBackground(Colour.YELLOW); // 背景顏色
		 cellFormat.setAlignment(Alignment.LEFT); // 對齊方式
		 
		 //欄寬自動調整
		 CellView cellView = new CellView();
		 cellView.setAutosize(true);
		 sheet.setColumnView(0, cellView);
		 sheet.setColumnView(1, cellView);
		 
		 Label label = new Label(0 , 0 , "檔名" , cellFormat);//設定儲存格內容，格式非必填
		 Label label2 = new Label(1 , 0 , "路徑" , cellFormat);
		 
		 //塞入儲存格內容
		 sheet.addCell(label);
		 sheet.addCell(label2);
		 
		 //塞檔名
		 for(int i = 0 ; i < rows.size() ; i++){
			 Label cellValue = new Label(0 , i + 1 , rows.get(i).get("FileName"));
			 Label cellValue2 = new Label(1 , i + 1 , rows.get(i).get("FileUrl"));
			 sheet.addCell(cellValue);
			 sheet.addCell(cellValue2);	 
		 }

		 //開始寫入excel，停止寫入
		 workBook.write();
		 workBook.close(); 
	}
	
	public static void typeCsv(List<Map<String,String>> rows) throws IOException{
		//新增檔案，類型是csv
		FileWriter fw = new FileWriter("bbb.csv");
		
		//開是塞入，逗號分隔列，\n是分行
        fw.write("檔名");  
        fw.write(",路徑\n");
		for(int i = 0 ; i < rows.size() ; i++){
	        fw.write(rows.get(i).get("FileName"));  
	        fw.write("," + rows.get(i).get("FileUrl") + "\n");  
		}
		        
        //寫入結束
        fw.close();
	}
	
	public static void typeXlsx(List<Map<String,String>> rows) throws IOException{
		//新增excel,只能是xlsx格式
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		 //新增頁籤
		XSSFSheet sheet = workbook.createSheet("清單");
		
		//第0行，標題
		Row rowHead = sheet.createRow(0);
		
		//跑兩列
		Cell cell_1 = rowHead.createCell(0,Cell.CELL_TYPE_STRING);//第一個參數：列數（從0開始），第二個參數：列類型	
		Cell cell_2 = rowHead.createCell(1,Cell.CELL_TYPE_STRING);
		
		cell_1.setCellValue("檔名");
		cell_2.setCellValue("路徑");
		
		//跑行數
		for(int rowNum = 0 ;  rowNum < rows.size() ; rowNum ++){
			//從第二行開始
			Row row = sheet.createRow(rowNum+1);
			
			Cell cell1 = row.createCell(0,Cell.CELL_TYPE_STRING);
			Cell cell2 = row.createCell(1,Cell.CELL_TYPE_STRING);
			cell1.setCellValue(rows.get(rowNum).get("FileName"));
			cell2.setCellValue(rows.get(rowNum).get("FileUrl"));
		}
		
		//設置自動調整大小
		sheet.autoSizeColumn(0);		
		sheet.autoSizeColumn(1);
		
		OutputStream out = new FileOutputStream("ddd.xlsx");
		workbook.write(out);
		out.flush();
		out.close();
		
	}
}
