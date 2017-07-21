package getfile;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.write.WriteException;

public class GetFileName {
	static String searchFile = "";
	static String fileType = "xls";
	static String folderName = System.getProperty("user.dir")+"\\";
	static String[] searchFileList;
//	static ArrayList<String> fileNameList = new ArrayList<String>();
//	static ArrayList<String> fileUrlList = new ArrayList<String>();
	static List<Map<String,String>> rows = new ArrayList();

	public static void main(String[] args) throws WriteException, IOException {
		// TODO Auto-generated method stub
		File file = new File(folderName);
		try{
			searchFile = args[0];						
			fileType = args[1];						
		}catch(Exception e){
			
		}

		//���P�_�O�ؿ��٬O�ɮ�
		IsDirectory(file);
		
		//�Ƨǥ�
		Collections.sort(rows, new MapComparatorAsc());
		
		switch (fileType) {
		case "csv":
			CreateExcel.typeCsv(rows);			
			System.out.println("�ϥ�csv����");
			break;
		case "xls":
			CreateExcel.typeXls(rows);
			System.out.println("�ϥ�xls����");
			break;
		case "xlsx":
			CreateExcel.typeXlsx(rows);
			System.out.println("�ϥ�xlsx����");
			break;
		default:
			CreateExcel.typeCsv(rows);						
			System.out.println("�S�����w���T�ɮ׮榡�A�ϥ�csv����");
			break;
		}
		
		System.out.println("����");
	}
	
	//��ؿ����U���W�r��ilist�h���P�_�O�ؿ����ɦW
	public static void doFileIsDirectory(File file){
		File [] fileList = file.listFiles();
		for(int i =0 ; i<fileList.length ; i++ ){
			IsDirectory(fileList[i]);
		}
		
	}
	
	//�O�ؿ��N���j��A�O�ɮ״N�P�_�ɦW�O�_���]�t�n�䪺�W�r
	public static void IsDirectory(File file){
		if(file.isDirectory()){
			doFileIsDirectory(file);
		}else{
			SearchFileName(file);
		}
	}
	
	//���]�t�n�䪺�W�r�N���ɦW
	public static void SearchFileName(File file){
		Map columns;
		
		searchFileList =  searchFile.split(",");
		for(int i = 0 ; i < searchFileList.length ; i++){
			if(file.getName().toLowerCase().indexOf(searchFileList[i].toLowerCase())!= -1){
//				fileNameList.add(file.getName());
//				fileUrlList.add(file.toPath().toString().replace(folderName, ""));
				System.out.println("�ɮ� : "+file.getName());
				columns = new HashMap();
				columns.put("FileName", file.getName());
				columns.put("FileUrl", file.toPath().toString().replace(folderName, ""));
				rows.add(columns);
		
//				System.out.println("���| : " + file.toPath().toString().replace(folderName, ""));
//				System.out.println("URI : " + file.toURI());
//				System.out.println("URI + ASCII : " + file.toURI().toASCIIString());
//				System.out.println("������| : " + file.getAbsolutePath());
			}				
		}
	}
	
	//�Ƨǥ�
	static class MapComparatorAsc implements Comparator<Map<String, String>> {
        public int compare(Map<String, String> m1, Map<String, String> m2) {
        	String  v1 = String.valueOf(m1.get("FileName"));
        	String  v2 = String.valueOf(m2.get("FileName"));
            if(v1 != null){
                return v1.compareTo(v2);
            }
            return 0;
        }
  
    }
	

}
