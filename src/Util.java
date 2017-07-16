import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Util {

	public static void readExcel() {
		try{
			File file=new File("C:/Users/link/Desktop/0324/test");
			if(file.isDirectory()){
				File[] files=file.listFiles();
				for(File f:files){
					String fileName=f.getName();
					fileName=fileName.substring(0, fileName.lastIndexOf('.'));
					
					InputStream iStream=new FileInputStream(f);
					XSSFWorkbook xssfWorkbook = new XSSFWorkbook(iStream);
					
					XSSFSheet xssfSheet0 = xssfWorkbook.getSheetAt(0);
					for(int rowNum0 = 0; rowNum0 <= xssfSheet0.getLastRowNum(); rowNum0++){
//						String tmpSheetName=new String();
//						System.out.print(xssfSheet0.getSheetName());
						XSSFRow xssfRow0 = xssfSheet0.getRow(rowNum0);
						String industrySubName=xssfRow0.getCell(0).toString();
						String templateName=xssfRow0.getCell(1).toString();
						
						for (int numSheet = 1; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++){
							XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
//							System.out.println();
//							System.out.println("numSheet:"+numSheet+"\tsheetName:"+xssfSheet.getSheetName()+"\ttmpSheetName:"+tmpSheetName);
							if(!xssfSheet.getSheetName().equals(templateName))
								continue;
//							System.out.println(xssfSheet.getSheetName()+"\t");
							for(int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++){
//								System.out.print(xssfSheet.getSheetName());
								System.out.print(fileName+"\t"+industrySubName+"\t"+templateName+"\t");
								XSSFRow xssfRow = xssfSheet.getRow(rowNum);
								for (int colNum = 0; colNum < xssfRow.getLastCellNum(); colNum++) {
									System.out.print(xssfRow.getCell(colNum)+"\t");
								}
								System.out.println();
							}
						}
						
					}
					iStream.close();
					
//					for (int numSheet = 1; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++){
//						XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
//						System.out.println(file.getName()+"\t"+xssfSheet.getSheetName());
//					}
					
				}
			}
			
			
//			for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++){
//				XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
//				System.out.println(xssfSheet.getSheetName());
//				for(int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++){
//					System.out.print(xssfSheet.getSheetName());
//					XSSFRow xssfRow = xssfSheet.getRow(rowNum);
//					for (int colNum = 0; colNum < xssfRow.getLastCellNum(); colNum++) 
//						System.out.print("\t"+xssfRow.getCell(colNum));
//					System.out.println();
//				}
//			}
		}catch(Exception e){
			System.out.println(e.getMessage());
		}finally{
			
		}
	}
}
