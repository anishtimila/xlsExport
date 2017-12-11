package xlsExport;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class excel2txt {
    
	public void readExcel() throws IOException{
    try{
    	int i;
    	BufferedWriter output = null;
        
    	//Create an object of File class to open xlsx file
	    File file = new File("C:\\Users\\soar-it\\Desktop\\PropertyInventory2017.xlsx");
	    
	    //Create an object of FileInputStream class to read excel file
	    FileInputStream inputStream = new FileInputStream(file);
	    
	    //If it is xlsx file then create object of XSSFWorkbook class
	    //If it is xlsx file then create object of XSSFWorkbook class
        Workbook workbook = new XSSFWorkbook(inputStream);
                    for (i=0; i<workbook.getNumberOfSheets(); i++) {
        							//System.out.println("start i="+i);
                                    //Read sheet inside the workbook by its name
                                    Sheet sheet = workbook.getSheetAt(i);
                                    File txt = new File("C:\\Users\\soar-it\\Desktop\\test\\"+workbook.getSheetName(i)+".txt");
			                        output = new BufferedWriter(new FileWriter(txt));
                                       
                                    //Find number of rows in excel file
                                    int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
                                    String[] CompName= new String[70];
                                    //Create a loop over all the rows of excel file to read it
                                    for (int a = 0; a < rowCount+1; a++) {
                                        Row row = sheet.getRow(a);
                						if(row != null) {
                                    			if(row.getCell(10)!=null && row.getCell(10).getStringCellValue()!="") {
                                								CompName[a]=row.getCell(10).getStringCellValue();
                                                                if(row.getCell(10).getStringCellValue().toString()!="Computer Name") {
                                                                	System.out.println(row.getCell(10).getStringCellValue());
	                                                                output.write(row.getCell(10).getStringCellValue() + "\n");
	                                                                output.newLine();
                                                                }
                                                }
                                        }
                                   }
                                    output.close();
                    }
                    workbook.close();
    } catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
}

    

//Main function is calling readExcel function to read data from excel file
public static void main(String...strings) throws IOException{
	    //Create an object of class
	    excel2txt objExcelFile = new excel2txt();
	    //Call read file method of the class to read data
	    objExcelFile.readExcel();
    }
}