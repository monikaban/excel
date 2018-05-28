package convexcel;

/**
 * Converts the Excel Spreadsheet into CSV files. Both .xls and .xlsx file formats are supported.
 * Each Tab in the spreadsheet will be converted to a new csv file with name same as the Sheet name. 
 * 1st arg is Excel file name (with relative path from current folder). 2nd arg is Sheet name to be translated to csv file. 
 * If not specified, convert all sheets to separate csv files.
 * 
 * java ./sample-xlsx-file.xlsx Sheet1
 */

import org.apache.poi.ss.usermodel.*;
import java.util.Iterator;
import java.util.ArrayList;
import java.io.FileWriter;
import au.com.bytecode.opencsv.CSVWriter;
import java.io.File;
import java.io.IOException;

public class Excel2Csv {
    
    public static void main(String[] args) throws Exception{
    	
    	    System.out.println("Input args length:" + args.length);
    	    // 1st arg is the input Excel file   	
    		// 2nd arg is Sheet name to be translated to csv file. If not specified, convert all sheets to separate csv files.
    		String inputExcelFile = null;
    		String outputSheetName = null;
    		if(args.length <= 0) {
    			System.out.println("Please specify input filename");
    			System.exit(0);    		
    		}else if(args.length <= 1) {
    	 		inputExcelFile = args[0];
    	 	}else if(args.length > 1) {
    	 		inputExcelFile = args[0];
    	 		outputSheetName = args[1];
    	 	}
    	    final String outputSheet = outputSheetName;
    	    boolean csvFileGenerated = false;
 	        System.out.println("Inupt Excel file :" + inputExcelFile);
	    	 // Creating a Workbook from an Excel file (.xls or .xlsx)
	        Workbook workbook = WorkbookFactory.create(new File(inputExcelFile));
	        
	        System.out.println("Excel file read :" + inputExcelFile);
	        
	        // Retrieving the number of sheets in the Workbook
	        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets");

	        // Navigating Sheets. Loop through each Sheet and generate corresponding csv file
	        for(Sheet sheet : workbook) {	
	        	
	            System.out.println("Reading Sheet => " + sheet.getSheetName() + ". Total number of rows : " + sheet.getPhysicalNumberOfRows());
	            
	            if(outputSheet != null && !outputSheet.equals(sheet.getSheetName())) {
	            	continue; 
	            }
	            
	            // To iterate over the rows
	            Iterator<Row> rowIterator = sheet.iterator();
	            // OpenCSV writer object to create CSV file
	            FileWriter my_csv = null;
				try {
					my_csv = new FileWriter(sheet.getSheetName() +".csv");
				} catch (IOException e) {
					e.printStackTrace();
				}
	            CSVWriter my_csv_output=new CSVWriter(my_csv); 
	            int colCount = 0;
	            //Loop through rows.
	            while(rowIterator.hasNext()) {
	            	
	                Row row = rowIterator.next(); 
	                
	                ArrayList<String> csvList = new ArrayList<String>();
	                Iterator<Cell> cellIterator = row.cellIterator();
	                while(cellIterator.hasNext()) {
	                    Cell cell = cellIterator.next();                  
	                    printCellValue(cell, csvList, workbook);
	                }
	                my_csv_output.writeNext(csvList.toArray(new String[0]));
	                colCount = csvList.size();
	            }
	            System.out.println("Total number of columns :" + colCount);
	            try {
					my_csv_output.close(); //close the CSV file
				} catch (IOException e) {
					e.printStackTrace();
				}   
	            System.out.println("CSV file generated :" + sheet.getSheetName() +".csv");     
	            csvFileGenerated = true;
	        };
	        if(!csvFileGenerated) {
	        	System.out.println("No CSV files are generated for " + inputExcelFile);   
	        	if(outputSheet != null) {
	        		System.out.println("The Sheet name specified in 2nd argument ["+ outputSheetName +"] does not match any Sheets in the Excel file " + inputExcelFile); 
	        	}
            }
    }

    private static void printCellValue(Cell cell, ArrayList<String> csvList, Workbook workbook) {

        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
            	 csvList.add(Boolean.toString(cell.getBooleanCellValue()));
                break;
            case STRING:
            	 csvList.add(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    csvList.add(cell.getDateCellValue().toString());
                    //System.out.println("Date value | format" + cell.getDateCellValue() + "|" + cell.getCellStyle().getDataFormatString());
                } else {
                	DataFormatter dataFormatter = new DataFormatter();
                    String strValue = dataFormatter.formatCellValue(cell);
                   	csvList.add(strValue);
                }
                break;
            case FORMULA:            	   
            	        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            	        CellValue cellValue = evaluator.evaluate(cell);
            	        switch (cellValue.getCellType()) {
            	        case Cell.CELL_TYPE_BOOLEAN:
            	        	csvList.add(Boolean.toString(cellValue.getBooleanValue()));
            	            break;
            	        case Cell.CELL_TYPE_NUMERIC:
            	        	csvList.add(String.valueOf(cellValue.getNumberValue()));
            	            break;
            	        case Cell.CELL_TYPE_STRING:
            	        	csvList.add(cellValue.getStringValue());
            	            break;
            	        case Cell.CELL_TYPE_BLANK:
            	            break;
            	        case Cell.CELL_TYPE_ERROR:
            	            break;
            	        case Cell.CELL_TYPE_FORMULA: // CELL_TYPE_FORMULA will never happen
            	            break;
            	    }
                break;
            case BLANK:
            	 csvList.add("");
                break;
            default:
                csvList.add("");
        }
    }
}