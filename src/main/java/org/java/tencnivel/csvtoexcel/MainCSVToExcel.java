/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.java.tencnivel.csvtoexcel;


import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.commons.lang3.StringEscapeUtils;

// Import the necessary Native libraries
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
// import java.io.FileReader;
import java.io.InputStreamReader;

/**
 *
 * @author vlaugier
 */
public class MainCSVToExcel {

    public static void main(String[] args) throws Exception {
	    
		// Variable declaration
		Workbook wb;
		Sheet ws;
		Row row;
		Cell cell;
		int rownum;
		FileOutputStream outstream;
		String extension = "", delim = "", infiles = "", infile = "", outfile = "", line = "", dirchar = "", sheetname = "", encoding = "", usage = "Usage: Csv2Excel -t [xlsx|xls] -o [outfile] -d [delimiter] -i [infile1:infile2:infile3...]";
		String[] data, os, files;
		BufferedReader br = null;
		
		try {
			// Check enough arguments were supplied
			if(args.length < 10) {
				throw new Exception(usage);
			}
			
			// Loop through and assign the arguments
			for(int a = 0; a < args.length; a++) {
				// If -t option assign the next arg to the extension
				if(args[a].equals("-t")) {
					a++;
					extension = args[a].toLowerCase();
				}
				// Else if -o option assign the next arg to the outfile
				else if(args[a].equals("-o")) {
					a++;
					outfile = args[a];
				}
				// Else if -d option assign the next arg to the delimiter
				else if(args[a].equals("-d")) {
					a++;
					delim = "\\" + args[a];
				}
				// Else if -i option assign the next arg to the input files array
				else if(args[a].equals("-i")) {
					a++;
					infiles = args[a];
				}
				// Else if -e option assign the encoding for the input file
				else if(args[a].equals("-e")) {
					a++;
					encoding = args[a];
				}
			}
			
			// Check the file extension is valid
			if(!extension.equals("xls") && !extension.equals("xlsx")) {
				throw new Exception(usage);
			}
			
			// Create the workbook based upon the file type arg
			if(extension.equals("xls")) {
				// Excel 2003 workbook
				wb = new HSSFWorkbook();
			}
			else {
				// Excel 2007+ workbook
				wb = new XSSFWorkbook();
			}
			
			// Set the dirchar
			os = System.getProperty("os.name").split(" ");
			if(os[0].equals("Windows")) {
				dirchar = "\\";
			}
			else {
				dirchar = "/";
			}
		    
			// Loop through the input files and create a sheet for each
			files = infiles.split("\\:");
			for(int k = 0; k < files.length; k++) {
				
				// Read the input file
				infile = StringEscapeUtils.escapeJava(files[k]);
				// br = new BufferedReader(new FileReader(infile));
				br = new BufferedReader(new InputStreamReader(new FileInputStream(infile), encoding));
				
				// Create the new worksheet
				//sheetname = infile.substring(infile.lastIndexOf(dirchar) + 1, infile.lastIndexOf('.'));
                                sheetname = infile.substring(infile.lastIndexOf(dirchar) + 1);                                
				ws = wb.createSheet(sheetname);
				System.out.println("Adding worksheet: " + sheetname);
				
				// Reset the rownum
				rownum = 0;
				
				// Loop through the rows of the CSV file
				while((line = br.readLine()) != null) {
					
					//Split the current line into an array
					data = line.split(delim);
					
					// Create the row and check if there is data in the input row
					row = ws.createRow(rownum);
					if(data == null) continue;
					
					// Loop through the cells of the row
					for(int j = 0; j < data.length; j++) {
						
						// Create the cell and set the value of the cell
						cell = row.createCell(j);
						cell.setCellValue(data[j]);
						
					}
					
					// Increment the rownum
					rownum++;
					
				}
				
				// Close the buffered reader
				br.close();
				
			}
			
			// Define the output extension
			if(wb instanceof XSSFWorkbook) {
				outfile += ".xlsx";
			}
			else {
				outfile += ".xls";
			}
			
			// Create the output file
			System.out.println("Saving: " + outfile);
			outstream = new FileOutputStream(outfile);
			wb.write(outstream);
			outstream.close();
			System.out.println("\nSaved successfully!");
		}
		catch(Exception e) {
			System.out.println(e.getMessage());
		}

	}
    
}
