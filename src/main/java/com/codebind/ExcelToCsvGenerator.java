package com.codebind;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.Map.Entry;

import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.OptionBuilder;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.io.FilenameUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToCsvGenerator {


	final static Logger logger=Logger.getLogger(ExcelToCsvGenerator.class);
	/**
	 * Method in which we will specified product name and If
	 * Product is not present in the List it print the appropriate
	 * message to the customer.
	 *
	 * It will generate output file i.e. comma separated files in the D drive of your System
	 *
	 * @param inputFile	  input file name
	 * @param outputFile  output file name
	 * @param product	  product name whose comma separated file has to be generated
	 * @param path        where file is located in the system 
	 * 
	**/
	static void forSpecifiedProduct(File inputFile,String product,String path)  {
		
		StringBuffer data = new StringBuffer();
		HashMap<String, Integer> hashMap = new HashMap<>();
		final String separator = System.getProperty("line.separator");
		Workbook workbook = null;
		
		String ext= FilenameUtils.getExtension(path);
		Set<String> set=new HashSet<String>();
		try 
		{
			try
			{
				// Platform new line
				if(ext.equals("xlsx")||ext.equals("xls"))
				{
					if (ext.equals("xlsx"))
					{
						workbook = new XSSFWorkbook(new FileInputStream(inputFile));
					} 
					else if (ext.equals("xls"))
					{
						workbook = new HSSFWorkbook(new FileInputStream(inputFile));
					} 
					else 
					{
						throw new FileNotFoundException();
					}
				}
				else
				{
					throw new NullPointerException();
				}
			}
			catch(NullPointerException e)
			{
				logger.info("Only .xlsx and .xls format are required. "+ext.toUpperCase() + " format is not allowed");
				System.exit(0);
			}
			catch(FileNotFoundException e)
			{
				logger.info("File Extension is valid "+ext.toUpperCase()+" but File is NOT present in given directory "+path.toUpperCase());
				System.exit(0);
			}
			Date date = new Date();
			Format formatter = new SimpleDateFormat("YYYY-MM-dd_hh-mm-ss");
		    
			File outputFile = new File("D:\\"+product+formatter.format(date)+".txt");

			FileOutputStream fos = new FileOutputStream(outputFile);
			// Get first sheet from the workbook
			Sheet sheet = workbook.getSheetAt(0);
			Row row;
			Cell cell;
			String temp1 ="";
			String temp2 ="";
			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				
				row = rowIterator.next();
				// just skip the rows if row number is 0 or 1
				if (row.getRowNum() == 0) {
					continue; 
				}
				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {

					cell = cellIterator.next();
					temp1 = cell.getStringCellValue();

					cell = cellIterator.next();
					temp2 = cell.getStringCellValue();
				}
				
				set.add(temp2);
				if (temp2.equals(product)) {
					if (hashMap.containsKey(temp1)) {
						logger.info("Duplicate Account key "+temp1+" found.");
						hashMap.put(temp1, (hashMap.get(temp1)) + 1);
					} else{
						logger.info("New Account key Found "+temp1+" found.");
						hashMap.put(temp1, 1);
					}

				}
			}
		
			if(set.contains(product))
			{
				logger.info("\n\n\n\n\n\n\nBRIEF REPORT OF YOUR PRODUCT IS\n\n\n\n\n\n\n "+product);
			}
			else
			{
				logger.info("YOUR LISTED PRODUCT IS NOT PRESENT IN THE LIST "+product);
			}

			// CODE TO PRINT THE HASH MAP

			Set<Entry<String, Integer>> entrySet = hashMap.entrySet();
			Iterator<Entry<String, Integer>> iter = entrySet.iterator();
			while (iter.hasNext()) {
				Entry<String, Integer> entry = iter.next();
				logger.info(" Account key "+entry.getKey()+" .Total  number of times repeated "+entry.getValue());
				data.append(" Account key "+entry.getKey()+" .Total  number of times repeated "+entry.getValue());
				data.append(separator);
			}
			fos.write(data.toString().getBytes());
			fos.close();

		}

		catch (Exception ioe) {
			ioe.printStackTrace();
		}
	}

	
	/**
	 *	Here we generating comma separated file for all the products when there is no product
	 *	has been specified.
	 * 	
	 *  It will generate output file i.e. comma separated files in the D drive of your System
	 *  
	 *  
	 * @param inputFile file from where we have to read the data
	 * @param path		the location were our file is present in the system
	 */
	@SuppressWarnings("resource")
	static void forAnyProduct(File inputFile, String path) {
		final String lineSeparator = System.getProperty("line.separator");
		Set<String> set = new HashSet<String>();
		Workbook workbook = null;
		String ext= FilenameUtils.getExtension(path);
		try {

			try
			{
				
				if(ext.equals("xlsx")||ext.equals("xls"))
				{
					if (ext.equals("xlsx"))
					{
						System.out.println(path.endsWith("xlsx"));
						workbook = new XSSFWorkbook(new FileInputStream(inputFile));
					} 
					else if (ext.equals("xls"))
					{
						workbook = new HSSFWorkbook(new FileInputStream(inputFile));
					} 
					else 
					{
						throw new FileNotFoundException();
					}
				}
				else
				{
					throw new NullPointerException();
				}
			}
			catch(NullPointerException e)
			{
				logger.info("Only .xlsx and .xls format are required. "+ext.toUpperCase() + " format is not allowed");
				System.exit(0);
			}
			catch(FileNotFoundException e)
			{
				logger.info("File Extension is valid "+ext.toUpperCase()+" but File is NOT present in given directory "+path.toUpperCase());
				System.exit(0);
			}
			// Get first sheet from the workbook
			Sheet sheet =  workbook.getSheetAt(0);
			Row row;
			Cell cell;
			String toStoreProductName = "";
			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {

				row = rowIterator.next();
				// just skip the rows if row number is 0 or 1
				if (row.getRowNum() == 0) {
					continue; 
				}
				
				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {

					cell = cellIterator.next();

					cell = cellIterator.next();
					toStoreProductName = cell.getStringCellValue();
				}

				set.add(toStoreProductName);

			}

			// Making the multiple files at a time
			Iterator<String> itr = set.iterator();
			while (itr.hasNext()) {
				StringBuffer data = new StringBuffer();
				HashMap<String, Integer> hashMap = new HashMap<>();
				String productName = itr.next();
				Date date = new Date();
				Format formatter = new SimpleDateFormat("YYYY-MM-dd_hh-mm-ss");
				
				File outputFile = new File("D:\\"+ FilenameUtils.getBaseName(path)+"."+ext+" "+productName+" "+formatter.format(date)+".txt");
				FileOutputStream fos = new FileOutputStream(outputFile);
				// Get the workbook object for XLSX file
				XSSFWorkbook wBook1 = new XSSFWorkbook(new FileInputStream(
						inputFile));

				// Get first sheet from the workbook
				XSSFSheet sheet1 = wBook1.getSheetAt(0);
				Row eachRow;
				Cell eachCell;
				String first = "";
				String second = "";
				// Iterate through each rows from first sheet
				Iterator<Row> rowIterator1 = sheet1.iterator();
				while (rowIterator1.hasNext()) {

					eachRow = rowIterator1.next();
					 // just skip the rows if row number is 0 or 1
					if (eachRow.getRowNum() == 0) {
						continue;
					}
					// For each row, iterate through each columns
					Iterator<Cell> eachCellIterator = eachRow.cellIterator();

					while (eachCellIterator.hasNext()) {

						eachCell = eachCellIterator.next();
						first = eachCell.getStringCellValue();
						eachCell = eachCellIterator.next();
						second = eachCell.getStringCellValue();
					}

					if (second.equals(productName)) {

						if (hashMap.containsKey(first)) {
							logger.info("Duplicate account key "+first+" found.");
							hashMap.put(first, (hashMap.get(first)) + 1);
						} else{
							logger.info("New Account key Found "+first+" found.");
							hashMap.put(first, 1);
						}
					}
				}

				// CODE TO PRINT THE HASH MAP
				Set<Entry<String, Integer>> entrySet = hashMap.entrySet();
				Iterator<Entry<String, Integer>> iter = entrySet.iterator();
				
				while (iter.hasNext()) {
					Entry<String, Integer> entry = iter.next();
					logger.info(" Account key "+entry.getKey()+".Total  number of times repeated "+entry.getValue());
					data.append(" Account key "+entry.getKey()+".Total  number of times repeated "+entry.getValue());
					data.append(lineSeparator);
				}

				fos.write(data.toString().getBytes());
				fos.close();
			}
		} catch (Exception ioe) {
			System.out.println("File is not Found in given path "+path);
		}

	}

	@SuppressWarnings({ "static-access" })
	public static void main(String[] args) throws ParseException 
	{

		Options options = new Options();
		
		options.addOption(OptionBuilder.withArgName("<destination>").hasArg()
				.isRequired().withDescription("File path name(Where file is present)").create("f"));
		
		options.addOption(OptionBuilder.withArgName("<product ID>").hasArg()
				.withDescription("Product Name followed by -f").create("p"));


		CommandLine line;

		try 
		{
			line = new BasicParser().parse(options, args);
			String path = line.getOptionValue("f");

			File inputFile = new File(path);

			if (line.getOptionValue("p") != null) 
			{
				
				forSpecifiedProduct(inputFile,line.getOptionValue("p"),path);

			} 
			else if (line.getOptionValue("f") != null) 
			{
				forAnyProduct(inputFile, path);
			}
			
			
		}

		catch (Exception e) 
		{
			
			logger.info("\n\n\nFile path is mandatory\n\n\n");
			HelpFormatter usage = new HelpFormatter();
			usage.printHelp("Please Only allowed Options with Excel To Csv Generator are following", options);
		}

	}
}