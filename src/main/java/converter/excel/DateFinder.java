package converter.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DateFinder {
	static String username=System.getProperty("user.name");

	public static ArrayList<Integer> readExcel(int sheet, int readingCol, String fileName) throws IOException {
		ArrayList<Integer> storeDay = new ArrayList<>();
		String pathPrefix = "C:/Users/"+username+"/Desktop/";
		String extension = ".xlsx";
		String path = pathPrefix + fileName + extension;
		FileInputStream inputStream = new FileInputStream(new File(path));

		Workbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(sheet);
		Iterator<Row> iterator = firstSheet.iterator();

		while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			if (nextRow.getRowNum() == 0) {
				continue; // just skip the rows if row number is 0
			}
			Cell cell = nextRow.getCell(readingCol);
				String strDateFormat = "MM/dd/yyyy hh:mm:ss aa";
				try {
					Date dateBefore = new SimpleDateFormat(strDateFormat).parse(cell.getStringCellValue());
					Date dateAfter = new Date();
					long difference = dateAfter.getTime() - dateBefore.getTime();
					int daysBetween = (int) (difference / (1000 * 60 * 60 * 24));
					storeDay.add(daysBetween);
				} catch (ParseException e) {
					e.printStackTrace();
				}
		}
		inputStream.close();
		return storeDay;
	}

	static public void writeExcel(ArrayList<Integer> l1, int col, int sheet, String fileName) {

		FileInputStream in;
		try {
			String pathPrefix = "C:/Users/"+username+"/Desktop/";
			String extension = ".xlsx";
			String path = pathPrefix + fileName + extension;
			in = new FileInputStream(new File(path));
			XSSFWorkbook workbook = new XSSFWorkbook(in);
			XSSFSheet firstSheet = workbook.getSheetAt(sheet);

			Iterator<Integer> i = l1.iterator();

			XSSFRow row2 = firstSheet.getRow(0);
			if (row2 == null)
				row2 = firstSheet.createRow(0);
			Cell cell2 = row2.getCell(col, Row.CREATE_NULL_AS_BLANK);
			/*
			 * start styling
			 */
			XSSFCellStyle style = workbook.createCellStyle();
			Font font = workbook.createFont();
			font.setColor(IndexedColors.WHITE.getIndex());
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			font.setFontName(HSSFFont.FONT_ARIAL);
			font.setFontHeightInPoints((short) 10);
			style.setFont(font);
			XSSFColor myColor = new XSSFColor(new java.awt.Color(102, 102, 153)); // Blue BG
			style.setFillForegroundColor(myColor);
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			row2.getCell(col).setCellStyle(style);
			cell2.setCellValue("No. of Days");
			/*
			 * end styling
			 */
			int rownum = 1;
			int cellnum = col;
			firstSheet.autoSizeColumn(col);
			while (i.hasNext()) {
				XSSFRow row = firstSheet.getRow(rownum++);
				int temp = (Integer) i.next();
				Cell cell = row.createCell(cellnum);
				cell.setCellValue(temp);
				System.out.println("Row : " + rownum + " inserted with value = " +temp);
			}

			in.close();
			FileOutputStream fos = new FileOutputStream(new File(path));
			workbook.write(fos);
			fos.close();
		} catch (Exception e2) {
			e2.printStackTrace();
		}

	}

	public static void main(String[] args) {
		Scanner scn = new Scanner(System.in);
		System.out.println("Enter file name");
		String filename = scn.nextLine();
		System.out.println("Enter sheet no.");
		int sheet = scn.nextInt();
		System.out.println("Enter column to read data");
		int readingCol = scn.nextInt();
		System.out.println("Enter column to enter data");
		int col = scn.nextInt();
		try {
			ArrayList<Integer> al = readExcel(sheet, readingCol, filename);
			writeExcel(al, col, sheet, filename);
			System.err.println("Data inserted successfully");
		} catch (IOException e) {
			e.printStackTrace();
		}
		scn.close();
	}
}
