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
	String username = System.getProperty("user.name");
	String pathPrefix = "C:/Users/" + username + "/Desktop/";
	String extension = ".xlsx";
	String fileName;
	String path;

	public ArrayList<CRDate> readExcel() throws IOException {
		ArrayList<CRDate> storeDay = new ArrayList<>();

		FileInputStream inputStream = new FileInputStream(new File(path));

		Workbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = firstSheet.iterator();

		while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			if (nextRow.getRowNum() == 0) {
				continue; // just skip the rows if row number is 0
			}
			String strDateFormat = "MM/dd/yy hh:mm:ss aa";
			Date dateBefore = null;
			Date dateAfter = null;
			Date dateAfterResolved = null;
			Cell createdDateCell = nextRow.getCell(1);
			try {
				dateBefore = new SimpleDateFormat(strDateFormat).parse(createdDateCell.getStringCellValue());
			} catch (ParseException e) {
				System.err.println("Empty/invalid row found!!!");
				continue;
			}
			Cell closedDateCell = nextRow.getCell(2);
			try {
				dateAfter = new SimpleDateFormat(strDateFormat).parse(closedDateCell.getStringCellValue());
			} catch (ParseException e) {
				dateAfter = new Date();
			}
			Cell resolvedDateCell = nextRow.getCell(3);
			try {
				dateAfterResolved = new SimpleDateFormat(strDateFormat).parse(resolvedDateCell.getStringCellValue());
			} catch (ParseException e) {
				dateAfterResolved = new Date();
			}
			CRDate crd = new CRDate();
			long closedDifference = dateAfter.getTime() - dateBefore.getTime();
			crd.daysBetweenClosed = (int) (closedDifference / (1000 * 60 * 60 * 24));
			long resolvedDifference = dateAfterResolved.getTime() - dateBefore.getTime();
			crd.daysBetweenResolved = (int) (resolvedDifference / (1000 * 60 * 60 * 24));
			crd.timeTaken = crd.daysBetweenClosed - crd.daysBetweenResolved;
			storeDay.add(crd);
		}
		inputStream.close();
		return storeDay;
	}

	public void writeExcel(ArrayList<CRDate> l1) {

		FileInputStream in;
		try {
			in = new FileInputStream(new File(path));
			XSSFWorkbook workbook = new XSSFWorkbook(in);
			XSSFSheet firstSheet = workbook.getSheetAt(0);
			Iterator<CRDate> i = l1.iterator();

			XSSFRow row2 = firstSheet.getRow(0);
			if (row2 == null)
				row2 = firstSheet.createRow(0);
			Cell cell2 = row2.getCell(4, Row.CREATE_NULL_AS_BLANK);
			Cell cell3 = row2.getCell(5, Row.CREATE_NULL_AS_BLANK);
			Cell cell4 = row2.getCell(6, Row.CREATE_NULL_AS_BLANK);
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
			row2.getCell(4).setCellStyle(style);
			row2.getCell(5).setCellStyle(style);
			row2.getCell(6).setCellStyle(style);
			cell2.setCellValue("Closed Duration");
			cell3.setCellValue("Resolved Duration");
			cell4.setCellValue("Difference");
			firstSheet.autoSizeColumn(4);
			firstSheet.autoSizeColumn(5);
			firstSheet.autoSizeColumn(6);
			/*
			 * end styling
			 */
			int rownum = 1;

			while (i.hasNext()) {
				XSSFRow row = firstSheet.getRow(rownum++);
				if (row != null) {
					CRDate temp = (CRDate) i.next();
					Cell closedCell = row.createCell(4);
					closedCell.setCellValue(temp.daysBetweenClosed);
					Cell resolvedCell = row.createCell(5);
					resolvedCell.setCellValue(temp.daysBetweenResolved);
					Cell diffCell = row.createCell(6);
					diffCell.setCellValue(temp.timeTaken);
					System.out.println("Row : " + rownum + " inserted with value = " + temp.daysBetweenClosed + ", "
							+ temp.daysBetweenResolved + " and " + temp.timeTaken);
				}
			}

			in.close();
			try {
				FileOutputStream fos = new FileOutputStream(new File(path));
				workbook.write(fos);
				fos.close();
				System.err.println("\nDATA INSERTED SUCCESSFULLY");
			} catch (Exception e3) {
				System.err.println("WRITING FAILED. File already open.");
			}

		} catch (Exception e2) {
			e2.printStackTrace();
		}

	}

	public static void main(String[] args) {
		Scanner scn = new Scanner(System.in);
		System.out.println("Enter file name");
		DateFinder df = new DateFinder();
		df.fileName = scn.nextLine();
		df.path = df.pathPrefix + df.fileName + df.extension;
		System.out.println(df.fileName);
		System.out.println(df.path);
		try {
			ArrayList<CRDate> al = df.readExcel();
			df.writeExcel(al);
		} catch (IOException e) {
			e.printStackTrace();
		}
		scn.close();
	}
}
