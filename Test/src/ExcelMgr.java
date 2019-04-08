
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMgr {

	public static void main(String[] args) {
		String fileName = "C:\\temp1\\PIID_Dev_1.3_��������_V1.3.xlsx";
		// System.out.println(ExcelMgr.findExcelByWord(fileName, "����"));
		//ExcelMgr.replaceHyperLink(fileName, fileName + ".xlsx", "����", "�ຣ");
		ExcelMgr.replaceExcelWords(fileName, fileName + ".xlsx", "x", "Y");
		List<String> list = null;

	}

	public static String replaceHyperLink(String filePath, String newFileName, String oldStr, String newStr) {

		// ������ʼ��
		XSSFWorkbook wb = null;
		FileInputStream fs;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;
		CreationHelper createrHelper = null;
		XSSFHyperlink link = null;

		// ��ʱ������������������
		String temp = "";

		// ����������
		StringBuffer resStrBuffer = new StringBuffer();

		try {
			fs = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fs);
		} catch (IOException e) {
			e.printStackTrace();
			resStrBuffer.append(e.getStackTrace());
		}

		// ѭ����ȡ�������
		for (int i = 0; i < wb.getNumberOfSheets(); i++) { // ��ȡÿ��ҳǩ
			sheet =wb.getSheetAt(i);
			createrHelper = wb.getCreationHelper();

			for (int j = 0; j <= sheet.getLastRowNum(); j++) { // ÿ��
				row = sheet.getRow(j);
				// Ϊ�գ�ֱ������
				if (row == null)
					continue;
				for (int k = 0; k <= row.getLastCellNum(); k++) {
					// ÿ��
					cell = row.getCell(k);
					if (cell == null)
						continue;
					System.out.println(cell.getRawValue());
					if (cell != null) {
						System.out.println(sheet.getSheetName() + "," + j + "," + k + ":");
						link = (XSSFHyperlink) createrHelper.createHyperlink(HyperlinkType.URL);
						if (cell.getHyperlink() == null)
							continue;
						temp = cell.getHyperlink().getAddress();
						if (temp.contains(oldStr)) {
							resStrBuffer.append(sheet.getSheetName() + "," + j + "," + k + ":" + temp + "-->>");
							temp = temp.replaceAll(oldStr, newStr);
							link.setAddress(temp);
							cell.setHyperlink(link);
							cell.setCellValue("x");
							resStrBuffer.append(temp + "\n");
							System.out.println(oldStr + "   " + newStr + "   " + temp);
						}

					}
				}
			}
		}

		// д��Ŀ���ļ�
		OutputStream out;
		try {
			out = new FileOutputStream(newFileName);
			wb.write(out);
			out.flush();
			out.close();
			wb.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			resStrBuffer.append(e.getStackTrace());
		} catch (IOException e) {
			e.printStackTrace();
			resStrBuffer.append(e.getStackTrace());
		}
		return resStrBuffer.toString();
	}

	/**
	 * �滻Excel������
	 * 
	 * @param filePath
	 * @param oldStr
	 * @param newStr
	 * @return
	 */
	public static String replaceExcelWords(String filePath, String newFilePath, String oldStr, String newStr) {

		XSSFWorkbook wb = null;
		FileInputStream fs;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;

		String temp = "";

		StringBuffer resStrBuffer = new StringBuffer();

		try {
			fs = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fs);
		} catch (IOException e) {
			resStrBuffer.append(e.getStackTrace());
		}

		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			sheet = wb.getSheetAt(i);

			for (int j = 0; j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null)
					continue;
				for (int k = 0; k <= row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					if (cell == null)
						continue;
					if (cell != null) {

						temp = ExcelMgr.getValue(cell);
						// if (temp.contains(oldStr)||StrMgr.getMatchDate(temp).size()>0) {
						if (temp.contains(oldStr)) {
							resStrBuffer.append(sheet.getSheetName() + "," + j + "," + k + ":" + temp + "-->>");
							temp = temp.replaceAll(oldStr, newStr);
							System.out.println(oldStr);
							cell.setCellValue(temp);
							resStrBuffer.append(temp + "\n");
						}

					}
				}

			}

		}

		FileOutputStream out;
		try {
			out = new FileOutputStream(new File(filePath.replaceAll(".xls", "") + "_new.xls"));
			wb.write(out);
			out.flush();
			out.close();
		} catch (FileNotFoundException e) {
			resStrBuffer.append(e.getStackTrace());
		} catch (IOException e) {
			resStrBuffer.append(e.getStackTrace());
		}

		return resStrBuffer.toString();
	}

	public static boolean findExcelByWord(String filePath, String word) {

		XSSFWorkbook wb = null;
		FileInputStream fs;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;

		String temp = "";

		try {
			fs = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fs);
		} catch (IOException e) {

		}

		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			sheet = wb.getSheetAt(i);

			for (int j = 0; j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null)
					continue;
				for (int k = 0; k <= row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					if (cell == null)
						continue;
					if (cell != null) {

						temp = ExcelMgr.getValue(cell);
						if (temp.contains(word)) {
							return true;
						} else {
							// int t = StrMgr.getMatchDate(temp).size();
							int t = 2;
							if (t > 0)
								return true;
						}

					}
				}

			}

		}

		return false;
	}

	// ��ȡ��Ԫ�������ֵ�������ַ�������
	public static String getValue(XSSFCell cell) {
		if (null != cell) {
			switch (cell.getCellType()) {
			case NUMERIC: // ����
				return cell.getNumericCellValue() + "   ";
			case STRING: // �ַ���
				return cell.getStringCellValue() + "   ";
			case BOOLEAN: // Boolean
				return cell.getBooleanCellValue() + "   ";
			case FORMULA: // ��ʽ
				return cell.getCellFormula() + "   ";
			case BLANK: // ��ֵ
				return "";
			case ERROR: // ����
				return "";
			default:
				return "δ֪����   ";
			}
		} else {
			return "";
		}
	}

	public static List<String> getAllHyperLink(String filePath) {

		XSSFWorkbook wb = null;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;
		List<String> files = new ArrayList<String>();
		FileInputStream fis = null;

		try {
			fis = new FileInputStream(filePath);

			wb = new XSSFWorkbook(fis);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			sheet = wb.getSheetAt(i);

			for (int j = 0; j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null)
					continue;
				for (int k = 0; k <= row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					if (cell == null || cell.getHyperlink() == null)
						continue;
					files.add(cell.getHyperlink().getAddress());
					System.out
							.println(sheet.getSheetName() + ":" + j + ":" + k + ":" + cell.getHyperlink().getAddress());

				}
				//
			}

		}
		try {
			fis.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return files;
	}

	public static List<String> getAllHyperLinkWithPost(String filePath) {

		XSSFWorkbook wb = null;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;
		List<String> files = new ArrayList<String>();
		FileInputStream fis = null;

		try {
			fis = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fis);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			sheet = wb.getSheetAt(i);

			for (int j = 0; j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null)
					continue;
				for (int k = 0; k <= row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					if (cell == null || cell.getHyperlink() == null)
						continue;
					files.add(sheet.getSheetName() + ":" + j + ":" + k + ":" + cell.getHyperlink().getAddress());

				}
				//
			}

		}
		try {
			fis.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return files;
	}

	public static Set<String> getUniqueHyperLink(String filePath) {

		XSSFWorkbook wb = null;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;
		Set<String> files = new HashSet<String>();
		FileInputStream fis = null;

		try {
			fis = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fis);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			sheet = wb.getSheetAt(i);

			for (int j = 0; j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null)
					continue;
				for (int k = 0; k <= row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					if (cell == null || cell.getHyperlink() == null)
						continue;
					files.add(cell.getHyperlink().getAddress());

				}
			}

		}
		try {
			fis.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return files;
	}
}
