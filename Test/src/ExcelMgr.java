
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
		String fileName = "C:\\temp1\\PIID_Dev_1.3_西藏中设_V1.3.xlsx";
		// System.out.println(ExcelMgr.findExcelByWord(fileName, "西藏"));
		//ExcelMgr.replaceHyperLink(fileName, fileName + ".xlsx", "拉萨", "青海");
		ExcelMgr.replaceExcelWords(fileName, fileName + ".xlsx", "x", "Y");
		List<String> list = null;

	}

	public static String replaceHyperLink(String filePath, String newFileName, String oldStr, String newStr) {

		// 变量初始化
		XSSFWorkbook wb = null;
		FileInputStream fs;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;
		CreationHelper createrHelper = null;
		XSSFHyperlink link = null;

		// 临时变量，保存链接内容
		String temp = "";

		// 保存操作结果
		StringBuffer resStrBuffer = new StringBuffer();

		try {
			fs = new FileInputStream(filePath);
			wb = new XSSFWorkbook(fs);
		} catch (IOException e) {
			e.printStackTrace();
			resStrBuffer.append(e.getStackTrace());
		}

		// 循环读取表格内容
		for (int i = 0; i < wb.getNumberOfSheets(); i++) { // 读取每个页签
			sheet =wb.getSheetAt(i);
			createrHelper = wb.getCreationHelper();

			for (int j = 0; j <= sheet.getLastRowNum(); j++) { // 每行
				row = sheet.getRow(j);
				// 为空，直接跳过
				if (row == null)
					continue;
				for (int k = 0; k <= row.getLastCellNum(); k++) {
					// 每列
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

		// 写到目标文件
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
	 * 替换Excel中内容
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

	// 获取单元格各类型值，返回字符串类型
	public static String getValue(XSSFCell cell) {
		if (null != cell) {
			switch (cell.getCellType()) {
			case NUMERIC: // 数字
				return cell.getNumericCellValue() + "   ";
			case STRING: // 字符串
				return cell.getStringCellValue() + "   ";
			case BOOLEAN: // Boolean
				return cell.getBooleanCellValue() + "   ";
			case FORMULA: // 公式
				return cell.getCellFormula() + "   ";
			case BLANK: // 空值
				return "";
			case ERROR: // 故障
				return "";
			default:
				return "未知类型   ";
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
