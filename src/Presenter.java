import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map.Entry;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Presenter {

	public TreeMap<String, Tweak> tweaks;
	private static XSSFCellStyle normalTweakNameStyle;
	private static XSSFCellStyle outdatedTweakNameStyle;
	private static XSSFCellStyle myRepoTweakNameStyle;
	private static XSSFColor green = new XSSFColor(new Color(100, 200, 100));
	private static XSSFColor red = new XSSFColor(new Color(250, 100, 100));
	private static XSSFColor headerColor = new XSSFColor(new Color(200, 200, 220));
	private int minOSVer;
	private int maxOSVer;

	public Presenter(int minVer, int maxVer) {
		tweaks = new TreeMap<String, Tweak>();
		minOSVer = minVer;
		maxOSVer = maxVer;
	}
	
	public Presenter() {
		this(Version.iOSVersionMin, Version.iOSVersionMax);
	}

	public void addTweak(Tweak tweak) {
		tweaks.put(tweak.name, tweak);
	}

	public void addTweak(String name, boolean supports64Bit, boolean available, boolean fromMyRepo, String comments) {
		addTweak(new Tweak(name, supports64Bit, available, fromMyRepo, comments));
	}

	public void addTweak(String name, boolean supports64Bit, boolean available, boolean fromMyRepo) {
		addTweak(name, supports64Bit, available, fromMyRepo, null);
	}

	public void addTweak(String tweakDesc, String supportDesc) {
		Tweak t = new Tweak(tweakDesc);
		t.setSupport(supportDesc, minOSVer, maxOSVer);
		addTweak(t);
	}

	public static void setCenter(CellStyle style) {
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
	}
	
	public static void setTweakNameStyle(Cell cell, boolean available, boolean fromMyRepo) {
		if (available && fromMyRepo)
			cell.setCellStyle(myRepoTweakNameStyle);
		else if (!available)
			cell.setCellStyle(outdatedTweakNameStyle);
		else
			cell.setCellStyle(normalTweakNameStyle);
	}
	
	public static void setSupportStyle(Workbook wb, Cell cell, XSSFFont boldFont, Support.SupportType type) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		switch (type) {
		case Full:
			style.setFillForegroundColor(green);
			break;
		case CodePartial:
			style.setFillForegroundColor(IndexedColors.YELLOW.index);
			break;
		case OSPartial:
			style.setFillForegroundColor(IndexedColors.GREEN.index);
			break;
		case Maybe:
			style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.index);
			break;
		case Unknown:
			style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
			break;
		case No:
			style.setFillForegroundColor(red);
			break;
		case NAI:
		case NA:
			style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
			break;
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldFont);
		setCenter(style);
		cell.setCellStyle(style);
	}
	
	public static void set64bitSupportStyle(Workbook wb, Cell cell, XSSFFont boldFont, boolean support) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		cell.setCellValue(support ? "Yes" : "No");
		style.setFillForegroundColor(support ? green : red);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldFont);
		setCenter(style);
		cell.setCellStyle(style);
	}
	
	public static void load(Workbook wb) {
		normalTweakNameStyle = (XSSFCellStyle) wb.createCellStyle();
		outdatedTweakNameStyle = (XSSFCellStyle) wb.createCellStyle();
		myRepoTweakNameStyle = (XSSFCellStyle) wb.createCellStyle();
		XSSFFont normalFont = (XSSFFont) wb.createFont();
		normalFont.setBold(true);
		normalTweakNameStyle.setFont(normalFont);
		XSSFFont outdatedFont = (XSSFFont) wb.createFont();
		outdatedFont.setBold(true);
		outdatedFont.setColor(IndexedColors.RED.getIndex());
		outdatedTweakNameStyle.setFont(outdatedFont);
		XSSFFont myRepoFont = (XSSFFont) wb.createFont();
		myRepoFont.setBold(true);
		myRepoFont.setColor(IndexedColors.DARK_BLUE.getIndex());
		myRepoTweakNameStyle.setFont(myRepoFont);
		XSSFColor bg = new XSSFColor(new Color(200, 220, 220));
		normalTweakNameStyle.setFillForegroundColor(bg);
		normalTweakNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		outdatedTweakNameStyle.setFillForegroundColor(bg);
		outdatedTweakNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		myRepoTweakNameStyle.setFillForegroundColor(bg);
		myRepoTweakNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

	public static void output(String fileName, Presenter p) throws IOException {
		TreeMap<String, Tweak> tweaks = p.tweaks;
		Workbook wb = new XSSFWorkbook();
		load(wb);
		Sheet sheet = wb.createSheet(fileName);
		Row header = sheet.createRow(0);
		XSSFCellStyle headerStyle = (XSSFCellStyle) wb.createCellStyle();
		headerStyle.setFillForegroundColor(headerColor);
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		XSSFFont font = (XSSFFont) wb.createFont();
		font.setBold(true);
		headerStyle.setFont(font);
		setCenter(headerStyle);
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		style.setFont(font);
		setCenter(style);
		Cell name = header.createCell(0);
		name.setCellValue("Name");
		name.setCellStyle(headerStyle);
		Cell iOSVer = header.createCell(1);
		iOSVer.setCellValue("Supported iOS Versions");
		iOSVer.setCellStyle(headerStyle);
		int numOS = p.maxOSVer - p.minOSVer + 1;
		for (int i = 1; i <= numOS; i++)
			header.createCell(i + 1);
		Cell s64 = header.createCell(1 + numOS);
		s64.setCellValue("64-bit");
		s64.setCellStyle(headerStyle);
		Cell comments = header.createCell(2 + numOS);
		comments.setCellValue("Comments");
		comments.setCellStyle(headerStyle);
		Row iOSVers = sheet.createRow(1);
		for (int i = 1; i <= numOS; i++) {
			Cell ver = iOSVers.createCell(i);
			ver.setCellValue((p.minOSVer + i - 1) + ".x");
			ver.setCellStyle(headerStyle);
		}
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, numOS));
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 1 + numOS, 1 + numOS));
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 2 + numOS, 2 + numOS));
		int tRow = 2;
		for (Entry<String, Tweak> entry : tweaks.entrySet()) {
			String tweakName = entry.getKey();
			Tweak tweak = entry.getValue();
			Row row = sheet.createRow(tRow);
			Cell nameCell = row.createCell(0);
			nameCell.setCellValue(tweakName);
			setTweakNameStyle(nameCell, tweak.available, tweak.fromMyRepo);
			sheet.autoSizeColumn(0);
			Support supports[] = tweak.supports.values().toArray(new Support[tweak.supports.size()]);
			int i, start = 0;
			Cell typeCell = row.createCell(1);
			typeCell.setCellValue(supports[start].toString());
			setSupportStyle(wb, typeCell, font, supports[start].type);
			// Merging algorithm, thanks to Erbazz
			for (i = start + 1; i < supports.length; i++) {
				if (supports[i].type != supports[i - 1].type) {
					if (i - start > 1)
						sheet.addMergedRegion(new CellRangeAddress(tRow, tRow, start + 1, i));
					typeCell = row.createCell(start + 1);
					typeCell.setCellValue(supports[start].toString());
					setSupportStyle(wb, typeCell, font, supports[start].type);
					start = i;
				}
			}
			typeCell = row.createCell(start + 1);
			typeCell.setCellValue(supports[start].toString());
			setSupportStyle(wb, typeCell, font, supports[start].type);
			if (i - start > 1)
				sheet.addMergedRegion(new CellRangeAddress(tRow, tRow, start + 1, i));
			Cell s64Cell = row.createCell(++i);
			set64bitSupportStyle(wb, s64Cell, font, tweak.supports64Bit);
			Cell commentsCell = row.createCell(++i);
			commentsCell.setCellValue(tweak.comments);
			sheet.autoSizeColumn(i);
			tRow++;
			supports = null;
		}
		FileOutputStream fileOut = new FileOutputStream(fileName + ".xlsx");
		wb.write(fileOut);
		fileOut.close();
		wb.close();
	}

}
