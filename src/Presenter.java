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

	public Presenter() {
		tweaks = new TreeMap<String, Tweak>();
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
		t.setSupport(supportDesc);
		System.out.println(t);
		addTweak(t);
	}

	public void setCenter(CellStyle style) {
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
	}
	
	public void setTweakNameStyle(Cell cell, boolean available, boolean fromMyRepo) {
		if (available && fromMyRepo)
			cell.setCellStyle(myRepoTweakNameStyle);
		else if (!available)
			cell.setCellStyle(outdatedTweakNameStyle);
		else
			cell.setCellStyle(normalTweakNameStyle);
	}
	
	public void setSupportStyle(Workbook wb, Cell cell, XSSFFont boldFont, Support.SupportType type) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		switch (type) {
		case Full:
			style.setFillForegroundColor(green);
			break;
		case CodePartial:
		case OSPartial:
			style.setFillForegroundColor(IndexedColors.YELLOW.index);
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
	
	public void set64bitSupportStyle(Workbook wb, Cell cell, XSSFFont boldFont, boolean support) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		cell.setCellValue(support + "");
		style.setFillForegroundColor(support ? green : red);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldFont);
		setCenter(style);
		cell.setCellStyle(style);
	}
	
	public void load(Workbook wb) {
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

	public void output() throws IOException {
		Workbook wb = new XSSFWorkbook();
		load(wb);
		Sheet sheet = wb.createSheet();
		Row header = sheet.createRow(0);
		XSSFCellStyle headerStyle = (XSSFCellStyle) wb.createCellStyle();
		headerStyle.setFillForegroundColor(new XSSFColor(new Color(200, 200, 220)));
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
		int numOS = Version.iOSVersionMax - Version.iOSVersionMin + 1;
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
			ver.setCellValue((Version.iOSVersionMin + i - 1) + ".x");
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
			for (i = 1; i < supports.length; i++) {
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
		}
		FileOutputStream fileOut = new FileOutputStream("tweaks.xlsx");
		wb.write(fileOut);
		fileOut.close();
		wb.close();
	}

	public static void main(String[] args) throws IOException {
		Presenter p = new Presenter();
		p.addTweak("60fps,n64,my,[Intend to unlock 60 fps capability in iPhone 4s and some old iPads only]", "4:na,5:?,6-9:1,10:m");
		p.addTweak("AFVideo", "4:na,5-10:1");
		p.addTweak("Animated Weather Enabler,n64", "4-6:na,7:1,dis");
		p.addTweak("AppColorClose", "4-6:0,7-10:1");
		p.addTweak("Auto HDR Enabler,n64,[Too internal, couldn't make it]", "4-6:na,7:0,8-9:cp,dis");
		p.addTweak("Better Pano Button,n64", "4-5:na,6:0,dis");
		p.addTweak("BlurryBadges", "4-6:0,7-10:1");
		p.addTweak("BlurryBar", "4-6:0,7-9:1,10:m");
		p.addTweak("BlurryContrast", "4-6:na,7-10:1");
		p.addTweak("Burst mode", "4:?,5-9:1,10:0");
		p.addTweak("CamBlur7", "4-6:na,5-9:1,10:0");
		p.addTweak("CamElapsed+", "4-6:0,7-9:1,10:0");
		p.addTweak("Camera Button UI Mod,n64", "4-5:na,6:1,dis");
		p.addTweak("CameraModes", "4-6:na,7-9:1,10:0");
		p.addTweak("CamModeList,my", "4-6:na,7-9:1,10:0");
		p.addTweak("CamPad,my", "4-6:0,7-8:1,9-10:0");
		p.addTweak("CamRotate", "4:?,5-9:1,10:m");
		p.addTweak("CamSwitchDown,my", "4-6:na,7-8:0,9:1,stop");
		p.addTweak("CamToggleBlur", "4-6:na,7-8:1,9:?,10:0");
		p.addTweak("CamTouch", "4-8:na,9:1,10:0");
		p.addTweak("CamVolNormal,n64", "4:?,5-6:1,dis");
		p.addTweak("CamVolZoom", "4:?,5-8:1,9-10:0");
		p.addTweak("CamZoomNoReset", "4-7:1,dis");
		p.addTweak("CCFlashLightLevel", "4-6:na,7-8:1,9:?,dis");
		p.addTweak("Contrast70", "4-6:na,7:op:7.0,dis");
		p.addTweak("DismissProgress", "4-10:1");
		p.addTweak("Effects+", "4-6:na,7-8:1,9-10:0");
		p.addTweak("Emoji10 (iOS 6.0-8.2),my", "4:na,5:0,6-8:1,stop");
		p.addTweak("Emoji10 (iOS 8.3/8.4),my", "4-7:na,8:op;8.3+,stop");
		p.addTweak("Emoji10 (iOS 9.0+),my", "4-8:na,9:1,stop");
		p.addTweak("Emoji83,my", "4:na,5:0,6:cp,7-8:1,dis");
		p.addTweak("Emoji83+", "4-7:na,8:op;8.3+,9:1,dis");
		p.addTweak("EmojiAttributes,my", "4:na,5:?,6-9:1,10:?");
		p.addTweak("EmojiLayout,my", "4:na,5:0,6-7:1,8:op;<8.3,stop");
		p.addTweak("EmojiLocalization,my", "4:na,5:0,6-9:1,stop");
		p.addTweak("EmojiResources,my", "4:na,5:0,6-8:1,9:op:<9.2,stop");
		p.addTweak("exKeyboard,[Incomplete code injection in IOS 9.1+ Jailbreak]", "4-7:na,8:1,9:cp,10:m");
		p.addTweak("FaceDetectionDuringVideo", "4:na,5:?,6-9:1,10:m");
		p.addTweak("FastLoading", "4-10:1");
		p.addTweak("FB Unlimited Chat Heads[Deprecated feature]", "4-9:1,dis");
		p.addTweak("Flashorama", "4:na,5:0,6-9:1,10:0");
		p.addTweak("ForceReach", "4-7:na,8:1,dis");
		p.addTweak("Front HDR", "4:na,5:0,6-9:1,10:0");
		p.addTweak("FrontFlash[Require iOS 5+ since 1.7-1]", "4-9:1,10:0");
		p.addTweak("FullCameraLog", "4:?,5-8:1,dis");
		p.addTweak("FullNoPop", "4-8:na,9-10:1");
		p.addTweak("Handoff4S", "4-7:na,8-9:1,dis");
		p.addTweak("HDR Badge 7.0", "4-6:na,7:op;7.0,dis");
		p.addTweak("HighGraphics,my", "4-6:na,7-10:1");
		p.addTweak("IB Graphics Selector", "4:na,5-8:1,9-10:?");
		p.addTweak("InternalPhotos", "4-6:na,7-10:1");
		p.addTweak("KBSwipe9", "4-8:na,9-10:1");
		p.addTweak("LandscapeProximity", "4:0,5-10:1");
		p.addTweak("LetMeSwitch", "4-7:na,8-9:1,10:m");
		p.addTweak("LetterPress9", "4-8:na,9:1,stop");
		p.addTweak("Live Effects Enabler", "4-6:na,7-9:1,10:0");
		p.addTweak("LLBiPT5", "4-5:na,6-8:1,9:m,stop");
		p.addTweak("LLBPano", "4-5:na,6-8:1,9-10:0");
		p.addTweak("LocationRemindersEnabler7", "4-6:na,7:1,dis");
		p.addTweak("LockEmoji", "4-6:na,7-9:1,10:m");
		p.addTweak("LockPredict", "4-7:na,8-9:1,10:0");
		p.addTweak("MoreDictation[Replaced by MoreSiri]", "4:na,5-7:1,dis");
		p.addTweak("MorePredict", "4-7:na,8-9:1,10:m");
		p.addTweak("MoreSiri", "4:na,5-9:1,stop");
		p.addTweak("MoreTimer", "4-7:na,8-9:1,10:0");
		p.addTweak("MyAssistive", "4:na,5-7:1,dis");
		p.addTweak("MyBias[iOS 9+ support is unlikely]", "4-7:na,8:1,9:0,dis");
		p.addTweak("MyLapse,my", "4-7:na,8:1,9-10:0");
		p.addTweak("NoGrayContrast", "4-6:na,7:op;7.1,8-9:1,10:m");
		p.addTweak("NoKeyPop", "4:0,5-8:1,stop");
		p.addTweak("NoPhotoGestures", "4-6:na,7-9:1,10:m");
		p.addTweak("NoSquare[Replaced by CameraModes]", "4-6:na,7-8:1,dis");
		p.addTweak("NoUpperCaseTable", "4-6:na,7-10:1");
		p.addTweak("NoWallpaperZoomAnimation", "4-6:na,7-10:1");
		p.addTweak("Nyan Cat for Velox,n64,[No Velox update iOS 9+]", "4-5:na,6-8:1,dis");
		p.addTweak("PanoMod", "4:na,5:0,6-9:1,10:0");
		p.addTweak("PhotoRes", "4-5:?,6-9:1,10:m");
		p.addTweak("PhotoScrubber,n64", "4-8:na,9:1,10:m");
		p.addTweak("PhotoTorch", "4-5:0,6-9:1,10:0");
		p.addTweak("ProximityCam", "4-6:0,7-8:1,9-10:0");
		p.addTweak("Randomy", "4-6:0,7:1,8-9:0,dis");
		p.addTweak("ReachFix,my", "4-7:na,8:1,9-10:?");
		p.addTweak("Record 'n' Torch", "4:0,5-9:1,10:0");
		p.addTweak("RecordPause", "4-7:0,8-9:1,10:0");
		p.addTweak("RoundedTable,my", "4-6:na,7-9:1,10:m");
		p.addTweak("SiriNoConfirm", "4:na,5-8:1,9-10:?");
		p.addTweak("Sketch9,n64", "4-8:1,dis");
		p.addTweak("Slo-mo Mod", "4-6:na,7-9:1,10:0");
		p.addTweak("SmoothCursor", "4-10:1");
		p.addTweak("SmoothPop", "4-6:0,7-10:1");
		p.addTweak("StaticZoom", "4-9:1,10:m");
		p.addTweak("Still Capture Enabler 2", "4-9:1,10:cp");
		p.addTweak("SwipeForMore", "4-7:0,8-9:1,10:m");
		p.addTweak("SwipeKey,my", "4-5:?,6-8:1,9-10:0");
		p.addTweak("SwitchAutofocus", "4:?,5-9:1,10:m");
		p.addTweak("TapForMore,my,[armv7 only]", "4-7:1,stop");
		p.addTweak("TransparentCameraBar", "4-9:1,10:m");
		p.addTweak("TypeAndTalk", "4-6:na,7-10:1");
		p.addTweak("UnlimShortcut", "4-8:na,9:1,10:m");
		p.addTweak("UnlockVol", "4-9:1,10:m");
		p.addTweak("Unrestricted Folders Naming", "4-10:1");
		p.addTweak("Video Zoom Mod", "4-6:na,7-10:1");
		p.addTweak("WeatherFix8", "4-7:na,8:1,dis");
		p.addTweak("Yellow Flash 7.0", "4-6:na,7:op;7.0,dis");

		p.output();
	}

}
