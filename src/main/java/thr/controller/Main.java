package thr.controller;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.pdfbox.io.IOUtils;
import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

//import com.aspose.ocr.AsposeOCR;
//import com.aspose.ocr.InputType;
//import com.aspose.ocr.OcrInput;
//import com.aspose.ocr.RecognitionResult;
import com.fasterxml.jackson.databind.ObjectMapper;

import net.sourceforge.tess4j.Tesseract;
import thr.model.Friend;
import thr.model.Record;
import thr.model.RecordData;
import thr.model.Totals;
import thr.utils.MavenProperties;

public class Main {

	private static Logger logger = LogManager.getLogger("tsumRecordReport");
	private static String TSUM_RECORD_PATH   = null;
	private static File   TSUM_RECORD_DIR    = null;
	private static String RECORD_TXT_NAME    = "record.txt";
	private static File   RECORD_TXT_FILE    = null;
	private static File[] FRIEND_IMAGE_FILES = null;
	
	private static boolean REMOVE_TRAILING_ELLIPSE = true;
	private static String  ELLIPSE = "...";
	
	private static int HEADER_ROW = 0;
	private static int HINT_ROWS  = 10;
	
	private static double PIXEL_TO_POINTS_RATIO_WIDTH = ((double)256 / 7);
	private static double PIXEL_TO_POINTS_RATIO_HEIGHT = ((double)72 / 96);
	
	private static Dimension imageSizePixels = new Dimension(255, 69); // Each image is 255x69
	private static Dimension imageSizePoints = new Dimension(calcPoints(imageSizePixels.width, true), calcPoints(imageSizePixels.height, false));
	
	private static Map<String, String>  fileName2friendName = new LinkedHashMap<String, String>();
	private static Map<String, Integer> fileName2imageId    = new LinkedHashMap<String, Integer>();
	private static Map<String, Integer> fileName2rowNum     = new LinkedHashMap<String, Integer>();
	
	private static String COMMON_HELP_MSG =
			"\t(1) Make sure Robotmon is configured...\n" +
			"\t\t(a) To auto-send/receive hearts" +
			"\t\t(b) To record sender (with enlarged images)\n" +
			"\t(2) Wait for Robotmon to perform a send/receive cycle.\n" +
			"\t(3) Confirm the '" + RECORD_TXT_NAME + "' file has been created.";
	
	private enum Sheets {
		SHEET0(RECORD_TXT_NAME),
		SHEET1("all friends");
		
		String label;
		
		Sheets(String label) {
			this.label = label;
		}
	}
	
	private enum ColumnsSheet0 {
		FRIEND_NAME("Friend Name", calcPoints(110, true)),
		SEQUENCE("#", calcPoints(35, true)),
		FILE_NAME("File Name", calcPoints(175, true)),
		IMAGE("Image", imageSizePoints.width), // 328
		LAST_RECEIVED("Last Received", calcPoints(245, true)),
		NUMBER_RECEIVED("# Received", calcPoints(90, true));
		
		String label;
		Integer fixedWidth;
		
		ColumnsSheet0(String label, Integer fixedWidth) {
			this.label = label;
			this.fixedWidth = fixedWidth;
		}
	}

	private enum ColumnsSheet1 {
		FRIEND_NAME("Friend Name", ColumnsSheet0.FRIEND_NAME.fixedWidth),
		SEQUENCE("#", ColumnsSheet0.SEQUENCE.fixedWidth),
		FILE_NAME("File Name", ColumnsSheet0.FILE_NAME.fixedWidth),
		IMAGE("Image", ColumnsSheet0.IMAGE.fixedWidth),
		LAST_RECEIVED("Last Received", ColumnsSheet0.LAST_RECEIVED.fixedWidth),
		NUMBER_RECEIVED("# Received", ColumnsSheet0.NUMBER_RECEIVED.fixedWidth);
		
		String label;
		Integer fixedWidth;
		
		ColumnsSheet1(String label, Integer fixedWidth) {
			this.label = label;
			this.fixedWidth = fixedWidth;
		}
	}
	
	@SuppressWarnings("unused")
	private static int calcPixels(int points, boolean width) {
		return (int)((double)points / (width?PIXEL_TO_POINTS_RATIO_WIDTH:PIXEL_TO_POINTS_RATIO_HEIGHT));
	}
	private static int calcPoints(int pixels, boolean width) {
		return (int)((double)pixels * (width?PIXEL_TO_POINTS_RATIO_WIDTH:PIXEL_TO_POINTS_RATIO_HEIGHT));
	}
	
	/**
	 * Primary entry point into this application.
	 * <ol>Parameter(s)...
	 *  <li>Path to 'tsum_record' folder</li>
	 * </ol>
	 * 
	 * @param args String[] containing argument(s) for this class.
	 */
	public static void main(String[] args) {
		logger.info("...\r\n***\r\n*** " + MavenProperties.obtainArtifactId() + " (v" + MavenProperties.obtainVersion() + ")\r\n***");
		if (args == null || args.length < 1) {
			logger.error("Missing required argument(s).");
			System.err.println(" Syntax: java -jar " + MavenProperties.obtainArtifactId() + "-" + MavenProperties.obtainVersion() + ".jar ('tsum_record' folder)");
			System.err.println(" *Where: ('tsum_record' folder) is the path to the 'tsum_record' folder");
			System.err.println("Example: java -jar " + MavenProperties.obtainArtifactId() + "-" + MavenProperties.obtainVersion() + ".jar C:\\temp\\tsum_record");
		} else {
			TSUM_RECORD_PATH = args[0];
			TSUM_RECORD_DIR = new File(TSUM_RECORD_PATH);
			if (!TSUM_RECORD_DIR.exists()) {
				TSUM_RECORD_PATH += System.getProperty("file.separator") + "tsum_record";
				TSUM_RECORD_DIR = new File(TSUM_RECORD_PATH);
			}
			if (!TSUM_RECORD_DIR.exists() ) {
				logger.error("Specified folder cannot be found (" + args[1] + ").");
			} else {
				RECORD_TXT_FILE = new File(TSUM_RECORD_DIR, RECORD_TXT_NAME);
				if (!RECORD_TXT_FILE.exists()) {
					logger.error("'" + RECORD_TXT_NAME + "' cannot be found in specified folder.\n" + COMMON_HELP_MSG);
				} else {
					FRIEND_IMAGE_FILES = TSUM_RECORD_DIR.listFiles(new FilenameFilter() {
						public boolean accept(File dir, String name) {
			                return name.endsWith(".png");
			            }
					});
					if (FRIEND_IMAGE_FILES == null || FRIEND_IMAGE_FILES.length == 0) {
						logger.error("'Friend' image files cannot be found in the specified folder.\n" + COMMON_HELP_MSG);
					}
				}
				Main instance = new Main();
				instance.payload();
			}
		}
	}

	public void payload() {
		List<Friend> friends = null;
		try {
			FormulaEvaluator formulaEvaluator = null;
			// If XLS exists, index 1st sheet data and use it to augment 'record.txt' data
			File xlsFile = new File(TSUM_RECORD_DIR, "record.xlsx");
			Workbook workbook = null;
			if (xlsFile.exists()) {
				logger.info("Processing existing XLS file (" + xlsFile.getName() + ").");
				// ..open/load it if it already exists
				FileInputStream fis = new FileInputStream(xlsFile);
				workbook = new XSSFWorkbook(fis);
				fis.close();
				formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
				Sheet worksheet = workbook.getSheet(Sheets.SHEET0.label);
				// ..index 1st sheet's friend names and image ids (to use later)
				for (int x = 0; x < worksheet.getLastRowNum() + 1; x++) {
					if (x == HEADER_ROW) continue;
					Row row = worksheet.getRow(x);
					String fileName = getCellValue(row.getCell(2), formulaEvaluator);
					String friendName = getCellValue(row.getCell(0), formulaEvaluator);
					Cell imageCell = row.getCell(ColumnsSheet0.IMAGE.ordinal());
					Integer imageId = (int)Double.parseDouble(getCellValue(imageCell, formulaEvaluator));
					if (fileName != null && !fileName2friendName.containsKey(fileName)) {
						fileName2friendName.put(fileName, friendName);
					}
					if (imageId != null && !fileName2imageId.containsKey(fileName)) {
						fileName2imageId.put(fileName, imageId);
					}
					if (x > 0 && !fileName2rowNum.containsKey(fileName)) {
						fileName2rowNum.put(fileName, x);
					}
				}
				logger.info("Processed existing XLS data (size=" + fileName2friendName.size() + ").");
			}
			
			// Process 'record.txt'
			logger.info("Processing existing JSON file (" + RECORD_TXT_FILE.getName() + ").");
			String jsonString = Files.readString(RECORD_TXT_FILE.getAbsoluteFile().toPath());
			ObjectMapper objectMapper = new ObjectMapper();
			Record jsonData = objectMapper.readValue(jsonString, Record.class);
			logger.info("Cached JSON data (size=" + jsonData.getData().size() + " including totals).");
			friends = new ArrayList<Friend>();
			Totals totals = null;
			if (jsonData != null) {
				for (Map.Entry<String, RecordData> entry : jsonData.getData().entrySet()) {
					RecordData record = entry.getValue();
					record.setId(entry.getKey());
					if (record.getId().startsWith("f_") && record.getId().endsWith(".png")) {
						Friend friend = new Friend();
						friend.setId(record.getId());
						friend.setLastReceiveTime(record.getLastReceiveTime());
						friend.setReceiveCounts(record.getReceiveCounts());
						File imageFile = new File(TSUM_RECORD_DIR, friend.getId());
						if (imageFile.exists()) {
							if (!friends.contains(friend)) {
								friends.add(friend);
							}
						} else {
							logger.warn("Friend image file not found: " + friend.getId());
						}
					} else if (record.getId().equals("hearts_count")) {
						totals = new Totals();
						totals.setId(record.getId());
						totals.setReceivedCount(record.getReceivedCount());
						totals.setSentCount(record.getSentCount());
					} else {
						logger.warn("Unexpected record type: " + record.getId());
					}
				}
			}
			logger.info("Processed existing JSON data (size=" + friends.size() + ", totals=" + totals + ").");
			
			// Process 'friend' image file(s)
			if (!friends.isEmpty()) {
				logger.info("Processing friend data (size=" + friends.size() + ").");
				// Initialize OCR engine
//				AsposeOCR api = new AsposeOCR();
//				OcrInput images  = null;
//				ArrayList<RecognitionResult> results = null;
				// TODO Allow OCR scanning to be optional
				Tesseract tesseract = new Tesseract();
				// TODO Externalize Tesseract configuration
				//tesseract.setDatapath("src/main/resources/tessdata");
				tesseract.setDatapath("C:/Program Files/Tesseract-OCR/tessdata");
				tesseract.setLanguage("eng");
				tesseract.setPageSegMode(1);
				tesseract.setOcrEngineMode(1);
				File imageFile = null;
				for (int x = 0; x < friends.size(); x++) {
					Friend friend = friends.get(x);
					logger.debug("Processing friend (" + (x+1) + " of " + friends.size() + "): " + friend.getId());
					String friendName = fileName2friendName.get(friend.getId());
					if (friendName != null) {
						friend.setName(friendName);
						logger.debug("Pulled existing friend name (" + friendName + ") from XLS.");
					} else {
						imageFile = new File(TSUM_RECORD_DIR, friend.getId());
						// Scrape text from image
						// TODO Narrow area of OCR scan to make results more accurate
//						images = new OcrInput(InputType.SingleImage);
//						images.add(imageFile.getAbsolutePath());
//						results = api.Recognize(images);
						String result = tesseract.doOCR(imageFile);				
						// for friends name
//						if (results != null && results.size() > 0) {
//							friend.setName(results.get(1).recognitionText);
//						}
						String[] text = (result == null?null:result.split("\n"));
						if (text != null && text.length > 1) {
							friendName = StringUtils.defaultIfEmpty(text[1], StringUtils.EMPTY);
							if (REMOVE_TRAILING_ELLIPSE && friendName.endsWith(ELLIPSE)) {
								friendName = friendName.substring(0, friendName.length() - ELLIPSE.length());
							}
							friend.setName(friendName);
							if (StringUtils.isBlank(friendName)) {
								logger.warn("Empty friend name: " + friendName + " (from " + text + ")");
							}
						} else {
							logger.warn("Could not determine friend name: " + friend.getId());
						}
					}
				}
				logger.info("Procesed friend data (size=" + friends.size() + ").");
			} else {
				logger.warn("No friend data found.");
			}
			
			// Generate output file(s)
			if (!friends.isEmpty()) {
				// HTML
				StringBuilder sb = new StringBuilder();
				sb.append("<html>");
				sb.append("\t<main>");
				sb.append("\t\t<title>" + MavenProperties.obtainName() + "</title>");
				sb.append("\t</main>");
				sb.append("\t<body>");
				sb.append("\t\t<table border=1 cellpadding=1 cellspacing=1>");
				sb.append("\t\t\t<tr " + getHeaderRowStyleHTML() + ">");
				for (int y = 0; y < ColumnsSheet0.values().length; y++) {
					sb.append("<th>" + ColumnsSheet0.values()[y].label + "</th>");
				}
				sb.append("</tr>");
				for (int x = 0; x < friends.size(); x++) {
					Friend friend = friends.get(x);
					int totalReceived = 0;
					for (Map.Entry<String, Integer> receivedCount : friend.getReceiveCounts().entrySet()) {
						totalReceived += receivedCount.getValue();
					}
					sb.append("\t\t\t<tr>");
					sb.append("<td>" + (x+1) + "</td>");
					sb.append("<td>" + friend.getId() + "</td>");
					sb.append("<td><img src='" + friend.getId() + "'/></td>");
					sb.append("<td>" + friend.getName() + "</td>");
					sb.append("<td>" + friend.getLastReceiveTime() + "<br/>(" + friend.getLastReceivedDate() + ")" + "</td>");
					sb.append("<td>" + totalReceived + "</td>");
					sb.append("</tr>");
				}
				sb.append("\t\t</table>");
				sb.append("\t</body>");
				sb.append("</html>");
				File htmlFile = new File(TSUM_RECORD_DIR, "record.html");
				Files.write(htmlFile.toPath(), sb.toString().getBytes(), (htmlFile.exists()?StandardOpenOption.TRUNCATE_EXISTING:StandardOpenOption.CREATE));
				logger.info("Generated HTML file (" + htmlFile.getName() + ").");
				// XLS
				if (!xlsFile.exists()) {
					// ..create if/when it doesn't already exist
					workbook = createWorkbook(friends);
				} else {
					// ..update (first sheet) based on latest 'friends' data
					updateWorkbook(friends, workbook);
				}
				// ..save
				FileOutputStream fos = new FileOutputStream(xlsFile);
				workbook.write(fos);
				fos.close();
				logger.info("Generated XLS file (" + xlsFile.getName() + ").");
			}
		} catch (Exception e) {
			logger.error(e.getLocalizedMessage(), e);
		} finally {
			logger.info("Done.");
		}
	}

	private Workbook createWorkbook(List<Friend> friends) throws IOException, FileNotFoundException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		// Sheet 0: record.txt (with mapping to friend names)
		Sheet[] worksheet = { null, null };
		worksheet[0] = workbook.createSheet(Sheets.SHEET0.label);
		Row row = null;
		Cell cell = null;
		// Header row
		row = worksheet[0].createRow(HEADER_ROW);
		for (int y = 0; y < ColumnsSheet0.values().length; y++) {
			cell = row.createCell(y);
			cell.setCellStyle(getHeaderRowStyleXLSX(workbook));
			cell.setCellValue(ColumnsSheet0.values()[y].label);
		}
		// Friend rows
		CellStyle editableFriendCellStyle = getFriendRowStyleXLSX(workbook, true);
		CellStyle lockedFriendCellStyle   = getFriendRowStyleXLSX(workbook, false);
		for (int x = 0; x < friends.size(); x++) {
			Friend friend = friends.get(x);
			row = worksheet[0].createRow(x+1);
			int totalReceived = 0;
			for (Map.Entry<String, Integer> receivedCount : friend.getReceiveCounts().entrySet()) {
				totalReceived += receivedCount.getValue();
			}
			for (int y = 0; y < ColumnsSheet0.values().length; y++) {
				cell = row.createCell(y);
				switch (y) {
					case 0: // Columns.FRIEND_NAME.ordinal():
						cell.setCellValue(friend.getName());
						cell.setCellStyle(editableFriendCellStyle);
						break;
					case 1: // Columns.SEQUENCE.ordinal():
						cell.setCellValue(x+1);
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 2: // Columns.FILE_NAME.ordinal():
						cell.setCellValue(friend.getId());
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 3: // Columns.IMAGE.ordinal():
						int imagePictureId = -1;
						if (fileName2imageId.containsKey(friend.getId())) {
							imagePictureId = fileName2imageId.get(friend.getId());
							logger.debug("Pulled existing image (" + imagePictureId + ") from XLS.");
						} else {
							InputStream imageInputStream = new FileInputStream(new File(TSUM_RECORD_DIR, friend.getId()));
							byte[] imageBytes = IOUtils.toByteArray(imageInputStream);
							imageInputStream.close();
							imagePictureId = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
							logger.debug("Loaded new image (" + friend.getId() + ") from file-system.");
						}
						XSSFDrawing imageDrawing = (XSSFDrawing)worksheet[0].createDrawingPatriarch();
						XSSFClientAnchor imageAnchor = new XSSFClientAnchor();
						imageAnchor.setCol1(ColumnsSheet0.IMAGE.ordinal());
						imageAnchor.setCol2(ColumnsSheet0.LAST_RECEIVED.ordinal());
						imageAnchor.setRow1(row.getRowNum());
						imageAnchor.setRow2(row.getRowNum()+1);
						imageDrawing.createPicture(imageAnchor, imagePictureId);
						fileName2imageId.put(friend.getId(), imagePictureId);
						cell.setCellValue(imagePictureId);
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 4: // Columns.LAST_RECEIVED.ordinal():
						cell.setCellValue(friend.getLastReceiveTime() + "\n(" + friend.getLastReceivedDate() + ")");
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 5: // Columns.NUMBER_RECEIVED.ordinal():
						cell.setCellValue(totalReceived);
						cell.setCellStyle(lockedFriendCellStyle);
						break;
				}
			}
		}
		// Sheet 1: all friends (user-entered list of all their friends)
		worksheet[1] = workbook.createSheet(Sheets.SHEET1.label);
		// Header row
		row = worksheet[1].createRow(HEADER_ROW);
		for (int y = 0; y < ColumnsSheet1.values().length; y++) {
			cell = row.createCell(y);
			cell.setCellStyle(getHeaderRowStyleXLSX(workbook));
			cell.setCellValue(ColumnsSheet1.values()[y].label);
		}
		// Hint rows
		for (int x = 0; x < HINT_ROWS; x++) {
			row = worksheet[1].createRow(x+1);
			for (int y = 0; y < ColumnsSheet1.values().length; y++) {
				cell = row.createCell(y);
				switch (y) {
					case 0: // Columns.FRIEND_NAME.ordinal():
						cell.setCellValue("(Enter Name)");
						cell.setCellStyle(editableFriendCellStyle);
						break;
					case 1: // Columns.SEQUENCE.ordinal():
						cell.setCellFormula("INDEX('" + Sheets.SHEET0.label + "'!$B:$B,MATCH($A" + (row.getRowNum() + 1) + ",'" + Sheets.SHEET0.label + "'!$A:$A,0))");
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 2: // Columns.FILE_NAME.ordinal():
					case 3: // Columns.IMAGE.ordinal():
					case 4: // Columns.LAST_RECEIVED.ordinal():
					case 5: // Columns.NUMBER_RECEIVED.ordinal():
						cell.setCellFormula("VLOOKUP($B" + (row.getRowNum() + 1) + ",'" + Sheets.SHEET0.label + "'!$B:$F," + y + ",FALSE)");
						cell.setCellStyle(lockedFriendCellStyle);
						break;
				}
			}
		}
		// Create conditional formatting to highlight rows not associated with a friend
		SheetConditionalFormatting sheetCF = worksheet[1].getSheetConditionalFormatting();
		for (int rule = 0; rule < sheetCF.getNumConditionalFormattings(); rule++) {
			sheetCF.removeConditionalFormatting(rule);
		}
		ConditionalFormattingRule ruleFormula = sheetCF.createConditionalFormattingRule("ISNA($F2)");
        ruleFormula.createPatternFormatting().setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
        BorderFormatting borderFormatting = ruleFormula.createBorderFormatting();
        borderFormatting.setBorderTop(BorderStyle.DOTTED);
        borderFormatting.setBorderRight(BorderStyle.DOTTED);
        borderFormatting.setBorderBottom(BorderStyle.DOTTED);
        borderFormatting.setBorderLeft(BorderStyle.DOTTED);
        CellRangeAddress[] rangeFormula = {
        	new CellRangeAddress(HEADER_ROW + 1, worksheet[1].getLastRowNum(), 0, ColumnsSheet1.values().length - 1)
        };
        sheetCF.addConditionalFormatting(rangeFormula, ruleFormula);
		// Create a new data validation to cover all of column A
		// @see https://thecodeshewrites.com/2020/08/11/apache-poi-excel-java-dropdown-list-dependent/
		String dvReference = "'" + Sheets.SHEET0.label + "'!$A:$A";
		Name dvNamedArea = workbook.createName();
		dvNamedArea.setNameName("myFriends");
		dvNamedArea.setRefersToFormula(dvReference);
		CellRangeAddressList dvAddressList = new CellRangeAddressList(HEADER_ROW + 1, worksheet[1].getLastRowNum(), ColumnsSheet1.FRIEND_NAME.ordinal(), ColumnsSheet1.FRIEND_NAME.ordinal());
		DataValidationHelper dvHelper = worksheet[1].getDataValidationHelper();
		DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint("myFriends");
		DataValidation dataValidation = dvHelper.createValidation(dvConstraint, dvAddressList);
		dataValidation.setSuppressDropDownArrow(true);
		dataValidation.setShowPromptBox(true);
		worksheet[1].addValidationData(dataValidation);
		// Resize each sheets rows & columns
		for (int sheet = 0; sheet < Sheets.values().length; sheet++) {
			// Protecting entire sheet, however, column A cells are "unlocked" and therefore editable
			worksheet[sheet].protectSheet("");
			// Freeze the header row
			worksheet[sheet].createFreezePane(0, 1);
			// Row height (all rows except Header row)
			for (int x = 1; x < worksheet[sheet].getLastRowNum() + 1; x++) {
				worksheet[sheet].getRow(x).setHeightInPoints(imageSizePoints.height);
			}
			// Column width (all columns)
			for (int y = 0; y < (sheet == 0?ColumnsSheet0.values().length:ColumnsSheet1.values().length); y++) {
				if (sheet == Sheets.SHEET0.ordinal()) {
					worksheet[sheet].setColumnWidth(y, ColumnsSheet0.values()[y].fixedWidth);	
				} else if (sheet == Sheets.SHEET1.ordinal()) {
					worksheet[sheet].setColumnWidth(y, ColumnsSheet1.values()[y].fixedWidth);	
				}
			}
		}
		// Finalize workbook
		workbook.setActiveSheet(Sheets.SHEET0.ordinal());
		return workbook;
	}

	private void updateWorkbook(List<Friend> friends, Workbook workbook) throws IOException, FileNotFoundException {
		FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		// Sheet 0: record.txt (with mapping to friend names)
		Sheet[] worksheet = { null, null };
		worksheet[0] = workbook.getSheet(Sheets.SHEET0.label);
		Row row = null;
		Cell cell = null;
//		// Header row
//		row = worksheet[0].getRow(HEADER_ROW);
//		for (int y = 0; y < ColumnsSheet0.values().length; y++) {
//			cell = row.getCell(y);
//			cell.setCellStyle(getHeaderRowStyleXLSX(workbook));
//			cell.setCellValue(ColumnsSheet0.values()[y].label);
//		}
		// Friend rows
		CellStyle editableFriendCellStyle = getFriendRowStyleXLSX(workbook, true);
		CellStyle lockedFriendCellStyle   = getFriendRowStyleXLSX(workbook, false);
		// ..Update existing rows (and create new rows as needed)
		for (int x = 0; x < friends.size(); x++) {
			Friend friend = friends.get(x);
			if (fileName2rowNum.containsKey(friend.getId())) {
				// Update existing row
				row = worksheet[0].getRow(fileName2rowNum.get(friend.getId()));
				logger.debug("Updating (existing) row [" + row.getRowNum() + "]: " + friend.getId() + " (" + friend.getName() + ")");
			} else {
				// Create new row
				row = worksheet[0].createRow(worksheet[0].getLastRowNum() + 1);
				if (!fileName2rowNum.containsKey(friend.getId())) {
					fileName2rowNum.put(friend.getId(), row.getRowNum());
				}
				logger.debug("Creating (new) row [" + row.getRowNum() + "]: " + friend.getId() + " (" + friend.getName() + ")");
			}
			int totalReceived = 0;
			for (Map.Entry<String, Integer> receivedCount : friend.getReceiveCounts().entrySet()) {
				totalReceived += receivedCount.getValue();
			}
			for (int y = 0; y < ColumnsSheet0.values().length; y++) {
				cell = row.getCell(y);
				if (cell == null) {
					// Allow for new data
					cell = row.createCell(y);
				}
				switch (y) {
					case 0: // Columns.FRIEND_NAME.ordinal():
						cell.setCellValue(friend.getName());
						cell.setCellStyle(editableFriendCellStyle);
						break;
					case 1: // Columns.SEQUENCE.ordinal():
						cell.setCellValue(x+1);
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 2: // Columns.FILE_NAME.ordinal():
						cell.setCellValue(friend.getId());
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 3: // Columns.IMAGE.ordinal():
						int imagePictureId = -1;
						if (fileName2imageId.containsKey(friend.getId())) {
							imagePictureId = fileName2imageId.get(friend.getId());
							logger.debug("Pulled existing image (" + imagePictureId + ") from XLS.");
						} else {
							InputStream imageInputStream = new FileInputStream(new File(TSUM_RECORD_DIR, friend.getId()));
							byte[] imageBytes = IOUtils.toByteArray(imageInputStream);
							imageInputStream.close();
							imagePictureId = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
							logger.debug("Loaded new image (" + friend.getId() + ") from file-system.");
						}
						XSSFDrawing imageDrawing = (XSSFDrawing)worksheet[0].createDrawingPatriarch();
						XSSFClientAnchor imageAnchor = new XSSFClientAnchor();
						imageAnchor.setCol1(ColumnsSheet0.IMAGE.ordinal());
						imageAnchor.setCol2(ColumnsSheet0.LAST_RECEIVED.ordinal());
						imageAnchor.setRow1(row.getRowNum());
						imageAnchor.setRow2(row.getRowNum()+1);
						imageDrawing.createPicture(imageAnchor, imagePictureId);
						cell.setCellValue(imagePictureId);
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 4: // Columns.LAST_RECEIVED.ordinal():
						cell.setCellValue(friend.getLastReceiveTime() + "\n(" + friend.getLastReceivedDate() + ")");
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 5: // Columns.NUMBER_RECEIVED.ordinal():
						cell.setCellValue(totalReceived);
						cell.setCellStyle(lockedFriendCellStyle);
						break;
				}
			}
		}
		// ..remove any rows which are no longer needed
		List<Integer> redundantRows = new ArrayList<Integer>();
		for (int x = 0; x < worksheet[0].getLastRowNum() + 1; x++) {
			if (x == HEADER_ROW) continue;
			if (!fileName2rowNum.containsValue(x)) {
				redundantRows.add(x);
			}
		}
		if (redundantRows.size() > 0) {
			for (int x = redundantRows.size() - 1; x >= 0; x--) {
				int redundantRow = redundantRows.get(x);
				XSSFPicture imageToDelete = findImage(worksheet[0], redundantRow, ColumnsSheet0.IMAGE.ordinal());
				if (imageToDelete != null) {
					deleteCTAnchor(imageToDelete);
				}
				logger.debug("Removing (redundant) row [" + row.getRowNum() + "]:" + row.getCell(ColumnsSheet0.FILE_NAME.ordinal()) + " (" + row.getCell(ColumnsSheet0.FRIEND_NAME.ordinal()) + ")");
				worksheet[0].shiftRows(redundantRow + 1, worksheet[0].getLastRowNum(), -1);
			}
		}
		// Sheet 1: all friends (user-entered list of all their friends)
		worksheet[1] = workbook.getSheet(Sheets.SHEET1.label);
		// Remove any/all existing images on this sheet
		for (XSSFShape shape : ((XSSFDrawing)worksheet[1].createDrawingPatriarch()).getShapes()) {
			//deleteEmbeddedXSSFPicture((XSSFPicture)shape);
			deleteCTAnchor((XSSFPicture)shape);
		}
//		// Header row
//		row = worksheet[1].getRow(x);
//		for (int y = 0; y < ColumnsSheet1.values().length; y++) {
//			cell = row.getCell(y);
//			cell.setCellStyle(getHeaderRowStyleXLSX(workbook));
//			cell.setCellValue(ColumnsSheet1.values()[y].label);
//		}
		// Friend rows (user-created)
		for (int x = HEADER_ROW + 1; x < worksheet[1].getLastRowNum() + 1; x++) {
			row = worksheet[1].getRow(x);
			for (int y = 0; y < ColumnsSheet1.values().length; y++) {
				cell = row.getCell(y);
				if (cell == null) {
					// Allow for new data
					cell = row.createCell(y);
				}
				switch (y) {
					case 0: // Columns.FRIEND_NAME.ordinal():
						cell.setCellStyle(editableFriendCellStyle);
						break;
					case 1: // Columns.SEQUENCE.ordinal():
						cell.setCellFormula("INDEX('" + Sheets.SHEET0.label + "'!$B:$B,MATCH($A" + (row.getRowNum() + 1) + ",'" + Sheets.SHEET0.label + "'!$A:$A,0))");
						cell.setCellStyle(lockedFriendCellStyle);
						break;
					case 2: // Columns.FILE_NAME.ordinal():
					case 3: // Columns.IMAGE.ordinal():
					case 4: // Columns.LAST_RECEIVED.ordinal():
					case 5: // Columns.NUMBER_RECEIVED.ordinal():
						cell.setCellFormula("VLOOKUP($B" + (row.getRowNum() + 1) + ",'" + Sheets.SHEET0.label + "'!$B:$F," + y + ",FALSE)");
						cell.setCellStyle(lockedFriendCellStyle);
						if (y == 3) {
							String cellValue = getCellValue(cell, formulaEvaluator);
							if (StringUtils.isNotBlank(cellValue) && !"#N/A".equals(cellValue)) {
								Integer imagePictureId = (int)Double.parseDouble(cellValue);
								XSSFDrawing imageDrawing = (XSSFDrawing)worksheet[1].createDrawingPatriarch();
								XSSFClientAnchor imageAnchor = new XSSFClientAnchor();
								imageAnchor.setCol1(ColumnsSheet1.IMAGE.ordinal());
								imageAnchor.setCol2(ColumnsSheet1.LAST_RECEIVED.ordinal());
								imageAnchor.setRow1(row.getRowNum());
								imageAnchor.setRow2(row.getRowNum()+1);
								imageDrawing.createPicture(imageAnchor, imagePictureId);
							}
						}
						break;
				}
			}
		}
		// [Re]Create conditional formatting to highlight rows not associated with a friend
		SheetConditionalFormatting sheetCF = worksheet[1].getSheetConditionalFormatting();
		for (int rule = 0; rule < sheetCF.getNumConditionalFormattings(); rule++) {
			sheetCF.removeConditionalFormatting(rule);
		}
		ConditionalFormattingRule ruleFormula = sheetCF.createConditionalFormattingRule("ISNA($F2)");
        ruleFormula.createPatternFormatting().setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
        BorderFormatting borderFormatting = ruleFormula.createBorderFormatting();
        borderFormatting.setBorderTop(BorderStyle.DOTTED);
        borderFormatting.setBorderRight(BorderStyle.DOTTED);
        borderFormatting.setBorderBottom(BorderStyle.DOTTED);
        borderFormatting.setBorderLeft(BorderStyle.DOTTED);
        CellRangeAddress[] rangeFormula = {
        	new CellRangeAddress(HEADER_ROW + 1, worksheet[1].getLastRowNum(), 0, ColumnsSheet1.values().length - 1)
        };
        sheetCF.addConditionalFormatting(rangeFormula, ruleFormula);
		// [Re]Create existing data validation to cover all of column A
		// @see https://thecodeshewrites.com/2020/08/11/apache-poi-excel-java-dropdown-list-dependent/
        CTWorksheet ctWorksheet = ((XSSFSheet)worksheet[1]).getCTWorksheet();
        ctWorksheet.unsetDataValidations();
//		String dvReference = "'" + Sheets.SHEET0.label + "'!$A:$A";
//		Name dvNamedArea = workbook.createName();
//		dvNamedArea.setNameName("myFriends");
//		dvNamedArea.setRefersToFormula(dvReference);
		CellRangeAddressList dvAddressList = new CellRangeAddressList(HEADER_ROW + 1, worksheet[1].getLastRowNum(), ColumnsSheet1.FRIEND_NAME.ordinal(), ColumnsSheet1.FRIEND_NAME.ordinal());
		DataValidationHelper dvHelper = worksheet[1].getDataValidationHelper();
		DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint("myFriends");
		DataValidation dataValidation = dvHelper.createValidation(dvConstraint, dvAddressList);
		dataValidation.setSuppressDropDownArrow(true);
		dataValidation.setShowPromptBox(true);
		worksheet[1].addValidationData(dataValidation);
		// Resize each sheets rows & columns
		for (int sheet = 0; sheet < Sheets.values().length; sheet++) {
			// Protecting entire sheet, however, column A cells are "unlocked" and therefore editable
			worksheet[sheet].protectSheet("");
			// Freeze the header row
			worksheet[sheet].createFreezePane(0, 1);
			// Row height (all rows except Header row)
			for (int x = 1; x < worksheet[sheet].getLastRowNum() + 1; x++) {
				worksheet[sheet].getRow(x).setHeightInPoints(imageSizePoints.height);
			}
			// Column width (all columns)
			for (int y = 0; y < (sheet == 0?ColumnsSheet0.values().length:ColumnsSheet1.values().length); y++) {
				if (sheet == Sheets.SHEET0.ordinal()) {
					worksheet[sheet].setColumnWidth(y, ColumnsSheet0.values()[y].fixedWidth);	
				} else if (sheet == Sheets.SHEET1.ordinal()) {
					worksheet[sheet].setColumnWidth(y, ColumnsSheet1.values()[y].fixedWidth);	
				}
			}
		}
		// Finalize workbook
		workbook.setActiveSheet(0);
	}

	private String getCellValue(Cell cell, FormulaEvaluator formulaEvaluator) {
		String result = null;
		if (cell != null) {
			switch (cell.getCellType()) {
				case BOOLEAN:
					result = String.valueOf(cell.getBooleanCellValue());
					break;
				case FORMULA:
					result = formulaEvaluator.evaluate(cell).formatAsString();
					break;
				case NUMERIC:
					result = String.valueOf(cell.getNumericCellValue());
					break;
				default:
					result = cell.getStringCellValue();
					break;
			}
		}
		return result;
	}

	private String getHeaderRowStyleHTML() {
		return "style='background-color:black;color:white;'";
	}
	private CellStyle getHeaderRowStyleXLSX(Workbook workbook) {
		CellStyle result = workbook.createCellStyle();
		result.setFillBackgroundColor(IndexedColors.BLACK.index);
		result.setFillForegroundColor(IndexedColors.BLACK.index);
		result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		result.setBorderBottom(BorderStyle.THIN);
		result.setBorderTop(BorderStyle.THIN);
		result.setBorderLeft(BorderStyle.THIN);
		result.setBorderRight(BorderStyle.THIN);
		Font font = workbook.createFont();
		font.setBold(true);
		font.setColor(IndexedColors.WHITE.index);
		result.setFont(font);
		return result;
	}
	private CellStyle getFriendRowStyleXLSX(Workbook workbook, boolean editable) {
		CellStyle result = workbook.createCellStyle();
		result.setWrapText(true);
		if (editable) {
			result.setLocked(false);
		} else {
			Font font = workbook.createFont();
			font.setColor(IndexedColors.GREY_50_PERCENT.index);
			result.setFont(font);
		}
		return result;
	}

	@SuppressWarnings("rawtypes")
	private static XSSFPicture findImage(Sheet sheet, int row1, int col1) {
		XSSFPicture xssfPictureToDelete = null;
		Drawing drawing = sheet.getDrawingPatriarch();
		if (drawing instanceof XSSFDrawing) {
			for (XSSFShape shape : ((XSSFDrawing) drawing).getShapes()) {
				if (shape instanceof XSSFPicture) {
					XSSFPicture xssfPicture = (XSSFPicture) shape;
					String shapename = xssfPicture.getShapeName();
					int row = xssfPicture.getClientAnchor().getRow1();
					int col = xssfPicture.getClientAnchor().getCol1();

					if (row == row1 && col == col1) {
						xssfPictureToDelete = xssfPicture;
						logger.debug("Sheet " + sheet.getSheetName() + " contains the matching shape " + shape + " (" + shapename + ")"
								+ " located at row " + row + " and column " + col + "!");
						break;
					} else {
						logger.trace("Sheet " + sheet.getSheetName() + " contains a shape " + shape + " (" + shapename + ")"
								+ " located at row " + row + " and column " + col + ".");
					}
				}
			}
		}
		return xssfPictureToDelete;
	}

	private static void deleteCTAnchor(XSSFPicture xssfPicture) {
		XSSFDrawing drawing = xssfPicture.getDrawing();
		XmlCursor cursor = xssfPicture.getCTPicture().newCursor();
		cursor.toParent();
		if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor) {
			for (int i = 0; i < drawing.getCTDrawing().getTwoCellAnchorList().size(); i++) {
				if (cursor.getObject().equals(drawing.getCTDrawing().getTwoCellAnchorArray(i))) {
					drawing.getCTDrawing().removeTwoCellAnchor(i);
					logger.debug("TwoCellAnchor for picture " + xssfPicture + " was deleted.");
				}
			}
		} else if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor) {
			for (int i = 0; i < drawing.getCTDrawing().getOneCellAnchorList().size(); i++) {
				if (cursor.getObject().equals(drawing.getCTDrawing().getOneCellAnchorArray(i))) {
					drawing.getCTDrawing().removeOneCellAnchor(i);
					logger.debug("OneCellAnchor for picture " + xssfPicture + " was deleted.");
				}
			}
		} else if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor) {
			for (int i = 0; i < drawing.getCTDrawing().getAbsoluteAnchorList().size(); i++) {
				if (cursor.getObject().equals(drawing.getCTDrawing().getAbsoluteAnchorArray(i))) {
					drawing.getCTDrawing().removeAbsoluteAnchor(i);
					logger.debug("AbsoluteAnchor for picture " + xssfPicture + " was deleted.");
				}
			}
		}
	}

//	private static void deleteEmbeddedXSSFPicture(XSSFPicture xssfPicture) {
//		if (xssfPicture.getCTPicture().getBlipFill() != null) {
//			if (xssfPicture.getCTPicture().getBlipFill().getBlip() != null) {
//				if (xssfPicture.getCTPicture().getBlipFill().getBlip().getEmbed() != null) {
//					String rId = xssfPicture.getCTPicture().getBlipFill().getBlip().getEmbed();
//					XSSFDrawing drawing = xssfPicture.getDrawing();
//					drawing.getPackagePart().removeRelationship(rId);
//					drawing.getPackagePart().getPackage()
//							.deletePartRecursive(drawing.getRelationById(rId).getPackagePart().getPartName());
//					logger.debug("Picture " + xssfPicture + " was deleted.");
//				}
//			}
//		}
//	}

//	private static void deleteHSSFShape(HSSFShape shape) {
//		HSSFPatriarch drawing = shape.getPatriarch();
//		drawing.removeShape(shape);
//		logger.debug("Shape " + shape + " was deleted.");
//	}
}
