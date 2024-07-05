package com.zenithst.common.util.excel;

import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelUtil {
	/**
	 * 엑셀 Title 스타일
	 * 
	 * @param request
	 * @return
	 */
	public static HSSFCellStyle getTitleCellStyle(HSSFWorkbook wb)  throws Exception {
		HSSFCellStyle titleStyle = wb.createCellStyle();		
		HSSFFont titleFont = wb.createFont();
		//titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		titleFont.setFontHeight((short) 300);
		titleStyle.setFont(titleFont);
		titleStyle.setAlignment(HorizontalAlignment.CENTER);
		titleStyle.setBorderTop(BorderStyle.THIN);
		titleStyle.setBorderLeft(BorderStyle.THIN);
		titleStyle.setBorderRight(BorderStyle.THIN);
		titleStyle.setBorderBottom(BorderStyle.THIN);
		titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		titleStyle.setFillForegroundColor(HSSFColorPredefined.LIGHT_YELLOW.getIndex());
		return titleStyle;
	}

	/**
	 * 엑셀 Header 스타일
	 * 
	 * @param request
	 * @return
	 */
	public static HSSFCellStyle getHeaderCellStyle(HSSFWorkbook wb)  throws Exception {
		HSSFCellStyle headerStyle = wb.createCellStyle();
		HSSFFont headerFont = wb.createFont();
		//headerFont.setBoldweight(headerFont.BOLDWEIGHT_BOLD);
		headerStyle.setFont(headerFont);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setBorderTop(BorderStyle.THIN);
		headerStyle.setBorderLeft(BorderStyle.THIN);
		headerStyle.setBorderRight(BorderStyle.THIN);
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		headerStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
		return headerStyle;
	}
	
	/**
	 * 엑셀 Body 스타일
	 * 
	 * @param request
	 * @return
	 */
	public static HSSFCellStyle getBodyCellStyleCenter(HSSFWorkbook wb)  throws Exception {
		HSSFCellStyle bodyStyle = wb.createCellStyle();
		bodyStyle.setAlignment(HorizontalAlignment.CENTER);
		bodyStyle.setBorderTop(BorderStyle.THIN);
		bodyStyle.setBorderLeft(BorderStyle.THIN);
		bodyStyle.setBorderRight(BorderStyle.THIN);
		bodyStyle.setBorderBottom(BorderStyle.THIN);
		return bodyStyle;
	}

	/**
	 * 엑셀 Body 스타일
	 * 
	 * @param request
	 * @return
	 */
	public static HSSFCellStyle getBodyCellStyleLeft(HSSFWorkbook wb)  throws Exception {
		HSSFCellStyle bodyStyle = wb.createCellStyle();
		bodyStyle.setAlignment(HorizontalAlignment.LEFT);
		bodyStyle.setBorderTop(BorderStyle.THIN);
		bodyStyle.setBorderLeft(BorderStyle.THIN);
		bodyStyle.setBorderRight(BorderStyle.THIN);
		bodyStyle.setBorderBottom(BorderStyle.THIN);
		return bodyStyle;
	}
	
	/**
	 * 엑셀 Body 스타일
	 * 
	 * @param request
	 * @return
	 */
	public static HSSFCellStyle getBodyCellStyleRight(HSSFWorkbook wb)  throws Exception {
		HSSFCellStyle bodyStyle = wb.createCellStyle();
		bodyStyle.setAlignment(HorizontalAlignment.RIGHT);
		bodyStyle.setBorderTop(BorderStyle.THIN);
		bodyStyle.setBorderLeft(BorderStyle.THIN);
		bodyStyle.setBorderRight(BorderStyle.THIN);
		bodyStyle.setBorderBottom(BorderStyle.THIN);
		return bodyStyle;
	}

	/**
	 * 브라우저 구분 얻기.
	 * 
	 * @param request
	 * @return
	 */
	private static String getBrowser(HttpServletRequest request) {
		String header = request.getHeader("User-Agent");
		if (header.indexOf("MSIE") > -1) {
			return "MSIE";
		} else if (header.indexOf("Chrome") > -1) {
			return "Chrome";
		} else if (header.indexOf("Opera") > -1) {
			return "Opera";
		}
		return "Firefox";
	}

	/**
	 * Disposition 지정하기.
	 * 
	 * @param filename
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	public static void setDisposition(String filename, HttpServletRequest request, HttpServletResponse response) throws Exception {
		String browser = getBrowser(request);

		String dispositionPrefix = "attachment; filename=";
		String encodedFilename = null;

		if (browser.equals("MSIE")) {
			encodedFilename = URLEncoder.encode(filename, "UTF-8").replaceAll("\\+", "%20");
		} else if (browser.equals("Firefox")) {
			encodedFilename = "\"" + new String(filename.getBytes("UTF-8"), "8859_1") + "\"";
		} else if (browser.equals("Opera")) {
			encodedFilename = "\"" + new String(filename.getBytes("UTF-8"), "8859_1") + "\"";
		} else if (browser.equals("Chrome")) {
			StringBuffer sb = new StringBuffer();
			for (int i = 0; i < filename.length(); i++) {
				char c = filename.charAt(i);
				if (c > '~') {
					sb.append(URLEncoder.encode("" + c, "UTF-8"));
				} else {
					sb.append(c);
				}
			}
			encodedFilename = sb.toString();
		} else {
			//throw new RuntimeException("Not supported browser");
			throw new IOException("Not supported browser");
		}

		response.setHeader("Content-Disposition", dispositionPrefix + encodedFilename);

		if ("Opera".equals(browser)){
			response.setContentType("application/octet-stream;charset=UTF-8");
		}
	}
	
	/**
	 * 엑셀 파일을 읽어온다.
	 * @param	ExcelReadOption.setFilePath(엑셀파일경로);
	 * @param	ExcelReadOption.setStartRow(데이터 로드 시작 row 번호);	//첫번째 row가 헤더일경우 건너 뛸수 있음
	 * @param	ExcelReadOption.setColumnNameUse(true);				//지정 컬럼 사용 여부
	 * @param	ExcelReadOption.setOutputColumns("A","B","C","D","E","F","G","H"); // 데이터를 읽어들일 컬럼명 EXCEL 상단 열명
	 * @throws IOException 
	 * */
	public static List<Map<String, String>> readBody(ExcelReadOption excelReadOption) throws IOException {
		//엑셀 파일 자체
		//엑셀파일을 읽어 들인다.
		//FileType.getWorkbook() <-- 파일의 확장자에 따라서 적절하게 가져온다.
		Workbook wb = ExcelFiletype.getWorkbook(excelReadOption.getFilePath());
		/**
		 * 엑셀 파일에서 첫번째 시트를 가지고 온다.
		 */
		Sheet sheet = wb.getSheetAt(0);

		//System.out.println("Sheet 이름: "+ wb.getSheetName(0)); 
		//System.out.println("데이터가 있는 Sheet의 수 :" + wb.getNumberOfSheets());
		/**
		 * sheet에서 유효한(데이터가 있는) 행의 개수를 가져온다.
		 */
		int numOfRows = sheet.getPhysicalNumberOfRows();
		int numOfCells = 0;

		Row row = null;
		Cell cell = null;

		String cellName = "";
		/**
		 * 각 row마다의 값을 저장할 맵 객체
		 * 저장되는 형식은 다음과 같다.
		 * put("A", "이름");
		 * put("B", "게임명");
		 */
		Map<String, String> map = null;
		/*
		 * 각 Row를 리스트에 담는다.
		 * 하나의 Row를 하나의 Map으로 표현되며
		 * List에는 모든 Row가 포함될 것이다.
		 */
		List<Map<String, String>> result = new ArrayList<Map<String, String>>(); 
		/**
		 * 각 Row만큼 반복을 한다.
		 */
		for(int rowIndex = excelReadOption.getStartRow() - 1; rowIndex < numOfRows; rowIndex++) {
			/*
			 * 워크북에서 가져온 시트에서 rowIndex에 해당하는 Row를 가져온다.
			 * 하나의 Row는 여러개의 Cell을 가진다.
			 */
			row = sheet.getRow(rowIndex);

			if(row != null) {
				/*
				 * 가져온 Row의 Cell의 개수를 구한다.
				 */
				numOfCells = row.getPhysicalNumberOfCells();
				/*
				 * 데이터를 담을 맵 객체 초기화
				 */
				map = new HashMap<String, String>();
				/*
				 * cell의 수 만큼 반복한다.
				 */
				for(int cellIndex = 0; cellIndex <= numOfCells; cellIndex++) {
					/*
					 * Row에서 CellIndex에 해당하는 Cell을 가져온다.
					 */
					cell = row.getCell(cellIndex);
					/*
					 * 현재 Cell의 이름을 가져온다
					 * 이름의 예 : A,B,C,D,......
					 */
					cellName = ExcelCellRef.getName(cell, cellIndex);

					/*
					 * 추출 대상 컬럼인지 확인한다
					 * 추출 대상 컬럼이 아니라면, 
					 * for로 다시 올라간다
					 */
					if( !excelReadOption.getOutputColumns().contains(cellName)&&excelReadOption.isColumnNameUse()) {
						continue;
					}
					/*
					 * map객체의 Cell의 이름을 키(Key)로 데이터를 담는다.
					 */
					if(excelReadOption.isColumnNameUse()) {
						map.put(cellName, ExcelCellRef.getValue(cell));
					}else {
						map.put(cellIndex+"", ExcelCellRef.getValue(cell));
					}

				}
				/*
				 * 만들어진 Map객체를 List로 넣는다.
				 */
				result.add(map);

			}

		}
		wb.close();
		return result;

	}
}
