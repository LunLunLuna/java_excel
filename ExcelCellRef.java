package com.zenithst.common.util.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.NumberToTextConverter;
 
public class ExcelCellRef {
    /**
     * Cell에 해당하는 Column Name을 가젼온다(A,B,C..)
     * 만약 Cell이 Null이라면 int cellIndex의 값으로
     * Column Name을 가져온다.
     * @param cell
     * @param cellIndex
     * @return
     */
    public static String getName(Cell cell, int cellIndex) {
        int cellNum = 0;
        if(cell != null) {
            cellNum = cell.getColumnIndex();
        }
        else {
            cellNum = cellIndex;
        }
        
        return CellReference.convertNumToColString(cellNum);
    }
    
    public static String getValue(Cell cell) {
        String value = "";
        
        if(cell == null) {
            value = "";
        }
        else {
            if( cell.getCellType() == CellType.FORMULA ) {
                value = cell.getCellFormula();
            }
            else if( cell.getCellType() == CellType.NUMERIC ) {
            	//숫자 데이터가 텍스트 서식으로 입력되어서 숫자형으로 판단해서 읽어 들이나. double 형으로 읽어 들이는 문제가 있음.
            	//제공 유틸로 문자형으로 변환
            	value = NumberToTextConverter.toText(cell.getNumericCellValue());
            }
            else if( cell.getCellType() == CellType.STRING ) {
                value = cell.getStringCellValue();
            }
            else if( cell.getCellType() == CellType.BOOLEAN ) {
                value = cell.getBooleanCellValue() + "";
            }
            else if( cell.getCellType() == CellType.ERROR ) {
                value = cell.getErrorCellValue() + "";
            }
            else if( cell.getCellType() == CellType.BLANK ) {
                value = "";
            }
            else {
                value = cell.getStringCellValue();
            }
        }
        
        return value;
    }
 
}