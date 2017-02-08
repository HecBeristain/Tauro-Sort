/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Model;

import java.io.*;
import java.util.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 *
 * @author HectorBeristainBermudez
 */
public class ExcelModel {
    Workbook innerWorkbook;
    
    public String ImportFile(File fileImported, JTable dataTable){
        String importResult = "System couldn't import the file";
        
        DefaultTableModel loadModel = new DefaultTableModel();
        dataTable.setModel(loadModel);
        dataTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        
        try{
            innerWorkbook = WorkbookFactory.create(new FileInputStream(fileImported));
            Sheet sheetTemporal = innerWorkbook.getSheetAt(0);
            Iterator rowIterator = sheetTemporal.rowIterator();
            
            int rowIndex = -1;
            while(rowIterator.hasNext()){
                rowIndex++;
                
                Row temporalRow = (Row) rowIterator.next();
                Iterator columnIterator = temporalRow.cellIterator();
                
                Object[] columnList = new Object[12];
                
                int columnIndex = -1;            
                while(columnIterator.hasNext()){
                    columnIndex++;
                    Cell temporalCell = (Cell) columnIterator.next();
                    if(rowIndex==0) loadModel.addColumn(temporalCell.getStringCellValue());
                    else {
                        if(temporalCell!=null){
                            switch(temporalCell.getCellType()){
                                case Cell.CELL_TYPE_NUMERIC:
                                    columnList[columnIndex]=(int)Math.round(temporalCell.getNumericCellValue());
                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    columnList[columnIndex]=temporalCell.getStringCellValue();
                                    break;       
                                default:
                                    columnList[columnIndex]=temporalCell.getDateCellValue();
                                    break;
                            }//End Switch/Case get.CellType
                        }//End temporalCell!=null condition
                    }//End Else rowIndex==0
                } //End columnIterator
                if (rowIndex!=0) loadModel.addRow(columnList);
            }//End rowIterator
            
            importResult = "Successful Import";
            
        } catch (Exception e){ }
        return importResult;
        
    }//End ImportFile Method
    
    public String ExportFile(File fileExported, JTable dataTable){
        String exportResult = "System couldn't export the file";
        
        int rowID = dataTable.getRowCount(), columnID = dataTable.getColumnCount();
        
        if(fileExported.getName().endsWith(".xls")) innerWorkbook = new HSSFWorkbook();
        else innerWorkbook = new XSSFWorkbook();
        
        Sheet sheetTemporal = innerWorkbook.createSheet("Test");
        
        try{
            for(int i=-1; i<rowID; i++){
                Row temporalRow = sheetTemporal.createRow(i+1);
                for(int j=0; j<rowID; j++){
                    Cell temporalCell = temporalRow.createCell(j);
                    
                    if(i==-1) temporalCell.setCellValue(String.valueOf(dataTable.getColumnName(j)));
                    else temporalCell.setCellValue(String.valueOf(dataTable.getValueAt(i,j)));
                    
                    innerWorkbook.write(new FileOutputStream(fileExported));
                }//End column for    
            }//End row for
            
            exportResult = "Succesful Export";
        } catch (Exception e){ }
        
        return exportResult;
    }//End ExportFile Method
    
}//End ExcelModel Class
