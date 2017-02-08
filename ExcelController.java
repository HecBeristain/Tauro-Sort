/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Controller;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import View.ExcelView;
import Model.ExcelModel;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 *
 * @author HectorBeristainBermudez
 */
public class ExcelController implements ActionListener{
    ExcelModel temporalModel = new ExcelModel();
    ExcelView temporalView = new ExcelView();
    
    JFileChooser fileSelecter = new JFileChooser();
    File localFile;
    
    int actionCounter=0;
    
    public ExcelController(ExcelView temporalView, ExcelModel temporalModel){
        this.temporalModel=temporalModel;
        this.temporalView=temporalView;
        
        this.temporalView.importButton.addActionListener(this);
        this.temporalView.exportButton.addActionListener(this);
    }
    
    public void AddFilter1(){
        fileSelecter.setFileFilter(new FileNameExtensionFilter("Excel (*.xls)","xls"));
        fileSelecter.setFileFilter(new FileNameExtensionFilter("Excel (*.xlsx)","xlsx"));
    }
    
    @Override
    public void actionPerformed(ActionEvent e) {
        actionCounter++;
        if(actionCounter==1)AddFilter1();
        
        if(e.getSource()==temporalView.importButton){
            if(fileSelecter.showDialog(null, "Select File")==JFileChooser.APPROVE_OPTION){
                localFile=fileSelecter.getSelectedFile();
                if(localFile.getName().endsWith("xls") || localFile.getName().endsWith("xlsx")){
                    JOptionPane.showMessageDialog(null,temporalModel.ImportFile(localFile, temporalView.dataTable));
                } else {
                    JOptionPane.showMessageDialog(null,"Invalid format, choose another");
                }//End localFile.getName().endsWith
            }//End fileSelecter.showDialog(null, "Select File")
        }//End e.getSource()==temporalView.importButton

        if(e.getSource()==temporalView.exportButton){
            if(fileSelecter.showDialog(null, "Export File")==JFileChooser.APPROVE_OPTION){
                localFile=fileSelecter.getSelectedFile();
                if(localFile.getName().endsWith("xls")||localFile.getName().endsWith("xlsx")){
                    JOptionPane.showMessageDialog(null,temporalModel.ExportFile(localFile, temporalView.dataTable));
                } else {
                    JOptionPane.showMessageDialog(null,"Invalid format, choose another");
                }//End localFile.getName().endsWith
            }//End fileSelecter.showDialog(null, "Select File")
        }//End e.getSource()==temporalView.importButton
        
    }//End ActionPerformed Method
    
}//End ExcelController Class
