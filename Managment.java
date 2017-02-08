/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tauro;

import View.ExcelView;
import Model.ExcelModel;
import Controller.ExcelController;
/**
 *
 * @author HectorBeristainBermudez
 */
public class Managment {
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        ExcelModel managmentModel = new ExcelModel();
        ExcelView managmentView = new ExcelView();
        ExcelController managmentController = new ExcelController(managmentView,managmentModel);
        managmentView.setVisible(true);
        managmentView.setLocationRelativeTo(null);
        
        //managmentView.dataTable.print();
        //Generate Package with last specifications
        //cottonPackage lastPackage = new cottonPackage();
    
    }//End Managment Main
    
}//End Managment Class
