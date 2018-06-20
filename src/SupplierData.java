import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.Number;
import jxl.write.DateTime;
import java.util.*;


public class SupplierData {

    Queue<String> suppliersNeed = new LinkedList();

    public static void main(String[] args){
        File myDataBase_1 = new File("../../SupplierData/PO Receiving 20161001-20170930.xls");
        File myDataBase_2 = new File("../../SupplierData/PO Receiving 20170606-20180606.xls");
        processData(myDataBase_1);
        processData(myDataBase_2);
        outputData();
    }

    public static void processData(File file2Process){
        try {
            InputStream is = new FileInputStream(file2Process.getAbsolutePath());
            Workbook wb = Workbook.getWorkbook(is);
            int sheet_size = wb.getNumberOfSheets();
            Sheet dataSheet = null;
            for (int index = 0; index < sheet_size; index++) {
                if(wb.getSheet(index).getName().equals("Data")){
                    dataSheet = wb.getSheet(index);
                }
            }
            for(int i = 0; i < dataSheet.getRows(); i++){
                String suppliers = dataSheet.getCell(6, i).getContents();
                for(int j = 0; j < ignoreSuppliers.length; j++){
                    if(ignoreSuppliers[j].equals(suppliers)){
                        usefulData.add(dataSheet.getRow(i));
                    }
                }
                //System.out.println(suppliers);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void outputData(){

    }
}
