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

    public static Queue<String> suppliersNeeded = new LinkedList();

    public static void main(String[] args){
        File myDataBase_1 = new File("../../SupplierData/PO Receiving 20161001-20170930.xls");
        File myDataBase_2 = new File("../../SupplierData/PO Receiving 20170606-20180606.xls");
        processData(myDataBase_1);
        processData(myDataBase_2);
        outputData();
    }

    public static void processData(File file2Process){
        try{
            InputStream is = new FileInputStream(file2Process.getAbsolutePath());
            Workbook wb = Workbook.getWorkbook(is);
            int sheet_size = wb.getNumberOfSheets();
            Queue<Sheet> dataSheets = new LinkedList();
            for (int index = 0; index < sheet_size; index++){
                Sheet dataSheet;
                if(wb.getSheet(index).getName().contains("Data")){
                    dataSheet = wb.getSheet(index);
                    dataSheets.add(dataSheet);
                }
            }
            while(!dataSheets.isEmpty()){
                Sheet currDataSheet = dataSheets.poll();
                for(int i = 0; i < currDataSheet.getRows(); i++){
                    String supplier = currDataSheet.getCell(6, i).getContents();
                    if(!suppliersNeeded.contains(supplier)){
                        suppliersNeeded.add(supplier);
                    }
                }
            }
        } catch (FileNotFoundException e){
            e.printStackTrace();
        } catch (BiffException e){
            e.printStackTrace();
        } catch (IOException e){
            e.printStackTrace();
        }
    }

    public static void outputData(){
        try{
            WritableWorkbook myFile = Workbook.createWorkbook(new File("../../SupplierData/usefulSuppliersData.xls"));
            WritableSheet sheet = myFile.createSheet("Suppliers", 0);
            int rowCnt = 0;
            //Label title = new Label(0, 0, "Suppliers Data");
            //sheet.addCell(title);
            while(!suppliersNeeded.isEmpty()){
                String currSupplier = suppliersNeeded.poll();
                Label currCell = new Label(0, rowCnt, currSupplier);
                sheet.addCell(currCell);
                rowCnt++;
            }
            myFile.write();
            myFile.close();
            System.out.println("Total number of suppliers: " + (rowCnt-1));
        } catch (Exception e){
            System.out.println(e);
        }
    }
}
