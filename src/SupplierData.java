import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;

import java.util.*;


public class SupplierData {

    public static Queue<String> suppliersNeeded = new LinkedList();
    public static String[] suppliersNeededArr;

    public static void main(String[] args){
        File myDataBase_1 = new File("../../SupplierData/PO Receiving 20161001-20170930.xls");
        File myDataBase_2 = new File("../../SupplierData/PO Receiving 20170606-20180606.xls");
        processData(myDataBase_1);
        System.out.println("Finish Data 1");
        processData(myDataBase_2);
        System.out.println("Finish Data 2");
        //suppliersNeededArr = (String[]) suppliersNeeded.toArray();
        //outputData();
        matchSuppliers();
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
                    String supplier = currDataSheet.getCell(6, i).getContents().toUpperCase();
                    if(!suppliersNeeded.contains(supplier) && !supplier.contains("BUC") && !supplier.contains("物流") &&
                            !supplier.contains("大学") && !supplier.contains("旅") && !supplier.contains("航空") &&
                            !supplier.contains("运") && !supplier.contains("测试") && !supplier.contains("检验")){
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

    public static void matchSuppliers(){
        try{
            InputStream is = new FileInputStream("../../SupplierData/SuppliersDataProcessing.xls");
            Workbook wb = Workbook.getWorkbook(is);
            WritableWorkbook myFile = Workbook.createWorkbook(new File("../../SupplierData/SuppliersDataProcessing.xls"), wb);
            //WritableSheet sheet = myFile.createSheet("Suppliers", 0);
            WritableSheet processSheet = myFile.getSheet(0);
            int rowNum = processSheet.getRows();
            WritableCellFormat noNeed= new WritableCellFormat();
            noNeed.setBackground(jxl.format.Colour.BLACK);
            System.out.println("No Problem 1");
            for(int i = 0; i < rowNum; i++){
                String engName = processSheet.getCell(1, i).getContents().toUpperCase();
                String chnName = processSheet.getCell(2, i).getContents().toUpperCase();
                //System.out.println("No Problem 2");
                if(engName.equals("") && chnName.equals("")){
                    continue;
                }
                if(!engName.equals("") && !chnName.equals("")){
                    if(suppliersNeeded.contains(engName) || suppliersNeeded.contains(chnName)){
                        if(suppliersNeeded.contains(engName)){
                            suppliersNeeded.remove(engName);
                        }
                        else{
                            suppliersNeeded.remove(chnName);
                        }
                        continue;
                    }
                    else{
                        System.out.println(engName);
                        System.out.println(chnName);
                        processSheet.getWritableCell(i, 0).setCellFormat(noNeed);
                        Label noNeedLabel = new Label(0, i, "X", noNeed);
                        processSheet.addCell(noNeedLabel);
                        //processSheet.getWritableCell(i, 1).setCellFormat(noNeed);
                        //processSheet.getWritableCell(i, 2).setCellFormat(noNeed);
                        continue;
                    }
                }
                else if(chnName.equals("")){
                    if(suppliersNeeded.contains(engName)){
                        suppliersNeeded.remove(engName);
                        continue;
                    }
                    else{
                        processSheet.getWritableCell(i, 0).setCellFormat(noNeed);
                        Label noNeedLabel = new Label(0, i, "X", noNeed);
                        processSheet.addCell(noNeedLabel);
                        //processSheet.getWritableCell(i, 1).setCellFormat(noNeed);
                        //processSheet.getWritableCell(i, 2).setCellFormat(noNeed);
                        continue;
                    }
                }
                else if(engName.equals("")){
                    if(suppliersNeeded.contains(chnName)){
                        suppliersNeeded.remove(chnName);
                    }
                    else{
                        processSheet.getWritableCell(i, 0).setCellFormat(noNeed);
                        Label noNeedLabel = new Label(0, i, "X", noNeed);
                        processSheet.addCell(noNeedLabel);
                        //processSheet.getWritableCell(i, 1).setCellFormat(noNeed);
                        //processSheet.getWritableCell(i, 2).setCellFormat(noNeed);
                        continue;
                    }
                }
            }
            while(!suppliersNeeded.isEmpty()){
                String currSupplier = suppliersNeeded.poll();
                Label currCell = new Label(1, rowNum, currSupplier);
                processSheet.addCell(currCell);
                rowNum++;
            }
            myFile.write();
            myFile.close();
            //System.out.println("Total number of suppliers: " + (rowCnt-1));
        } catch (Exception e){
            System.out.println(e);
        }
    }
}
