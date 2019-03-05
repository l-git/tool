package com.company.tool.excel;


import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;


public class WriteExcel {
    public static void main(String[] args) throws Exception {


        JSONArray jsonArray=new JSONArray();

        for(int i=0;i<10;i++){

            JSONObject jsonObject=new JSONObject();
            jsonObject.put("c1",i);
            jsonObject.put("c2",i);
            jsonObject.put("c3",i);
            jsonArray.add(jsonObject);
        }





        //Create blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet spreadsheet = workbook.createSheet(" Employee Info ");

        //Create row object
        XSSFRow row;

        //This data needs to be written (Object[])
        Map<String, Object[]> empinfo = new TreeMap<String, Object[]>();
        empinfo.put("1", new Object[]{
                "EMP ID", "EMP NAME", "DESIGNATION"});

        empinfo.put("2", new Object[]{
                "tp01", "Gopal", "Technical Manager"});

        empinfo.put("3", new Object[]{
                "tp02", "Manisha", "Proof Reader"});

        empinfo.put("4", new Object[]{
                "tp03", "Masthan", "Technical Writer"});

        empinfo.put("5", new Object[]{
                "tp04", "Satish", "Technical Writer"});

        empinfo.put("6", new Object[]{
                "tp05", "Krishna", "Technical Writer"});

        //Iterate over data and write to sheet
        Set<String> keyid = empinfo.keySet();
        int rowid = 0;

        for (String key : keyid) {
            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = empinfo.get(key);
            int cellid = 0;

            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String) obj);
            }
        }


        for(int i=0;i<jsonArray.size();i++){

            JSONObject jsonObject=jsonArray.getJSONObject(i);

            row = spreadsheet.createRow(rowid++);
            int cellid = 0;
            for(int j=1;j<4;j++){
                Cell cell = row.createCell(cellid++);

                int k=jsonObject.getInteger("c"+j);
                cell.setCellValue(k);
            }



        }

        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(
                new File("e:/temp/Writesheet.xlsx"));

        workbook.write(out);
        out.close();
        System.out.println("Writesheet.xlsx written successfully");
    }
}



