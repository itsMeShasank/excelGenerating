package com.cse.excelgenerating;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;

import android.os.Environment;
import android.provider.Settings;
import android.widget.Button;
import android.widget.Toast;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;

public class MainActivity extends AppCompatActivity {


    Button generate,design;
    ArrayList<ArrayList<String>> data;
    File filepath = new File(Environment.getExternalStorageDirectory()+"/Documents/GeneratedExcel.xls");//xlsx

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        askPermissionAndBrowseFile();


        data = new ArrayList<>();
        ArrayList<String> monlist = new ArrayList<>();
        monlist.add("mon");
        monlist.add("WT");
        monlist.add("WT");
        monlist.add("ML");
        monlist.add("BDA");
        monlist.add("LIBRARY");
        monlist.add("SCIRP");
        monlist.add("SCIRP");

        data.add(monlist);

        ArrayList<String> tuelist = new ArrayList<>();
        tuelist.add("tue");
        tuelist.add("ml");
        tuelist.add("ml");
        tuelist.add("library");
        tuelist.add("BDA");
        tuelist.add("cc");
        tuelist.add("library");
        tuelist.add("library");

        data.add(tuelist);

        ActivityCompat.requestPermissions(this,new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE,
        Manifest.permission.READ_EXTERNAL_STORAGE}, PackageManager.PERMISSION_GRANTED);

        generate = findViewById(R.id.gen);
        generate.setOnClickListener(view -> {

            generateExcel(data);

        });

        design = findViewById(R.id.announcement);
        design.setOnClickListener(view ->{
            startActivity(new Intent(getApplicationContext(),Announcement.class));
            finish();
        });


    }

    private void askPermissionAndBrowseFile() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
            if (Environment.isExternalStorageManager()){

                // If you don't have access, launch a new activity to show the user the system's dialog
                // to allow access to the external storage
            }else{
                Intent intent = new Intent();
                intent.setAction(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION);
                Uri uri = Uri.fromParts("package", this.getPackageName(), null);
                intent.setData(uri);
                startActivity(intent);
            }
        }
    }

    public void generateExcel(ArrayList<ArrayList<String>> data) {

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("IV");
        int i=0,j=0,rownum=2,cellnum=0;
        Row row;
        Cell cell;

        while(i<data.size()) {

            row = sheet.createRow(rownum);
            while(j<data.get(i).size()) {
                cell = row.createCell(cellnum);
                cell.setCellValue(data.get(i).get(j));
                cellnum+=1;
                j+=1;
            }
            rownum+=1;
            cellnum=0;
            i+=1;
            j=0;
        }

        /*while(i < data.size()) {
            System.out.println(data.get(i));

            row = sheet.createRow(rownum);

            cell = row.createCell(cellnum);
            cell.setCellValue(data.get(i).get(j));

            cell = row.createCell(1);
            cell.setCellValue(data.get(i).get(j+1));

            cell = row.createCell(2);
            cell.setCellValue(data.get(i).get(j+2));

            cell = row.createCell(3);
            cell.setCellValue(data.get(i).get(j+3));

            cell = row.createCell(4);
            cell.setCellValue(data.get(i).get(j+4));

            cell = row.createCell(5);
            cell.setCellValue(data.get(i).get(j+5));

            cell = row.createCell(6);
            cell.setCellValue(data.get(i).get(j+6));

            cell = row.createCell(7);
            cell.setCellValue(data.get(i).get(j+7));

            i+=1;
            j=0;
            rownum+=1;



        }*/

        try {
            if(!filepath.exists()) {
                filepath.createNewFile();
            }

            FileOutputStream fileOutputStream = new FileOutputStream(filepath);

            workbook.write(fileOutputStream);

            if(fileOutputStream!=null) {
                Toast.makeText(getApplicationContext(),"data saved",Toast.LENGTH_LONG).show();
                fileOutputStream.flush();
                fileOutputStream.close();
            }
        } catch(Exception e) {
            e.printStackTrace();
        }


    }

}