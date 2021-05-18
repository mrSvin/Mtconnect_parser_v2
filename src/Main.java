
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;

import org.jsoup.select.Elements;


import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.TimeUnit;


public class Main {
    private static int counter=0;
    public static int kol_device=3;

    public static void main(String[] args) throws IOException, InterruptedException {
        while (counter<1) {
            //Сбор данных для УФ5220
            uf5220.uf5220_parser();
            progress.progress_parser();
            ntx1000.ntx1000_parser();
        }
    }




}