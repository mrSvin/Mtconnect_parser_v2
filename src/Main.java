
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;

import org.jsoup.select.Elements;


import javax.imageio.ImageIO;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;
import java.util.function.UnaryOperator;



public class Main {
    private static int counter=0;
    static int uf5220_procent_work;
    static int uf5220_procent_pause;
    static int uf5220_procent_off;
    public static String name_data;
    public static String old_name_data;

    //Метод для сранения строк имени
    public boolean equals(Main name) {
        return this.name_data ==  name.old_name_data;
    }

    public static void main(String[] args) throws IOException, InterruptedException {
        while (counter<1) {
            org.jsoup.nodes.Document doc = Jsoup.connect("http://192.168.3.150:5000/uf5220/current").get();
            //System.out.println(doc);
            Elements element = doc.getElementsByTag("Events");
            //System.out.println(element);
            Elements element1 = doc.getElementsByTag("Execution");
            String status_work = element1.text();   //Статус оборудования

            Elements element2 = doc.getElementsByTag("SpindleSpeed");
            String status_cor_speed = element2.text();   //Коррекция скорости подачи
            //System.out.println(status_cor_speed);


            long unixTime = System.currentTimeMillis() / 1000L + 10800; //Определяем текущее время
            Date date = new java.util.Date(unixTime*1000L);
            SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            sdf.setTimeZone(java.util.TimeZone.getTimeZone("GMT"));
            String formattedDate = sdf.format(date);

            SimpleDateFormat sdf2 = new java.text.SimpleDateFormat("yyyy-MM-dd");
            sdf.setTimeZone(java.util.TimeZone.getTimeZone("GMT"));
            name_data = sdf2.format(date);  //даем название для таблицы за текущий день
            //System.out.println(name_data);

            //System.out.println(name_data.equals(old_name_data));

            if (status_work.equals("ACTIVE")) {
                uf5220_procent_work=uf5220_procent_work+1;
            }
            if (status_work.equals("UNAVAILABLE")) {
                uf5220_procent_off=uf5220_procent_off+1;
            }
            if ((status_work.equals("ACTIVE")==false) && (status_work.equals("UNAVAILABLE")==false)) {
                uf5220_procent_pause=uf5220_procent_pause+1;
            }


            //создаем единожды за сутки таблицу для заполения данными
            if (name_data.equals(old_name_data) == false) {
                Workbook wb = new XSSFWorkbook();
                // записываем созданный в памяти Excel документ в файл
                FileOutputStream out = new FileOutputStream("D:\\java\\exel\\uf5220_data\\" + name_data + ".xlsx");
                Sheet sheet0=wb.createSheet("Zagruzka");
                wb.write(out);
                System.out.println(name_data + " Excel файл успешно создан!" );
                wb.close();

                old_name_data =name_data;
            }

            //Записываем текущие значение в таблицу
            Workbook wb = new XSSFWorkbook();
            FileOutputStream out = new FileOutputStream("D:\\java\\exel\\uf5220_data\\" + name_data + ".xlsx");
            Sheet sheet0=wb.createSheet("Zagruzka");
            Row row = sheet0.createRow(0);


            Cell cell =row.createCell(0);
            cell.setCellValue(uf5220_procent_work);

            Cell cell1 =row.createCell(1);
            cell1.setCellValue(uf5220_procent_pause);

            Cell cell2 =row.createCell(2);
            cell2.setCellValue(uf5220_procent_off);

            wb.write(out);
            wb.close();

            System.out.println(status_work + " " + formattedDate + " " + uf5220_procent_work + " " + uf5220_procent_pause + " " + uf5220_procent_off);
            //System.out.println(element1.attr("sequence"));
            TimeUnit.SECONDS.sleep(5);
        }
    }




}