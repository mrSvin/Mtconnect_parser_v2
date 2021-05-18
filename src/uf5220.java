import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.select.Elements;

import javax.lang.model.element.Element;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.TimeUnit;

public class uf5220 {
    static int uf5220_procent_work;
    static int uf5220_procent_pause;
    static int uf5220_procent_off;
    public static String uf5220_name_data;
    public static String uf5220_old_name_data;

    public static String uf5220_formattedDate; // Текущее время
    public static String uf5220_formattedDate_for_hc; // Текущее время для highchar

    public static String uf5220_old_status_work="null";
    public static String uf5220_old_status_off="null";

    public static int uf5220_stroka_work;
    public static int uf5220_stroka_white;
    public static int uf5220_stroka_off;
    public static int uf5220_triger_pause;
    public static int uf5220_stroka_programname;

    public static ArrayList<String> uf5220_work_arrayList = new ArrayList<String>();
    public static ArrayList<String> uf5220_white_arrayList = new ArrayList<String>();
    public static ArrayList<String> uf5220_off_arrayList = new ArrayList<String>();
    public static ArrayList<String> uf5220_programname_arrayList = new ArrayList<String>();


    public static void uf5220_parser() throws IOException, InterruptedException {
        org.jsoup.nodes.Document doc = Jsoup.connect("http://192.168.3.41:5000/uf5220/current").get();
        //System.out.println(doc);

        //Номер инструмента
        Elements element0 = doc.getElementsByTag("ToolId");
        //Статус оборудования
        Elements element1 = doc.getElementsByTag("Execution");
        //Номер кадра
        Elements element2 = doc.getElementsByTag("Line");
        //Имя УП
        Elements element3 = doc.getElementsByTag("Program");
        //Коррекция подачи
        Elements element4 = doc.getElementsByTag("PathFeedrate");
        //Коррекция скорости
        Elements element5 = doc.getElementsByTag("SpindleSpeed");
        //Координата Х
        Elements element6 = doc.getElementsByTag("Position");
        //Координата Y
        Elements element7 = doc.getElementsByTag("Position");
        //Координата Z
        Elements element8 = doc.getElementsByTag("Position");
        //Нагрузка по Х
        Elements element9 = doc.getElementsByTag("Load");
        //Нагрузка по Y
        Elements element10 = doc.getElementsByTag("Load");
        //Нагрузка по Z
        Elements element11 = doc.getElementsByTag("Load");


        //Преобразовываем значения элементов в текст
        String toolid = element0.text();
        String status_work = element1.text();
        String nomer_kadr = element2.text();
        String name_program = element3.text();
        String kor_podach = element4.get(0).text();
        String kor_speed = element5.get(1).text();
        String pos_x = element6.get(0).text();
        String pos_y = element7.get(1).text();
        String pos_z = element8.get(2).text();
        String nagr_x = element9.get(0).text();
        String nagr_y = element10.get(1).text();
        String nagr_z = element11.get(2).text();


        long unixTime = System.currentTimeMillis() / 1000L + 10800; //Определяем текущее время
        Date date = new java.util.Date(unixTime*1000L);
        SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        SimpleDateFormat for_hc = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

        sdf.setTimeZone(java.util.TimeZone.getTimeZone("GMT"));
        for_hc.setTimeZone(java.util.TimeZone.getTimeZone("GMT"));
        //Записываем текущее время
        uf5220_formattedDate_for_hc = for_hc.format(date);
        uf5220_formattedDate = sdf.format(date);

        SimpleDateFormat sdf2 = new java.text.SimpleDateFormat("yyyy-MM-dd");
        sdf2.setTimeZone(java.util.TimeZone.getTimeZone("GMT"));
        uf5220_name_data = sdf2.format(date);  //даем название для таблицы за текущий день
        //System.out.println(name_data);

        //System.out.println(name_data.equals(old_name_data));


        //создаем единожды за сутки таблицу для заполения данными по названию даты таблицы
        if (uf5220_name_data.equals(uf5220_old_name_data) == false) {

            //Очищаем массивы и счетчики
            uf5220_work_arrayList = new ArrayList<String>();
            uf5220_white_arrayList = new ArrayList<String>();
            uf5220_off_arrayList = new ArrayList<String>();
            uf5220_programname_arrayList = new ArrayList<String>();

            uf5220_procent_work=0;
            uf5220_procent_pause=0;
            uf5220_procent_off=0;

            uf5220_stroka_work=0;
            uf5220_stroka_off=0;
            uf5220_stroka_white=0;
            uf5220_stroka_programname=0;

            uf5220_old_status_work="null";
            uf5220_old_status_off="null";
            uf5220_triger_pause=0;

            //Создаем новую таблицу
            Workbook wb = new XSSFWorkbook();
            // записываем созданный в памяти Excel документ в файл
            FileOutputStream out = new FileOutputStream("D:\\java\\exel\\uf5220_data\\" + uf5220_name_data + ".xlsx");
            Sheet sheet0=wb.createSheet("Zagruzka");
            wb.write(out);
            System.out.println(uf5220_name_data + " Excel файл успешно создан!" );
            wb.close();

            uf5220_old_name_data =uf5220_name_data;
        }


        //Счетчик для определения % загрузки оборудования
        //по программе
        if (status_work.equals("ACTIVE")) { uf5220_procent_work=uf5220_procent_work+1; }
        //выключен
        if (status_work.equals("UNAVAILABLE")) { uf5220_procent_off=uf5220_procent_off+1; }
        //ожидание
        if ((status_work.equals("ACTIVE")==false) && (status_work.equals("UNAVAILABLE")==false)) {
            uf5220_procent_pause=uf5220_procent_pause+1;
        }


        //Определяем хронологию работы оборудования фиксируя переключения состояний

        //Состояние "Работа"
        //Для статуса работы фиксируем состояние TON
        if (status_work.equals("ACTIVE") && status_work.equals(uf5220_old_status_work) == false) {
            uf5220_stroka_work=uf5220_stroka_work+1;
            uf5220_work_arrayList.add(uf5220_stroka_work-1, uf5220_formattedDate_for_hc);

            uf5220_stroka_programname=uf5220_stroka_programname+1;
            uf5220_programname_arrayList.add(uf5220_stroka_programname-1,name_program);
            System.out.println("Станок uf5220 перешел в состояние включен");
        }
        //Фиксируем статус для тригерной логики
        if (status_work.equals("ACTIVE")) {uf5220_old_status_work="ACTIVE";}

        //Для статуса работы определяем состояние TOF
        if (uf5220_old_status_work.equals("ACTIVE") && !status_work.equals("ACTIVE")) {
            uf5220_old_status_work="null";
            uf5220_stroka_work=uf5220_stroka_work+1;
            uf5220_work_arrayList.add(uf5220_stroka_work-1, uf5220_formattedDate_for_hc);
        }

        //Состояние "Выключен"
        if (status_work.equals("UNAVAILABLE") && status_work.equals(uf5220_old_status_off) == false) {
            uf5220_stroka_off=uf5220_stroka_off+1;
            uf5220_off_arrayList.add(uf5220_stroka_off-1, uf5220_formattedDate_for_hc);
            System.out.println("Станок uf5220 перешел в состояние выключен");
        }
        //Фиксируем статус для тригерной логики
        if (status_work.equals("UNAVAILABLE")) {uf5220_old_status_off="UNAVAILABLE";}

        //Для статуса работы определяем состояние TOF
        if (uf5220_old_status_off.equals("UNAVAILABLE") && !status_work.equals("UNAVAILABLE")) {
            uf5220_old_status_off="null";
            uf5220_stroka_off=uf5220_stroka_off+1;
            uf5220_off_arrayList.add(uf5220_stroka_off-1, uf5220_formattedDate_for_hc);
        }

        //Ожидание
        //TON
        if ((status_work.equals("ACTIVE")==false) && (status_work.equals("UNAVAILABLE")==false) && uf5220_triger_pause != 1) {
            uf5220_triger_pause=1;
            uf5220_stroka_white=uf5220_stroka_white+1;
            uf5220_white_arrayList.add(uf5220_stroka_white-1, uf5220_formattedDate_for_hc);
            System.out.println("Станок uf5220 перешел в состояние ожидание");
        }
        //TOF
        if (((status_work.equals("ACTIVE")==true) || (status_work.equals("UNAVAILABLE")==true)) && uf5220_procent_pause>0 && uf5220_triger_pause != 0) {
            uf5220_triger_pause=0;
            uf5220_stroka_white=uf5220_stroka_white+1;
            uf5220_white_arrayList.add(uf5220_stroka_white-1, uf5220_formattedDate_for_hc);
        }


        try {
        //Записываем текущие значение в таблицу
        Workbook wb = new XSSFWorkbook();
        FileOutputStream out = new FileOutputStream("D:\\java\\exel\\uf5220_data\\" + uf5220_name_data + ".xlsx");
        Sheet sheet0=wb.createSheet("Zagruzka");

        //Создаем строку 1
        Row row = sheet0.createRow(0);

        //Создаем и записываем в столбец 0
        Cell cell =row.createCell(0);
        cell.setCellValue(uf5220_procent_work);

        //Создаем и записываем в столбец 1
        Cell cell1 =row.createCell(1);
        cell1.setCellValue(uf5220_procent_pause);

        //Создаем и записываем в столбец 2
        Cell cell2 =row.createCell(2);
        cell2.setCellValue(uf5220_procent_off);


        //Создаем строку 2 для заполнения массивом "работа"
        row = sheet0.createRow(1);
        //определяем длину массива
        int length_array_work = uf5220_work_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_work; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = uf5220_work_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        //Создаем строку 3 для заполнения массивом "ожидание"
        row = sheet0.createRow(2);
        //определяем длину массива
        int length_array_white = uf5220_white_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_white; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = uf5220_white_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        //Создаем строку 4 для заполнения массивом "выключен"
        row = sheet0.createRow(3);
        //определяем длину массива
        int length_array_off = uf5220_off_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_off; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = uf5220_off_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        //Создаем строку 5 для заполнения массивом "имя программы"
        row = sheet0.createRow(4);
        int length_array_nameprogram = uf5220_programname_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_nameprogram; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = uf5220_programname_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        wb.write(out);
        wb.close();

        } catch (FileNotFoundException e) {
            System.out.println("Произошла ошибка при создании файла uf5220");
        }

        System.out.println(" УФ5220: " + status_work + " " + uf5220_formattedDate + " " + uf5220_procent_work + " " + uf5220_procent_pause + " " + uf5220_procent_off);
        TimeUnit.MILLISECONDS.sleep(5000/ Main.kol_device);

        //Записываем текущие значение в таблицу
        Workbook wb2 = new XSSFWorkbook();
        FileOutputStream out2 = new FileOutputStream("D:\\java\\exel\\uf5220_realtime_data.xlsx");
        Sheet sheet1=wb2.createSheet("Realtime");

        Row row3 = sheet1.createRow(0);

        //Создаем и записываем в столбец 1
        Cell cell3 =row3.createCell(0);
        cell3.setCellValue(toolid);

        Cell cell4 =row3.createCell(1);
        cell4.setCellValue(status_work);

        Cell cell5 =row3.createCell(2);
        cell5.setCellValue(nomer_kadr);

        Cell cell6 =row3.createCell(3);
        cell6.setCellValue(name_program);

        Cell cell7 =row3.createCell(4);
        cell7.setCellValue(kor_podach);

        Cell cell8 =row3.createCell(5);
        cell8.setCellValue(kor_speed);

        Cell cell9 =row3.createCell(6);
        cell9.setCellValue(pos_x);

        Cell cell10 =row3.createCell(7);
        cell10.setCellValue(pos_y);

        Cell cell11 =row3.createCell(8);
        cell11.setCellValue(pos_z);

        Cell cell12 =row3.createCell(9);
        cell12.setCellValue(nagr_x);

        Cell cell13 =row3.createCell(10);
        cell13.setCellValue(nagr_y);

        Cell cell14 =row3.createCell(11);
        cell14.setCellValue(nagr_z);


        wb2.write(out2);
        wb2.close();

    }
}
