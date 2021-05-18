import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.NodeList;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.TimeUnit;

public class ntx1000 {
    static int ntx1000_procent_work;
    static int ntx1000_procent_pause;
    static int ntx1000_procent_off;
    static int ntx1000_procent_avar;
    static int ntx1000_procent_nagruzka;
    public static String ntx1000_name_data;
    public static String ntx1000_old_name_data;

    public static String ntx1000_formattedDate; // Текущее время
    public static String ntx1000_formattedDate_for_hc; // Текущее время для highchar

    public static String ntx1000_old_status_work="null";
    public static String ntx1000_old_status_off="null";

    public static int ntx1000_stroka_work=0;
    public static int ntx1000_stroka_white=0;
    public static int ntx1000_stroka_off=0;
    public static int ntx1000_stroka_error=0;
    public static int ntx1000_stroka_nagruzka=0;
    public static int ntx1000_stroka_programname=0;

    public static int ntx1000_triger_pause=0;
    public static int ntx1000_triger_error=0;
    public static int ntx1000_triger_nagruzka=0;

    public static int ntx1000_sost_nagruzka=0;
    public static int ntx1000_sost_error=0;


    public static ArrayList<String> ntx1000_work_arrayList = new ArrayList<String>();
    public static ArrayList<String> ntx1000_white_arrayList = new ArrayList<String>();
    public static ArrayList<String> ntx1000_off_arrayList = new ArrayList<String>();
    public static ArrayList<String> ntx1000_avar_arrayList = new ArrayList<String>();
    public static ArrayList<String> ntx1000_nagruzka_arrayList = new ArrayList<String>();
    public static ArrayList<String> ntx1000_programname_arrayList = new ArrayList<String>();


    public static void ntx1000_parser() throws IOException, InterruptedException {
       Document doc = Jsoup.connect("http://192.168.3.150:5000/NTX1000/current").get();

        //Номер инструмента
        Elements element0 = doc.getElementsByTag("ToolId");
        //Статус оборудования
        Elements element1 = doc.getElementsByTag("Execution");
        //Номер кадра
        Elements element2 = doc.getElementsByTag("Line");
        //Имя УП
        Elements element3 = doc.getElementsByTag("Program");
        //Коррекция подачи
        Elements element4 = doc.getElementsByTag("JogOverride");
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
        //Нагрузка на шпинделе
        Elements element12 = doc.getElementsByTag("Load");
        //Состоние Аварии
        Elements element13 = doc.getElementsByTag("EmergencyStop");

        //Преобразовываем значения элементов в текст
        String toolid = element0.get(3).text();
        String status_work = element1.get(3).text();
        String nomer_kadr = element2.get(3).text();
        String name_program = element3.get(3).text();
        String kor_podach = element4.get(3).text();
        String kor_speed = element5.get(1).text();
        String pos_x = element6.get(8).text();
        String pos_y = element7.get(12).text();
        String pos_z = element8.get(17).text();
        String nagr_x = element9.get(11).text();
        String nagr_y = element10.get(15).text();
        String nagr_z = element11.get(20).text();
        String nagr_hpindel = element12.get(7).text();
        String avar_sost = element13.text();


        long unixTime = System.currentTimeMillis() / 1000L + 10800; //Определяем текущее время
        Date date = new java.util.Date(unixTime*1000L);
        SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        SimpleDateFormat for_hc = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

        sdf.setTimeZone(java.util.TimeZone.getTimeZone("GMT"));
        for_hc.setTimeZone(java.util.TimeZone.getTimeZone("GMT"));
        //Записываем текущее время
        ntx1000_formattedDate_for_hc = for_hc.format(date);
        ntx1000_formattedDate = sdf.format(date);

        SimpleDateFormat sdf2 = new java.text.SimpleDateFormat("yyyy-MM-dd");
        sdf2.setTimeZone(java.util.TimeZone.getTimeZone("GMT"));
        ntx1000_name_data = sdf2.format(date);  //даем название для таблицы за текущий день
        //System.out.println(name_data);

        //System.out.println(name_data.equals(old_name_data));


        //создаем единожды за сутки таблицу для заполения данными по названию даты таблицы
        if (ntx1000_name_data.equals(ntx1000_old_name_data) == false) {

            //Очищаем массивы и счетчики
            ntx1000_work_arrayList = new ArrayList<String>();
            ntx1000_white_arrayList = new ArrayList<String>();
            ntx1000_off_arrayList = new ArrayList<String>();
            ntx1000_avar_arrayList = new ArrayList<String>();
            ntx1000_nagruzka_arrayList = new ArrayList<String>();
            ntx1000_programname_arrayList = new ArrayList<String>();

            ntx1000_procent_work=0;
            ntx1000_procent_pause=0;
            ntx1000_procent_off=0;
            ntx1000_procent_nagruzka=0;
            ntx1000_procent_avar=0;

            ntx1000_stroka_work=0;
            ntx1000_stroka_off=0;
            ntx1000_stroka_white=0;
            ntx1000_stroka_error=0;
            ntx1000_stroka_nagruzka=0;
            ntx1000_stroka_programname=0;

            ntx1000_old_status_work="null";
            ntx1000_old_status_off="null";
            ntx1000_triger_pause=0;

            try {
                //Создаем новую таблицу
                Workbook wb = new XSSFWorkbook();
                // записываем созданный в памяти Excel документ в файл
                FileOutputStream out = new FileOutputStream("D:\\java\\exel\\ntx1000_data\\" + ntx1000_name_data + ".xlsx");
                Sheet sheet0 = wb.createSheet("Zagruzka");
                wb.write(out);
                System.out.println(ntx1000_name_data + " Excel файл успешно создан!");
                wb.close();

                ntx1000_old_name_data = ntx1000_name_data;
            } catch (FileNotFoundException e) {
                System.out.println("Произошла ошибка при создании файла ntx1000");
            }
        }


        //Счетчик для определения % загрузки оборудования
        //по программе
        if ((status_work.equals("ACTIVE") && (Integer.parseInt(nagr_hpindel)<1))) { ntx1000_procent_work=ntx1000_procent_work+1; }
        //выключен
        if (status_work.equals("UNAVAILABLE")) { ntx1000_procent_off=ntx1000_procent_off+1; }
        //ожидание
        if ((status_work.equals("ACTIVE")==false) && (status_work.equals("UNAVAILABLE")==false))
        {ntx1000_procent_pause=ntx1000_procent_pause+1; }
        //авария
        if ((status_work.equals("UNAVAILABLE")==false) && (avar_sost.equals("ARMED")==false))
        {
            ntx1000_procent_avar=ntx1000_procent_avar+1;
            ntx1000_sost_error=1;
        } else
        {
            ntx1000_sost_error=0;
        }

        //под нагрузкой
        if ((status_work.equals("UNAVAILABLE")==false) && (Integer.parseInt(nagr_hpindel)>2))
        {
            ntx1000_procent_nagruzka=ntx1000_procent_nagruzka+1;
            ntx1000_sost_nagruzka=1; //Для фиксации тригерной логике
        } else
            {
                ntx1000_sost_nagruzka=0;
            }


        //Определяем хронологию работы оборудования фиксируя переключения состояний

        //Состояние "Работа"
        //Для статуса работы фиксируем состояние TON
        if (status_work.equals("ACTIVE") && status_work.equals(ntx1000_old_status_work) == false) {
            ntx1000_stroka_work=ntx1000_stroka_work+1;
            ntx1000_work_arrayList.add(ntx1000_stroka_work-1, ntx1000_formattedDate_for_hc);

            ntx1000_stroka_programname=ntx1000_stroka_programname+1;
            ntx1000_programname_arrayList.add(ntx1000_stroka_programname-1,name_program);
        }
        //Фиксируем статус для тригерной логики
        if (status_work.equals("ACTIVE")) {ntx1000_old_status_work="ACTIVE";}

        //Для статуса работы определяем состояние TOF
        if (ntx1000_old_status_work.equals("ACTIVE") && !status_work.equals("ACTIVE")) {
            ntx1000_old_status_work="null";
            ntx1000_stroka_work=ntx1000_stroka_work+1;
            ntx1000_work_arrayList.add(ntx1000_stroka_work-1, ntx1000_formattedDate_for_hc);
        }

        //Состояние "Выключен"
        if (status_work.equals("UNAVAILABLE") && status_work.equals(ntx1000_old_status_off) == false) {
            ntx1000_stroka_off=ntx1000_stroka_off+1;
            ntx1000_off_arrayList.add(ntx1000_stroka_off-1, ntx1000_formattedDate_for_hc);
        }
        //Фиксируем статус для тригерной логики
        if (status_work.equals("UNAVAILABLE")) {ntx1000_old_status_off="UNAVAILABLE";}

        //Для статуса работы определяем состояние TOF
        if (ntx1000_old_status_off.equals("UNAVAILABLE") && !status_work.equals("UNAVAILABLE")) {
            ntx1000_old_status_off="null";
            ntx1000_stroka_off=ntx1000_stroka_off+1;
            ntx1000_off_arrayList.add(ntx1000_stroka_off-1, ntx1000_formattedDate_for_hc);
        }

        //Авария
        //TON
        if (ntx1000_sost_error==1 && ntx1000_triger_error != 1) {
            ntx1000_triger_error=1;
            ntx1000_stroka_error=ntx1000_stroka_error+1;
            ntx1000_avar_arrayList.add(ntx1000_stroka_error-1, ntx1000_formattedDate_for_hc);
            System.out.println("Станок ntx1000 перешел в состояние авария");
        }
        //TOF
        if  (ntx1000_sost_error==0 && ntx1000_procent_avar>0 && ntx1000_triger_error != 0) {
            ntx1000_triger_error=0;
            ntx1000_stroka_error=ntx1000_stroka_error+1;
            ntx1000_avar_arrayList.add(ntx1000_stroka_error-1, ntx1000_formattedDate_for_hc);
        }

        //Ожидание
        //TON
        if ((status_work.equals("ACTIVE")==false) && (status_work.equals("UNAVAILABLE")==false) && ntx1000_sost_error==0 && ntx1000_triger_pause != 1 && ntx1000_triger_error == 0) {
            ntx1000_triger_pause=1;
            ntx1000_stroka_white=ntx1000_stroka_white+1;
            ntx1000_white_arrayList.add(ntx1000_stroka_white-1, ntx1000_formattedDate_for_hc);
        }
        //TOF
        if (((status_work.equals("ACTIVE")==true) || (status_work.equals("UNAVAILABLE")==true)) && ntx1000_sost_error==1 && ntx1000_procent_pause>0 && ntx1000_triger_pause != 0) {
            ntx1000_triger_pause=0;
            ntx1000_stroka_white=ntx1000_stroka_white+1;
            ntx1000_white_arrayList.add(ntx1000_stroka_white-1, ntx1000_formattedDate_for_hc);
        }


        //Под нагрузкой
        //TON
        if (ntx1000_sost_nagruzka==1 && ntx1000_triger_nagruzka !=1) {
            ntx1000_triger_nagruzka=1;
            ntx1000_stroka_nagruzka=ntx1000_stroka_nagruzka+1;
            ntx1000_nagruzka_arrayList.add(ntx1000_stroka_nagruzka-1, ntx1000_formattedDate_for_hc);
        }
        //TOF
        if  (ntx1000_sost_nagruzka==0 && ntx1000_procent_nagruzka>0 && ntx1000_triger_nagruzka != 0) {
            ntx1000_triger_nagruzka=0;
            ntx1000_stroka_nagruzka=ntx1000_stroka_nagruzka+1;
            ntx1000_nagruzka_arrayList.add(ntx1000_stroka_nagruzka-1, ntx1000_formattedDate_for_hc);
        }

        //Записываем текущие значение в таблицу
        Workbook wb = new XSSFWorkbook();
        FileOutputStream out = new FileOutputStream("D:\\java\\exel\\ntx1000_data\\" + ntx1000_name_data + ".xlsx");
        Sheet sheet0=wb.createSheet("Zagruzka");

        //Создаем строку 1
        Row row = sheet0.createRow(0);

        //Создаем и записываем в столбец 0
        Cell cell =row.createCell(0);
        cell.setCellValue(ntx1000_procent_work);

        //Создаем и записываем в столбец 1
        Cell cell1 =row.createCell(1);
        cell1.setCellValue(ntx1000_procent_pause);

        //Создаем и записываем в столбец 2
        Cell cell2 =row.createCell(2);
        cell2.setCellValue(ntx1000_procent_off);

        //Создаем и записываем в столбец 3
        Cell cell33 =row.createCell(3);
        cell33.setCellValue(ntx1000_procent_avar);

        //Создаем и записываем в столбец 4
        Cell cell44 =row.createCell(4);
        cell44.setCellValue(ntx1000_procent_nagruzka);


        //Создаем строку 2 для заполнения массивом "работа"
        row = sheet0.createRow(1);
        //определяем длину массива
        int length_array_work = ntx1000_work_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_work; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = ntx1000_work_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        //Создаем строку 3 для заполнения массивом "ожидание"
        row = sheet0.createRow(2);
        //определяем длину массива
        int length_array_white = ntx1000_white_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_white; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = ntx1000_white_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        //Создаем строку 4 для заполнения массивом "выключен"
        row = sheet0.createRow(3);
        //определяем длину массива
        int length_array_off = ntx1000_off_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_off; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = ntx1000_off_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        //Создаем строку 5 для заполнения массивом "авария"
        row = sheet0.createRow(4);
        //определяем длину массива
        int length_array_avar = ntx1000_avar_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_avar; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = ntx1000_avar_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        //Создаем строку 6 для заполнения массивом "под нагрузкой"
        row = sheet0.createRow(5);
        //определяем длину массива
        int length_array_nagruzka = ntx1000_nagruzka_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_nagruzka; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = ntx1000_nagruzka_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        //Создаем строку 7 для заполнения массивом "имя программы"
        row = sheet0.createRow(6);
        int length_array_nameprogram = ntx1000_programname_arrayList.size();
        //Заполняем ячейки значениями из массива
        for (int i=0; i<length_array_nameprogram; i++) {
            //Преобразовываем значение массива в стринг
            String array_string = ntx1000_programname_arrayList.get(i);
            row.createCell(i).setCellValue(array_string);
        }

        wb.write(out);
        wb.close();

        System.out.print(" ntx1000: " + status_work + " " + ntx1000_formattedDate + " " + ntx1000_procent_work + " " + ntx1000_procent_pause + " " + ntx1000_procent_off);

        TimeUnit.MILLISECONDS.sleep(5000/ Main.kol_device);

        //Записываем текущие значение в таблицу
        Workbook wb2 = new XSSFWorkbook();
        FileOutputStream out2 = new FileOutputStream("D:\\java\\exel\\ntx1000_realtime_data.xlsx");
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

        Cell cell15 =row3.createCell(12);
        cell15.setCellValue(nagr_hpindel);

        Cell cell16 =row3.createCell(13);
        cell16.setCellValue(avar_sost);

        wb2.write(out2);
        wb2.close();
    }
}
