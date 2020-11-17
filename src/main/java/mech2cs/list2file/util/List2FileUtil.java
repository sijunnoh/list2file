package mech2cs.list2file.util;

import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class List2FileUtil {

    public static byte[] list2CSV(List<?> list) throws Exception {
        return list2CSV(list, ',', "yyyy-MM-dd hh:mm:ss");
    }

    public static byte[] list2CSV(List<?> list, char delimiter) throws Exception {
        return list2CSV(list, delimiter, "yyyy-MM-dd hh:mm:ss");
    }

    public static byte[] list2CSV(List<?> list, String datePattern) throws Exception {
        return list2CSV(list, ',', datePattern);
    }

    public static byte[] list2CSV(List<?> list, char delimiterChar, String datePattern) throws Exception {
        String delimiter = String.valueOf(delimiterChar);
        if(list.size() == 0){
            return new StringBuffer().toString().getBytes();
        }

        SimpleDateFormat format = new SimpleDateFormat(datePattern);

        StringBuffer data = new StringBuffer();

        Field[] fields = list.get(0).getClass().getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            fields[i].setAccessible(true);
            data.append(fields[i].getName());
            if(i!=fields.length-1) {
                data.append(delimiter);
            }
        }
        data.append('\n');
        for (int i = 0; i < list.size(); i++) {
            for (int j = 0; j < fields.length; j++) {

                if (fields[j].get(list.get(i)) instanceof Date || fields[j].get(list.get(i)) instanceof java.util.Date) {
                    data.append(StringEscapeUtils.escapeCsv(format.format((java.util.Date) fields[j].get(list.get(i)))));
                } else if (fields[j].get(list.get(i)) instanceof LocalDate) {
                    data.append(StringEscapeUtils.escapeCsv(((LocalDate) fields[j].get(list.get(i))).format(DateTimeFormatter.ofPattern(datePattern))));
                } else if (fields[j].get(list.get(i)) instanceof LocalDateTime) {
                    data.append(StringEscapeUtils.escapeCsv(((LocalDateTime) fields[j].get(list.get(i))).format(DateTimeFormatter.ofPattern(datePattern))));
                } else if (fields[j].get(list.get(i)) == null) {
                } else if (fields[j].get(list.get(i)) instanceof Boolean) {
                    StringEscapeUtils.escapeCsv(fields[j].get(list.get(i)).toString());
                } else {
                    String tempStr;
                    if(fields[j].get(list.get(i)).toString().contains(String.valueOf(delimiter))) {
                        tempStr = StringEscapeUtils.escapeCsv("," + fields[j].get(list.get(i)).toString());
                        tempStr = tempStr.replaceFirst(",","");
                    }else {
                        tempStr = StringEscapeUtils.escapeCsv(fields[j].get(list.get(i)).toString());
                    }
                    data.append(tempStr);
                }
                if(j!=fields.length-1) {
                    data.append(delimiter);
                }
            }
            data.append('\n');
        }

        return data.toString().getBytes();

    }

    public static byte[] list2WorkBook(List<?> list) throws IOException, IllegalAccessException{
        return list2WorkBook(list, "yyyy-MM-dd hh:mm:ss");
    }

    public static byte[] list2WorkBook(List<?> list, String datePattern) throws IOException, IllegalAccessException {

        Workbook workbook = WorkbookFactory.create(true);
        Sheet sheet = workbook.createSheet();

        if(list.size() == 0){
            return workbook.toString().getBytes();
        }

        DataFormat format = workbook.createDataFormat();

        // date style
        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(format.getFormat(datePattern));

        Row row;
        Cell cell;

        row = sheet.createRow(0);

        Field[] fields = list.get(0).getClass().getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            fields[i].setAccessible(true);
            cell = row.createCell(i);
            cell.setCellValue(fields[i].getName());
        }

        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.length; j++) {

                cell = row.createCell(j);

                if (fields[j].get(list.get(i)) instanceof Integer) {
                    cell.setCellValue((Integer) fields[j].get(list.get(i)));
                } else if (fields[j].get(list.get(i)) instanceof Double) {
                    cell.setCellValue((Double) fields[j].get(list.get(i)));
                } else if (fields[j].get(list.get(i)) instanceof Float) {
                    cell.setCellValue((Float) fields[j].get(list.get(i)));
                } else if (fields[j].get(list.get(i)) instanceof Date ||
                        fields[j].get(list.get(i)) instanceof java.util.Date) {

                    cell.setCellValue((java.util.Date) fields[j].get(list.get(i)));
                    cell.setCellStyle(dateStyle);

                } else if (fields[j].get(list.get(i)) instanceof LocalDate) {

                    cell.setCellValue((LocalDate) fields[j].get(list.get(i)));
                    cell.setCellStyle(dateStyle);

                } else if (fields[j].get(list.get(i)) instanceof LocalDateTime) {

                    cell.setCellValue((LocalDateTime) fields[j].get(list.get(i)));
                    cell.setCellStyle(dateStyle);

                } else if (fields[j].get(list.get(i)) instanceof String) {
                    cell.setCellValue(fields[j].get(list.get(i)).toString());
                } else if (fields[j].get(list.get(i)) instanceof Boolean) {
                    cell.setCellValue((Boolean)fields[j].get(list.get(i)));
                } else {
                    //                    cell.setCellValue("invalid data");
                }
            }
        }

        try(ByteArrayOutputStream bos = new ByteArrayOutputStream()){
            workbook.write(bos);
            byte[] bytes = bos.toByteArray();
            return bytes;
        }
    }
}
