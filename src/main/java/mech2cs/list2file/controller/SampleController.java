package mech2cs.list2file.controller;

import mech2cs.list2file.util.List2FileUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.util.StreamUtils;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

@Controller
public class SampleController {

    @GetMapping(value = "/csv")
    public void getCsvList(HttpServletResponse response) {

        List<SampleDTO> list = getSampleList();

        try {
            byte[] byteArray = List2FileUtil.list2CSV(list);
            response.setContentType("text/plain; charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=sampleCSV.csv");
            StreamUtils.copy(byteArray,response.getOutputStream());

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    @GetMapping(value = "/excel")
    public void getExcelList(HttpServletResponse response) {

        List<SampleDTO> list = getSampleList();

        try {
            byte[] byteArray = List2FileUtil.list2WorkBook(list);
            response.setContentType("ms-vnd/excel");
            response.setHeader("Content-Disposition", "attachment;filename=sampleExcel.xlsx");
            StreamUtils.copy(byteArray,response.getOutputStream());

        } catch (IllegalAccessException | IOException e) {
            e.printStackTrace();
        }
    }

    public List<SampleDTO> getSampleList(){
        List<SampleDTO> list = new ArrayList<>();
        SampleDTO person1 = new SampleDTO();
        person1.setAge(15);
        person1.setBirthday(LocalDateTime.of(2010,01,01,00,00));
        person1.setName("김씨");
        list.add(person1);

        SampleDTO person2 = new SampleDTO();
        person2.setAge(16);
        person2.setBirthday(LocalDateTime.of(2020,01,03,00,00));
        person2.setName("최씨");
        list.add(person2);

        SampleDTO person3 = new SampleDTO();
        person3.setAge(16);
        person3.setBirthday(LocalDateTime.of(2000,01,05,00,00));
        person3.setName("이씨");
        list.add(person2);


        return list;
    }

}
