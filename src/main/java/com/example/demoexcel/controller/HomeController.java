package com.example.demoexcel.controller;

import com.example.demoexcel.dto.Employee;
import lombok.extern.slf4j.Slf4j;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Slf4j
@RestController
public class HomeController {

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public ResponseEntity<Object> downloadFile() throws IOException {
        log.info("Running Object Collection demo");
        Employee employee = new Employee();
        employee.setName("Hoang");
        employee.setBirthDate(new Date());
        employee.setBonus(BigDecimal.valueOf(3200000));
        employee.setPayment(BigDecimal.valueOf(42400000));
        List<Employee> employees = new ArrayList<>();
        employees.add(employee);
        employees.add(employee);
        employees.add(employee);
        String filenames = "object_collection_output.xls";
        File files = new File(filenames);
        try(InputStream is = new FileInputStream(files)) {
            try (OutputStream os = new FileOutputStream("target/object_collection_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A2");
            }
        }

        String filename = "target/object_collection_output.xls";
        File file = new File(filename);
        InputStreamResource resource = new InputStreamResource(new FileInputStream(file));


        HttpHeaders headers = new HttpHeaders();

        headers.add("Content-Disposition", String.format("attachment; filename=\"%s\"", file.getName()));
        headers.add("Cache-Control", "no-cache, no-store, must-revalidate");
        headers.add("Pragma", "no-cache");
        headers.add("Expires", "0");

        ResponseEntity<Object>
                responseEntity = ResponseEntity.ok().headers(headers).contentLength(
                file.length()).contentType(MediaType.parseMediaType("application/txt")).body(resource);

        return responseEntity;
    }


}
