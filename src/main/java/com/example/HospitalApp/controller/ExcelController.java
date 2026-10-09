package com.example.HospitalApp.controller;

import com.example.HospitalApp.service.ExcelService2;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

@Controller
public class ExcelController {

    private final ExcelService2 excelService;

    public ExcelController(ExcelService2 excelService) {
        this.excelService = excelService;
    }

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @PostMapping("/upload")
    public ResponseEntity<StreamingResponseBody> uploadFiles(@RequestParam("files")List<MultipartFile> files) throws IOException {

        ByteArrayInputStream inputStream = excelService.processExcelFiles(files);

        StreamingResponseBody responseBody = outputStream -> {
          byte[] buffer = new byte[1024];
          int bytesRead;
          while((bytesRead = inputStream.read(buffer)) != -1) {
              outputStream.write(buffer,0,bytesRead);
          }
        };

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=result.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(responseBody);
    }
}
