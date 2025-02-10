package com.prglolinks.emailAutomation.Controller;

import com.prglolinks.emailAutomation.Service.ExcelService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController

// change the controller class name something genric to project
//BulkEmailController
//EmailSenderController
public class ExcelDownloadController {


    private final ExcelService excelService;
    private static final Logger logger = LoggerFactory.getLogger(ExcelDownloadController.class);


    public ExcelDownloadController(ExcelService excelService) {
            this.excelService = excelService;
        }

        @GetMapping("/downloadExcel")
        public ResponseEntity<byte[]> downloadExcel() {

            String fileName = "prglolinks.xlsx";

            try {
                byte[] excelBytes = excelService.getExcelFile(fileName);

                if (excelBytes == null) {
                    logger.warn("File not found: {}", fileName);
                    return ResponseEntity.notFound().build();
                }

                HttpHeaders headers = new HttpHeaders();
                headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                headers.setContentDispositionFormData("attachment", fileName);

                return new ResponseEntity<>(excelBytes, headers, HttpStatus.OK);
            } catch (IOException e) {
                logger.error("Error while reading the file: {}", fileName, e);
                return ResponseEntity.internalServerError().build();
            }
        }

    @PostMapping("/uploadExcel")
    public ResponseEntity<byte []> uploadExcel(@RequestParam("file") MultipartFile file) {
        try {
            byte[] excelBytes  = excelService.readExcelFile(file);

            if (excelBytes == null) {
                logger.warn("File not found: {}", file);
                return ResponseEntity.notFound().build();
            }

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", "file");

            return new ResponseEntity<>(excelBytes, headers, HttpStatus.OK);
        } catch (IOException e) {
            logger.error("Error while reading the file: {}", file, e);
            return ResponseEntity.internalServerError().build();
        }
    }

}
