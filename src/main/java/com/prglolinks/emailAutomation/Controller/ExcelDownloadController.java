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


/* CHANGES TO BE MADE
1. Remove unwanted code comments
2. Why ExcelService is in private final ? Why not @Service and @Autowired implementation??
3. Under Download Excel API -> Why are you throwing Internal Server Error?? 
4. Under Download Excel API -> What is e??, Follow Proper name conversion
5. Under UploadExcel API -> WHy there is a space between byte and [] symbol??
6. Under UploadExcel API -> Follow proper name conversation, LINE NUMBER: 71
7. Under UploadExcel API -> Remove extra spaces, Check line number 72
8. Under UploadExcel API -> What is e??, Follow Proper name conversion
9. Add more Logger to track the follow
*/

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
