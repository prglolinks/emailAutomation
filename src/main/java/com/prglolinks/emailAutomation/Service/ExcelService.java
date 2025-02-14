package com.prglolinks.emailAutomation.Service;

import jakarta.mail.MessagingException;
import jakarta.mail.Session;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeMessage;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.regex.Pattern;

/* CHANGES TO BE MADE
1. Remove unwanted code comments
2. Add more Logger to track the follow
3. Remove all sop statement
4. Follow proper name conversion
5. Remove unwanted spaces
6. Align the code properly
7. ClassPathResource Path URL, move to application properties or make it as static final and reuse it.
8. ClassPath URL is pointing to your local path, Code will not work on other's machine
9. SMTP FROM, HOST AND PASSWORD VALUES MUST BE USED FROM APPLICATION PROPERTIES, AND NOT FROM THE CODE.
*/


@Service
public class ExcelService {

    public static final Logger LOGGER = LoggerFactory.getLogger(ExcelService.class);

    @Autowired
    private JavaMailSender javaMailSender;

    public static String paths = "src/main/files/";

    @Value("${spring.mail.username}")
    private String username;

    public byte[] getExcelFile(String fileName) throws IOException {

        LOGGER.info("Inside getExcelFile method from service");
        byte [] data = null;
        ClassPathResource resource = new ClassPathResource("static/" + fileName);
        if (!resource.exists()) {
            LOGGER.warn("File Not Found in the resources/static directory");
            throw new FileNotFoundException("Template File not found in the specified directory");
        }
        else {
            try (InputStream inputStream = resource.getInputStream()) {
                LOGGER.info("Template Excel data read successfully");
                data = new byte[fileName.length()];
                data = inputStream.readAllBytes();
            } catch (Exception exception){
                LOGGER.error("Exception occured : " + exception);
//                throw new Exception()
            }
        }
        return data;
    }

    public byte[] readExcelFile(MultipartFile file) throws Exception {

        LOGGER.info("Getting into readExcelFile Method");
        try {
            File uploadedFile = new File(paths, file.getOriginalFilename());
            uploadedFile.createNewFile();
            FileOutputStream fos = new FileOutputStream(uploadedFile);
            fos.write(file.getBytes());
            LOGGER.info("Input file has been stored in src/main/files directory");

            LOGGER.info("Reading the stored file");
            FileInputStream inputStream = new FileInputStream(uploadedFile);

            Workbook workbook = createWorkbook(inputStream, file.getOriginalFilename());
            Sheet sheet = workbook.getSheetAt(0);
            LOGGER.info("Workbook created and got the sheet ");

            LOGGER.info("Loop started from 1st row excluding header row");
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);

                LOGGER.info("reading row details by each column");
                int no = (int) (row.getCell(0).getNumericCellValue());
                String name = getCellValueAsString(row.getCell(1));
                String email = getCellValueAsString(row.getCell(2));
                String dateOfBirth = getCellValueAsString(row.getCell(3));
                String subject = getCellValueAsString(row.getCell(4));
                String content = getCellValueAsString(row.getCell(5));

                LOGGER.info("Data validation started");
                StringBuffer ssb = validateData(name,email,dateOfBirth,subject,content);

                Cell cell = row.getCell(6);
                if (cell == null) {
                    cell = row.createCell(6, CellType.STRING);
                }
                cell.setCellValue(ssb.toString());
                LOGGER.info("Data has been written in result column i.e cell(6)");

                LOGGER.info("Email check started");
                if (ssb.isEmpty()) {

                    LOGGER.info("Getting system properties and password");
                    Properties properties = System.getProperties();

                    Session session = Session.getInstance(properties);

                    LOGGER.info("Trigerring email with mime message");

                        MimeMessage message = new MimeMessage(session);
                        MimeMessageHelper helper = new MimeMessageHelper(message, true);
                        helper.setFrom(new InternetAddress(username));
                        helper.setTo(email);
                        helper.setSubject(subject);
                        helper.setText(content);

                        Path path = Paths.get(paths,"/attach.pdf");
                        byte[] attachContent = Files.readAllBytes(path);
                        helper.addAttachment("attach.pdf",new ByteArrayResource(attachContent));
                        javaMailSender.send(message);

                        row.getCell(6).setCellValue("Mail Sent");
                        LOGGER.info("Mail sent and result appended");
                }
                else {
                    LOGGER.info("Invalid Data found in the current profile with no : {}", no);
                }
            }

            LOGGER.info("writing result data to the file stored");
            FileOutputStream foss = new FileOutputStream(uploadedFile);
            workbook.write(foss);
            LOGGER.info("Getting the path of stored data and reading as bytes");
            String path = paths + "/FinalSheet.xlsx";
            FileInputStream fis = new FileInputStream(path);
            return fis.readAllBytes();
        } catch (FileNotFoundException | SecurityException exception){
            LOGGER.error("Exception occurred " + exception.getMessage());
            return ("Exception occured " + exception.getMessage()).getBytes();
        } catch (IllegalArgumentException illegalArgumentException){
            LOGGER.info("Invalid file format. Only .xls and .xlsx are supported.");
            return "Invalid file format. Only .xls and .xlsx are supported.".getBytes();
        } catch (MessagingException messagingException) {
            LOGGER.error("Error sending the message to the current row", messagingException);
            return ("Error sending the message to the current row" + messagingException).getBytes();
        }

    }


    public StringBuffer validateData(String name, String email, String dob, String subject, String content) {
        LOGGER.info("Inside validateData Method");
        StringBuffer sb = new StringBuffer();
        if (StringUtils.isBlank(email) || !email.contains("@") || !email.contains(".")) {
            sb.append(" Email ID is invalid |");
        }
        if (StringUtils.isBlank(dob) || !dob.contains("/")) {
            sb.append(" Date of Birth is invalid |");
        }
        if (StringUtils.isBlank(name) || !Pattern.matches("[a-zA-Z]+",name)){
            sb.append(" Name is empty/only letters allowed |");
        }
        if (StringUtils.isBlank(subject)){
            sb.append(" Subject is empty |");
        }
        if (StringUtils.isBlank(content)){
            sb.append(" Content is empty |");
        }
        LOGGER.info("Validation done for all fields");
        return sb;
    }

    private Workbook createWorkbook(InputStream inputStream, String fileName) throws IOException {
        LOGGER.info("Inside createWorkbook method");
        if (fileName.endsWith(".xlsx")) {
            return new XSSFWorkbook(inputStream);
        } else if (fileName.endsWith(".xls")) {
            return new HSSFWorkbook(inputStream);
        } else {
            LOGGER.error("Invalid file format. Only .xls and .xlsx are supported.");
            throw new IllegalArgumentException("Invalid file format. Only .xls and .xlsx are supported.");
        }
    }

    private String getCellValueAsString(Cell cell) {
        LOGGER.info("Insode getCellValueAsString method");
        String result;
        if (cell == null || cell.getCellType() == CellType.BLANK){
            result = "";
        }
        else {
            switch (cell.getCellType()) {
                case STRING -> result = cell.getStringCellValue();
                case NUMERIC -> {
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        Date date = cell.getDateCellValue();
                                        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
                                        String strDate = formatter.format(date);
                                        result = strDate;
                                    }
                                    else {
                                        result = String.valueOf(cell.getNumericCellValue());
                                    }
                                }
                case BOOLEAN -> result = String.valueOf(cell.getBooleanCellValue());
                case FORMULA -> result = String.valueOf(cell.getNumericCellValue());
                default -> result = "";
            }
        }
        LOGGER.info("Data read successfully and stored in variables");
        return result;
    }
}
