package com.prglolinks.emailAutomation.Service;

import com.prglolinks.emailAutomation.Configuration.EmailConfig;
import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeMessage;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Properties;

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

    private Logger logger = LoggerFactory.getLogger(ExcelService.class);

    @Autowired
    private EmailConfig emailConfig;

    @Autowired
    private JavaMailSender javaMailSender;

    public byte[] getExcelFile(String fileName) throws IOException {
        logger.info("Inside getExcelFile method from service");

            ClassPathResource resource = new ClassPathResource("static/" + fileName);
            if (!resource.exists()) {
                logger.warn("File Not Found");
            }
            try (InputStream inputStream = resource.getInputStream()) {
                logger.info("Excel data read successfully");
                return inputStream.readAllBytes();
            }
        }

    public byte[] readExcelFile(MultipartFile file) throws IOException {

        logger.info("Getting into readExcelFile Method");
        File unzippedFile = new File("src/main/files/", file.getOriginalFilename());
        unzippedFile.createNewFile();

        FileOutputStream fos = new FileOutputStream(unzippedFile);
        fos.write(file.getBytes());
        logger.info("Input file has been stored in src/main/files directory");

        logger.info("Reading the stored file as inputstream");
        try (FileInputStream inputStream = new FileInputStream(unzippedFile) {
        }) {
            Workbook workbook = createWorkbook(inputStream, file.getOriginalFilename());
            Sheet sheet = workbook.getSheetAt(0);
        logger.info("Workbook created and got the sheet ");

        logger.info("Loop started from 1st row excluding header row");
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);

        logger.info("reading 1st row details by each column");
                int no = (int) (row.getCell(0).getNumericCellValue());
                String name = getCellValueAsString(row.getCell(1));
                String email = getCellValueAsString(row.getCell(2));
                String dateOfBirth = getCellValueAsString(row.getCell(3));
                String subject = getCellValueAsString(row.getCell(4));
                String content = getCellValueAsString(row.getCell(5));

        logger.info("Data validation started");
                String[] str = new String[row.getLastCellNum() - 1];
                str[0] = String.valueOf(validateData(name, false, false, "name"));
                str[1] = String.valueOf(validateData(email, true, false, "email"));
                str[2] = String.valueOf(validateData(dateOfBirth, false, true, "dateOfBirth"));
                str[3] = String.valueOf(validateData(subject, false, false, "subject"));
                str[4] = String.valueOf(validateData(content, false, false, "content"));

                StringBuffer ssb = new StringBuffer();
                for (String st : str) {
                    ssb.append(st);
                }
                Cell cell = row.getCell(6);
                if (cell == null) {
                    cell = row.createCell(6, CellType.STRING);
                }
                cell.setCellValue(String.valueOf(ssb));
        logger.info("Data has been written in result column i.e cell(6)");

        logger.info("Email check started");
                if (!String.valueOf(row.getCell(6)).contains("invalid")) {
                    String to = email;
                    String from = emailConfig.getUsername();

        logger.info("Getting system properties and password");
                    Properties properties = System.getProperties();
                    Session session = Session.getInstance(properties, new Authenticator() {
                        protected PasswordAuthentication getPasswordAuthentication() {
                            return new PasswordAuthentication(emailConfig.getUsername(), emailConfig.getUsername());
                        }
                    });

        logger.info("Trigerring email with mime message");
                    try {
                        MimeMessage message = new MimeMessage(session);
                        message.setFrom(new InternetAddress(from));
                        message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
                        message.setSubject(subject);
                        message.setText(content);
                        javaMailSender.send(message);
                        row.getCell(6).setCellValue("Mail Sent");
        logger.info("Mail sent and result appended");
                    } catch (MessagingException messagingException) {
                        logger.error("Error sending the message to : {}", name, messagingException);
                    }
                } else {
                        logger.info("Invalid Data found in the current profile : {}", name);
                }
            }

        logger.info("writing result data to the file stored");
            FileOutputStream foss = new FileOutputStream(unzippedFile);
            workbook.write(foss);
        }
        logger.info("Getting the path of stored data and reading as bytes");
        String path = "src/main/files/FinalSheet.xlsx";
        FileInputStream fis = new FileInputStream(path);
        return fis.readAllBytes();
    }

    public StringBuffer validateData(String header,boolean isEmail, boolean isdate, String key) {
        StringBuffer sb = new StringBuffer();
        logger.info("Validation for the row data");
        if (isEmail) {
            if (StringUtils.isBlank(header) || !header.contains("@")) {
                sb.append(key + " id is invalid |");
            }
            else {
                sb.append(key + " id is valid |");
            }
        }
            else if (isdate) {
                if (StringUtils.isBlank(header) || !header.contains("/")) {
                    sb.append(key + " is invalid |");
                }
                else {
                    sb.append(key + " id is valid |");
                }
            }
            else {
                if(StringUtils.isBlank(header)){
                    sb.append(key + " is invalid |");
                }
                else {sb.append(key + " is valid |");}
            }
        return sb;
    }

    private Workbook createWorkbook(InputStream inputStream, String fileName) throws IOException {
            if (fileName.endsWith(".xlsx")) {
                return new XSSFWorkbook(inputStream);
            } else if (fileName.endsWith(".xls")) {
                return new HSSFWorkbook(inputStream);
            } else {
                logger.error("Invalid file format. Only .xls and .xlsx are supported.");
                throw new IllegalArgumentException("Invalid file format. Only .xls and .xlsx are supported.");
            }
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd");
                    return localDate.format(formatter);
                }
                else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return String.valueOf(cell.getNumericCellValue());
            default:
                return "";
        }
    }

}
