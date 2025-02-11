package com.prglolinks.emailAutomation.Service;

import com.prglolinks.emailAutomation.Profile;
import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeMessage;
import jakarta.validation.Valid;
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

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Properties;

/* CHANGES TO BE MADE
1. Remove unwanted code comments
2. Add more Logger to track the follow
3. Remove all sop statement
4. Follow proper name conversion
5. Remove unwanted spaces
6. Align the code properly
7. ClassPathResource Path URL, move to application properties or make it as static final and reuse it.
*/


@Service

public class ExcelService {
    private static final Logger logger = LoggerFactory.getLogger(ExcelService.class);

    @Autowired
    private JavaMailSender javaMailSender;

    public void createUser(@Valid Profile profile) { // @Valid here!
        // Validation is automatically triggered before this method executes.
        // If validation fails, a ConstraintViolationException is thrown.

        // If validation succeeds, proceed with business logic:
        System.out.println("Creating user: " + profile);
        // ... your service logic to create the user
    }
    public byte[] getExcelFile(String fileName) throws IOException {
        ClassPathResource resource = new ClassPathResource("static/" + fileName);

        if (!resource.exists()) {
            return null; // Handle file not found case in controller
        }

        try (InputStream inputStream = resource.getInputStream()) {
            return inputStream.readAllBytes();
        }
    }



    public byte[] readExcelFile(@Valid MultipartFile file) throws IOException {

        List<Profile> profiles = new ArrayList<>();

        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = createWorkbook(inputStream, file.getOriginalFilename());

            Sheet sheet = workbook.getSheetAt(0);


            for (int j = 1; j <= sheet.getLastRowNum(); j++) {

                Row row = sheet.getRow(j);
                int no = (int) (row.getCell(0).getNumericCellValue());
                String name = getCellValueAsString(row.getCell(1));
                String email = getCellValueAsString(row.getCell(2));
                String dateOfBirth = getCellValueAsString(row.getCell(3));
                String subject = getCellValueAsString(row.getCell(4));
                String content = getCellValueAsString(row.getCell(5));

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
                Profile temp = new Profile(no, name, email, dateOfBirth, subject, content, String.valueOf(ssb));

// Check for any invalid data, if found email will not get triggered
                if (!String.valueOf(temp.getResult()).contains("invalid")) {
                    String to = temp.getEmail();

                    String from = "8056182081pa@gmail.com";     // Sender's email ID


                    String host = "smtp.gmail.com";             // sending through gmail


                    Properties properties = System.getProperties();      // Get system properties


                    Session session = Session.getInstance(properties, new Authenticator() {});  // Get the Session object.
//
                    try {

                        MimeMessage message = new MimeMessage(session);          // Create a default MimeMessage object.


                        message.setFrom(new InternetAddress(from));          // Set From: header field of the header.


                        message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));        // Set To: header field of the header.


                        message.setSubject(temp.getSubject());       // Set Subject: header field


                        message.setText(temp.getContent());             // Now set the actual message // or message.setContent(...) for HTML content


                        javaMailSender.send(message);                    // Send message
                        temp.setResult("Mail Sent");
                        profiles.add(temp);

                    } catch (MessagingException mex) {
                        logger.error("Error sending the message to : {}", temp, mex);
                    }
                } else {
                    profiles.add(temp);
                    logger.error("Invalid Data found in the current profile : {}", temp);

            }
                }


                    Field[] fields = Profile.class.getDeclaredFields();             // Get the fields of the class for the header
                    List<String> headers = new ArrayList<>();
                    for (Field field : fields) {
                        headers.add(field.getName());
                    }


                    Workbook dworkbook = new XSSFWorkbook();                    // Create Excel workbook and sheet
                    Sheet dsheet = dworkbook.createSheet("Sheet1");


                    Row headerRow = dsheet.createRow(0);                     // Create header row
                    for (int i = 0; i < headers.size(); i++) {
                        Cell cell = headerRow.createCell(i);
                        cell.setCellValue(headers.get(i));
                    }

                    int rowNum = 1;                                         // Create data rows using reflection
                    for (Object obj : profiles) {                           // Use the generic Object type
                        Row drow = sheet.createRow(rowNum++);
                        for (int i = 0; i < fields.length; i++) {
                            Cell cell = null;
                            try {
                                Field field = fields[i];
                                field.setAccessible(true);                      // Make private fields accessible
                                Object value = field.get(obj);                  // Get the field value
                                cell = drow.createCell(i);
                                if (value != null) {                                // Handle null values appropriately
                                    cell.setCellValue(value.toString());            // Convert to string
                                }
                            } catch (IllegalAccessException e) {
                                logger.error("Error finding the field, does not have access", e);
                                cell.setCellValue("Error");                         // Or a more appropriate error value
                            }
                        }
                    }

                    for (int i = 0; i < headers.size(); i++) {              // Auto-size columns (optional)
                        sheet.autoSizeColumn(i);
                    }
                    String tmpDir = "C:\\Spring Boot Dev\\prglolinks\\emailAutomationWithSpring\\src\\main\\resources\\static";
                    File excelFile = new File(tmpDir, "dworkbook.xlsx");                        // Create File object
                    try (FileOutputStream outputStream = new FileOutputStream(excelFile)) {
                        workbook.write(outputStream);
                    } catch (IOException e) {
                        logger.error("Error while writing the file", excelFile, e);
                    }
                }

        return getExcelFile("dworkbook.xlsx");
    }

    public StringBuffer validateData(String header,boolean isEmail, boolean isdate, String key) {
        StringBuffer sb = new StringBuffer();

        if (isEmail) {
            if (StringUtils.isBlank(header) || !header.contains("@")) {
                sb.append(key + " id is invalid |");
            } else {
                sb.append(key + " id is valid |");
            }
        }
            else if (isdate) {
                if (StringUtils.isBlank(header) || !header.contains("/")) {
                    sb.append(key + " is invalid |");
                } else {
                    sb.append(key + " id is valid |");
                    }
                } else {
                if(StringUtils.isBlank(header)){
                    sb.append(key + " is invalid |");
            } else {sb.append(key + " is valid |");}
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
                if (DateUtil.isCellDateFormatted(cell)) {               // Check if it's a date
                    Date date = cell.getDateCellValue();                   // Get the Date object
                    LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();    // Convert to LocalDate
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd");                // Your desired format
                    return localDate.format(formatter);
                } else {
                    return String.valueOf(cell.getNumericCellValue()); // It's numeric, not a date
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return String.valueOf(cell.getNumericCellValue());// Formulas usually return a numeric value

            default:
                return "";
        }
    }


    public byte[] getresponseExcelFile(String fileName) throws IOException {
        ClassPathResource resource = new ClassPathResource("C:\\Spring Boot Dev\\prglolinks\\emailAutomationWithSpring\\src\\main\\resources\\static" + fileName);

        if (!resource.exists()) {
            logger.warn("File not found : {} ", resource);
            return null; // Handle file not found case in controller
        }
        try (InputStream inputStream = resource.getInputStream()) {
            return inputStream.readAllBytes();
        }
    }


}
