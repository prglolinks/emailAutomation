package com.prglolinks.emailAutomation;

import java.io.File;
import java.io.IOException;

public class Practice {
    public static void main(String[] args) throws IOException {

//
//        // Recipient's email ID
//        String to = "takrawelango@gmail.com";
//
//        // Sender's email ID
//        String from = "8056182081pa@gmail.com";  // Must be a valid email
//
//        // Assuming you are sending through gmail
//        String host = "smtp.gmail.com";
//
//        // Get system properties
//        Properties properties = System.getProperties();
//
//        // Setup mail server properties
//        properties.put("mail.smtp.host", host);
//        properties.put("mail.smtp.port", "587"); // or 465
//        properties.put("mail.smtp.starttls.enable", "true"); // For TLS
//        properties.put("mail.smtp.auth", "true");
//
//        // Get the Session object.
//        Session session = Session.getInstance(properties, new Authenticator() {
//            protected PasswordAuthentication getPasswordAuthentication() {
//                return new PasswordAuthentication("8056182081pa@gmail.com", "baoz tgsq obra fzzx"); // Your Gmail password or App Password
//            }
//        });
//
//        try {
//            // Create a default MimeMessage object.
//            MimeMessage message = new MimeMessage(session);
//
//            // Set From: header field of the header.
//            message.setFrom(new InternetAddress(from));
//
//            // Set To: header field of the header.
//            message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
//
//            // Set Subject: header field
//            message.setSubject("Test Email For Java Application");
//
//            // Now set the actual message
//            message.setText("Hey Elango," +
//                    "Hope you received the email" +
//                    "Regards, Pani."); // or message.setContent(...) for HTML content
//
//            // Send message
//            Transport.send(message);
//            System.out.println("Sent message successfully....");
//        } catch (MessagingException mex) {
//            mex.printStackTrace();
//        }
//    }
        File file = new File("test.txt");

        System.out.println(System.getProperty("user.dir"));


    }
}





//        String path = "C:\\Users\\Pani Vignesh\\Downloads\\FinalSheet.xlsx";
//
//        try(FileInputStream file = new FileInputStream(new File(path))) {
//            Workbook workbook = null;
//
//            if(path.endsWith(".xlsx")){
//                workbook = new XSSFWorkbook(file);
//            }else if (path.endsWith(".xls")){
//                workbook = new HSSFWorkbook(file);
//            } else {
//                System.out.println("Invalid file format must be .xls or .xlsx");
//                return;
//            }
//
//            Sheet sheet = workbook.getSheetAt(0);
//            Row row = sheet.getRow(1);
//            System.out.println(row.getCell(0));
//
//            Iterator<Row> rowIterator = sheet.iterator();
////
////            while (rowIterator.hasNext()){
////                Row row = rowIterator.next();
////
////                Iterator<Cell> cellIterator = row.cellIterator();
////                while (cellIterator.hasNext()){
////                    Cell cell = cellIterator.next();
////                    if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
////                        // Cell is a date
////
////                        // Method 1: Get the Date object directly (Recommended)
////                        Date date = cell.getDateCellValue();
////                        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd"); // Customize date format
////                        String formattedDate = dateFormat.format(date);
////                        System.out.print(formattedDate + "\t");
////                        System.out.println();
////                    }
////                    else{
////                            switch (cell.getCellType()) {
////                                case STRING:
////                                    System.out.println(cell.getStringCellValue() + "\t");
////                                    break;
////                                case NUMERIC:
////                                    System.out.println(cell.getNumericCellValue() + "\t");
////                                    break;
////                                case BOOLEAN:
////                                    System.out.println(cell.getBooleanCellValue() + "\t");
////                                    break;
////                                case FORMULA:
////                                    System.out.println(cell.getCellFormula() + "\t");
////                                    break;
////                                case BLANK:
////                                    System.out.println("\t");
////                                    break;
////                                default:
////                                    System.out.println(cell.toString() + "\t");
////                            }
////                        }
////                }
////                System.out.println();
////            }
//
//
//            workbook.close();
//
//
//        } catch (FileNotFoundException e) {
//            throw new RuntimeException(e);
//        } catch (IOException e) {
//            throw new RuntimeException(e);
//        } ;


//
//    }
//}
