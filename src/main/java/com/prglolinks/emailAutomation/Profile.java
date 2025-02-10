package com.prglolinks.emailAutomation;

public class Profile {

    private int sNo;
    private String fullName;
    private String email;
    private String dateOfBirth;
    private String subject;
    private String content;
    private String result;


    public Profile(){};

    public Profile(int no,String fullName, String email, String dateOfBirth, String subject, String content, String result) {
        this.sNo = no;
        this.fullName = fullName;
        this.email = email;
        this.dateOfBirth = dateOfBirth;
        this.subject = subject;
        this.content = content;
        this.result = result;

    }

    public int getsNo() {
        return sNo;
    }

    public void setsNo(int sNo) {
        this.sNo = sNo;
    }

    public String getFullName() {
        return fullName;
    }

    public void setFullName(String fullName) {
        this.fullName = fullName;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getDateOfBirth() {
        return dateOfBirth;
    }

    public void setDateOfBirth(String dateOfBirth) {
        this.dateOfBirth = dateOfBirth;
    }

    public String getSubject() {
        return subject;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public String getResult() {
        return result;
    }

    public void setResult(String result) {
        this.result = result;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    @Override
    public String toString() {
        return "Profile{" +
                "sNo=" + sNo +
                ", fullName='" + fullName + '\'' +
                ", email='" + email + '\'' +
                ", dateOfBirth='" + dateOfBirth + '\'' +
                ", subject='" + subject + '\'' +
                ", result='" + result + '\'' +
                ", content='" + content + '\'' +
                '}';
    }
}
