package org.gymbrain.gymbrainexcel.processor;


import org.gymbrain.gymbrainexcel.annotations.ExcelDocument;
import org.gymbrain.gymbrainexcel.annotations.ExcelField;

@ExcelDocument(name = "document")
public class TestObject {
    private String firstName;
    private String lastName;
    private int age;

    @ExcelField(name = "نام", index = 0)
    public String getFirstName() {
        return firstName;
    }

    public void setFirstName(String firstName) {
        this.firstName = firstName;
    }

    @ExcelField(name = "سن", index = 1)
    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    @ExcelField(name = "نام خانوادگی")
    public String getLastName() {
        return lastName;
    }

    public void setLastName(String lastName) {
        this.lastName = lastName;
    }
}
