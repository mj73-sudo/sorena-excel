package org.gymbrain.gymbrainexcel.processor;

import org.gymbrain.gymbrainexcel.file.ExcelWriter;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.beans.IntrospectionException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

class ObjectToExcelTest {
    ObjectToExcel objectToExcel = null;
    TestObject object = null;
    TestObject object1 = null;

    @BeforeEach
    void setUp() throws InvocationTargetException, IllegalAccessException, IOException, IntrospectionException {
        object = new TestObject();
        object1 = new TestObject();
        object.setFirstName("mohammad");
        object.setLastName("mohammadi");
        object.setAge(2);

        object1.setFirstName("jalili");
        object1.setLastName("jali");
        object1.setAge(23);
        List<Object> objects = new ArrayList<>();
        objects.add(object);
        objects.add(object1);
        objectToExcel = new ObjectToExcel(objects);
    }

    @Test
    void t() throws IOException {
        Assertions.assertEquals(ExcelWriter.writeToFile(objectToExcel, "/home/mohammad/IdeaProjects/"), true);
    }
}