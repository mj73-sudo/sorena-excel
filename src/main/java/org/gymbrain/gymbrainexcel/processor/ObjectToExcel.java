package org.gymbrain.gymbrainexcel.processor;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.gymbrain.gymbrainexcel.annotations.ExcelDocument;
import org.gymbrain.gymbrainexcel.annotations.ExcelField;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

public class ObjectToExcel {

    Row header = null;
    CellStyle headerStyle = null;
    private byte[] excelByteArray;
    private Validator validator = null;
    private List<Object> objects = new ArrayList<>();
    private String documentName;
    private boolean isFirstObject = true;

    private ObjectToExcel() {
    }

    public ObjectToExcel(List<Object> objects) throws InvocationTargetException, IllegalAccessException, IOException, IntrospectionException {
        validator = new Validator();
        this.objects = objects;
        process();
    }

    private void process() throws InvocationTargetException, IllegalAccessException, IOException, IntrospectionException {
        Object object = objects.get(0);
        Workbook workbook = new XSSFWorkbook();

        ExcelDocument excelDocument = object.getClass().getAnnotation(ExcelDocument.class);
        validator.validateExcelDocument(excelDocument, object);
        documentName = excelDocument.name();

        int row = 1;

        Sheet sheet = workbook.createSheet(documentName);
        initHeader(workbook, sheet);

        for (Object o : objects) {
            processPerObject(o, sheet, row);
            row++;
        }
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        bos.close();
        workbook.close();
        excelByteArray = bos.toByteArray();
    }

    private void processPerObject(Object object, Sheet sheet, int rowNumber) throws InvocationTargetException, IllegalAccessException, IntrospectionException {
        Row row = sheet.createRow(rowNumber);
        BeanInfo beanInfo = Introspector.getBeanInfo(object.getClass());
        List<String> headers = new ArrayList<>();
        Map<String, Method> methodMap = new HashMap<>();
        Arrays.stream(beanInfo.getPropertyDescriptors()).forEach(prop -> {
            if (prop.getName().equals("class")) {
                return;
            }
            String name;
            ExcelField excelField = prop.getReadMethod().getAnnotation(ExcelField.class);
            if (excelField != null) {
                name = excelField.name();
            } else {
                excelField = prop.getWriteMethod().getAnnotation(ExcelField.class);
                if (excelField != null) {
                    name = excelField.name();
                } else {
                    name = prop.getName();
                }
            }
            methodMap.put(name, prop.getReadMethod());
            headers.add(prop.getName());
        });

        AtomicInteger index = new AtomicInteger();
        methodMap.forEach((key, value) -> addHeaderCell(key, index.getAndIncrement()));

        AtomicInteger valueIndex = new AtomicInteger();
        methodMap.forEach((key, method) -> {
            Cell cell = row.createCell(valueIndex.getAndIncrement());
            try {
                cell.setCellValue(String.valueOf(method.invoke(object)));
            } catch (IllegalAccessException | InvocationTargetException e) {
                throw new RuntimeException(e);
            }
        });

        isFirstObject = false;

    }

    private void addHeaderCell(String fieldName, int index) {
        Cell headerCell = header.createCell(index);
        headerCell.setCellValue(fieldName);
        headerCell.setCellStyle(headerStyle);
    }

    private void initHeader(Workbook workbook, Sheet sheet) {
        header = sheet.createRow(0);

        headerStyle = workbook.createCellStyle();

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);
    }

    public String getDocumentName() {
        return documentName;
    }

    public byte[] getBytes() {
        return excelByteArray;
    }
}
