package org.gymbrain.gymbrainexcel.processor;


import org.gymbrain.gymbrainexcel.annotations.ExcelDocument;
import org.gymbrain.gymbrainexcel.annotations.ExcelField;
import org.gymbrain.gymbrainexcel.error.ExcelProcessorException;

import java.lang.reflect.Method;

class Validator {
    protected void validateExcelDocument(ExcelDocument excelDocument, Object object) {
        if (excelDocument == null)
            throw new ExcelProcessorException(object.getClass().getSimpleName() + " not annotated with ExcelDocument");

        if (excelDocument.name().isEmpty())
            throw new ExcelProcessorException(object.getClass().getSimpleName() + " ExcelDocument annotation not defined value");
    }

    protected boolean isExcelMethod(Method method) {
        ExcelField annotation = method.getAnnotation(ExcelField.class);
        if (annotation == null)
            return false;
        if (annotation.name().isEmpty())
            throw new ExcelProcessorException(method.getName() + " ExcelField annotation not defined value");
        return true;
    }
}
