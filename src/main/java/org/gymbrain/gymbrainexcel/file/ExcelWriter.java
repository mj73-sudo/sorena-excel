package org.gymbrain.gymbrainexcel.file;


import org.gymbrain.gymbrainexcel.processor.ObjectToExcel;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {
    public static boolean writeToFile(ObjectToExcel objectToExcel, String path) throws IOException {
        byte[] bytes = objectToExcel.getBytes();
        String documentName = objectToExcel.getDocumentName();
        String fullPath = path + documentName + ".xlsx";
        try (FileOutputStream stream = new FileOutputStream(fullPath)) {
            stream.write(bytes);
        }
        return true;
    }
}
