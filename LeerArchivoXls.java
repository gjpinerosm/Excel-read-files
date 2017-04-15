package LeerExcel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;

/*
Author: Mandan712
Date: 14/04/2017
Code Refactor: Androstux - Giovanni Piñeros Mora
Country: Bogotá, Colombia
*/

public class LeerArchivoXls {
    public static void main(String[] args) throws Exception {

        HashMap<String, LinkedHashMap<Integer, List>> outerMap = new LinkedHashMap<String, LinkedHashMap<Integer, List>>();
        LinkedHashMap<Integer, List> hashMap = new LinkedHashMap<Integer, List>();

        InputStream ExcelFileToRead = new FileInputStream("C:/Users/GiovanniJavierRicard/Dropbox/Develop/datosXLS.xls");
        HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

        String sheetName = null;

        // Obtener la primera hoja del libro (sheet) de trabajo (workbook).
        for (int i = 0; i < wb.getNumberOfSheets(); i++)
        {
            HSSFSheet sheet = wb.getSheetAt(i);
            sheetName = wb.getSheetName(i);

            // Decidir que fila procesar
            int rowStart = Math.min(1, sheet.getFirstRowNum());
            int rowEnd = Math.max(6, sheet.getLastRowNum());

            for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                Row fila = sheet.getRow(rowNum);
                if (fila == null) {
                    // Agregar mensaje o acción dependiendo de la fila si está vacia
                    continue;
                }

                int lastColumn = Math.max(fila.getLastCellNum(), 5);

                List data = new LinkedList();

                for (int cn = 0; cn < lastColumn; cn++) {
                    Cell celda = fila.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    int FilaError = fila.getRowNum() + 1;
                    if (celda == null || celda.equals(CellType._NONE) || celda.equals(" ")) {
                        String mensajeErrorFila = "Error de dato en excel: " + "Fila(" + FilaError + ")";
                        System.out.println(mensajeErrorFila);
                        data.add(mensajeErrorFila);
                        hashMap.put(fila.getRowNum(), data);
                    } else {
                        // Se agregan a las celdas
                        data.add(celda);
                        hashMap.put(fila.getRowNum(), data);
                    }
                }
            }
            outerMap.put(sheetName, hashMap);
            hashMap = new LinkedHashMap<Integer, List>();
            for (int x = 0; x < outerMap.size(); x++){
                for (int j = 0; j < outerMap.get("Hoja1").size(); j++){
                    for (int z = 0; z < outerMap.get("Hoja1").get(j).size(); z++){
                        System.out.println("El valor es: " + outerMap.get("Hoja1").get(j).get(z));
                    }
                }
            }
        }

    }
}
