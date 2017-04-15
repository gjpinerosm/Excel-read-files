package LeerExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class LeerDatosExcelYGuardarHasMapv2 {

        public static void main(String[] args) throws Exception {
            File fileName = new File("C:/Users/GiovanniJavierRicard/Dropbox/Develop/datosXLS.xls");
            loadExcelLines(fileName);
        }

        public static HashMap loadExcelLines(File fileName){
            HashMap<String, LinkedHashMap<Integer, List>> outerMap = new LinkedHashMap<String, LinkedHashMap<Integer, List>>();
            LinkedHashMap<Integer, List> hashMap = new LinkedHashMap<Integer, List>();

            String sheetName = null;
            // Se crea un ArrayList para guardar los datos leidos desde la hoja de excel.
            // Esta línea realiza la lectura del los encabezados
            FileInputStream fis = null;
            try
            {
                fis = new FileInputStream(fileName);
                // Crear un libro de trabajo de excel (workbook) desde el archivo del sistema
                XSSFWorkbook workBook = new XSSFWorkbook(fis);

                // Obtener la primera hoja del libro (sheet) de trabajo (workbook).
                for (int i = 0; i < workBook.getNumberOfSheets(); i++)
                {
                    XSSFSheet sheet = workBook.getSheetAt(i);
                    sheetName = workBook.getSheetName(i);

                    // Decidir que fila procesar
                    int rowStart = Math.min(1, sheet.getFirstRowNum());
                    int rowEnd = Math.max(6, sheet.getLastRowNum());

                    for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                        Row row = sheet.getRow(rowNum);
                        if (row == null) {
                            // Agregar mensaje o acción dependiendo de la fila si está vacia
                            continue;
                        }

                        int lastColumn = Math.max(row.getLastCellNum(), 5);

                        List data = new LinkedList();

                        for (int cn = 0; cn < lastColumn; cn++) {
                            Cell celda = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                            int FilaError = row.getRowNum() + 1;
                            if (celda == null || celda.equals(CellType._NONE) || celda.equals(" ")) {
                                String mensajeErrorFila = "Error de dato en excel: " + "Fila(" + FilaError + ")";
                                System.out.println(mensajeErrorFila);
                                data.add(mensajeErrorFila);
                                hashMap.put(row.getRowNum(), data);
                            } else {
                                // Se agregan a las celdas
                                data.add(celda);
                                hashMap.put(row.getRowNum(), data);
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
            catch (IOException e)
            {
                e.printStackTrace();
            }
            finally
            {
                if (fis != null)
                {
                    try
                    {
                        fis.close();
                    }
                    catch (IOException e)
                    {
                        e.printStackTrace();
                    }
                }
            }
            return outerMap;
        }
}
