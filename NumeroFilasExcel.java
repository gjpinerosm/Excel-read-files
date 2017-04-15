package LeerExcel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

/**
 * Created by GiovanniJavierRicard on 02/04/2017.
 */
public class NumeroFilasExcel {
    public static void main(String[] args) {
        String data;
        try {
            InputStream is = new FileInputStream("C:\\Users\\GiovanniJavierRicard\\Dropbox\\Develop\\datosXLS.xlsx");
            Workbook libroTrabajo = WorkbookFactory.create(is);
            Sheet LOTE1 = libroTrabajo.getSheetAt(0);
            Iterator rowIter = LOTE1.rowIterator();
            Row fila = (Row) rowIter.next();
            short lastCellNum = fila.getLastCellNum();
            int[] dataCount = new int[lastCellNum];
            int col = 0;
            //rowIter = LOTE1.rowIterator();  //Se quita para no incluir la primera fi-la con el encabezado
            while (rowIter.hasNext()) {
                Iterator cellIter = ((Row) rowIter.next()).cellIterator();
                while (cellIter.hasNext()) {
                    int valor;
                    Cell cell = (Cell) cellIter.next();
                    col = cell.getColumnIndex();
                    dataCount[col] += 1;
                    valor = cell.getColumnIndex();
                }
            }
            is.close();
            System.out.println("Los cantidad de filas son: " + dataCount [0]);
            /*for (int x = 0; x < dataCount.length; x++) {
                System.out.println("col " + x + ": " + dataCount[x]);
            }*/
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }
    }
}
