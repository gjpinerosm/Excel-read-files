package LeerExcel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;


public class LeerExcelXLSX {

    public static void main(String[] args) throws Exception {
      File archivo = new File("C://Users/GiovanniJavierRicard/Dropbox/Develop/datosXLS.xlsx");
        FileInputStream archivoUsar = new FileInputStream(archivo);
        // Se crea el objeto que tendra el libro de excel

        XSSFWorkbook libroTrabajo = new XSSFWorkbook(archivoUsar);
        XSSFSheet LOTE1 = libroTrabajo.getSheetAt(0);

        String hUnoFilaUnoCeldaCero = LOTE1.getRow(1).getCell(0).getStringCellValue();
        System.out.print("Datos desde el excel: " + hUnoFilaUnoCeldaCero + "\n");
        String hUnoFilaUnoCeldaUno = LOTE1.getRow(1).getCell(1).getStringCellValue();
        System.out.print("Datos desde el excel: " + hUnoFilaUnoCeldaUno + "\n");

        String hDosFUnoCCero = LOTE1.getRow(2).getCell(0).getStringCellValue();
        System.out.print("Datos desde el excel: " + hDosFUnoCCero + "\n");
        String hDosFUnoCUno = LOTE1.getRow(2).getCell(1).getStringCellValue();
        System.out.print("Datos desde el excel: " + hDosFUnoCUno);
    }
}

