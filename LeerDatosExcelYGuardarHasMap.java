package LeerExcel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by GiovanniJavierRicard on 02/04/2017.
 */
public class LeerDatosExcelYGuardarHasMap {

    public static void main(String[] args) throws Exception {
        File fileName = new File("C:/Users/GiovanniJavierRicard/Dropbox/Develop/datosXLS.xlsx");
        loadExcelLines(fileName);
    }

    public static HashMap loadExcelLines(File fileName){
        HashMap<String, LinkedHashMap<Integer, List>> outerMap = new LinkedHashMap<String, LinkedHashMap<Integer, List>>();
        LinkedHashMap<Integer, List> hashMap = new LinkedHashMap<Integer, List>();

        String sheetName = null;
        // Se crea un ArrayList para guardar los datos leidos desde la hoja de excel.
        // Esta línea realiza la lectura del los encabezados
        // List sheetData = new ArrayList();
        FileInputStream fis = null;
        try
        {
            fis = new FileInputStream(fileName);
            // Create an excel workbook from the file system
            XSSFWorkbook workBook = new XSSFWorkbook(fis);
            // Get the first sheet on the workbook.
            for (int i = 0; i < workBook.getNumberOfSheets(); i++)
            {
                XSSFSheet sheet = workBook.getSheetAt(i);
                // XSSFSheet sheet = workBook.getSheetAt(0);
                sheetName = workBook.getSheetName(i);

                Iterator rows = sheet.rowIterator();
                while (rows.hasNext())
                {
                    XSSFRow row = (XSSFRow) rows.next();
                    Iterator cells = row.cellIterator();

                    List data = new LinkedList();
                    while (cells.hasNext())
                    {
                        XSSFCell cell = (XSSFCell) cells.next();

                        Boolean valor = isCellEmpty(cell);

                        if (valor.equals(false)){
                            data.add(cell);
                            //cell.setCellType(CellType.STRING);
                            hashMap.put(row.getRowNum(), data);
                            for (i = 0; i < data.size(); i++) {
                                data.get(i);
                            }
                        } else {
                            data.add("Celda vacia");
                            hashMap.put(row.getRowNum(), data);
                            for (i = 0; i < data.size(); i++) {
                                data.get(i);
                        }
                      }
                        //cell.setCellType(Cell.CELL_TYPE_STRING);

                    }

                    data.size();
                    //sheetData.add(data);
                }
                    outerMap.put(sheetName, hashMap);
                    hashMap = new LinkedHashMap<Integer, List>();

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
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }
        return outerMap;
    }

    public static boolean isCellEmpty(final XSSFCell cell) {
        if (cell == null || cell.equals(CellType.BLANK) || cell.equals(CellType._NONE)) {
            return true;
        }

        if (cell.equals("")) {
            return true;
        }
        return false;
    }
}
