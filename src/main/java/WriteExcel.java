import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteExcel {

    private static Object Static;

    public static void main(String[] args) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Student data");

        int id =1;

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
        data.put("2", new Object[] {"00"+(id++), "Leslie", "Palmeri"});
        data.put("3", new Object[] {"00"+(id++), "Katie", "Wilhelm"});
        data.put("4", new Object[] {"00"+(id++), "Kelly", "Marron"});
        data.put("5", new Object[] {"00"+(id++), "Evan", "Victor"});
        data.put("6", new Object[] {"00"+(id++), "Robert", "Grimes"});
        data.put("7", new Object[] {"00"+(id++), "Gennifer ", "Gonzales"});
        data.put("8", new Object[] {"00"+(id++), "Tyrrel", "Lee"});
        data.put("9", new Object[] {"00"+(id++), "Luis", "Otero"});
        data.put("10", new Object[] {"00"+(id++), "Randy", "Holt"});
        data.put("11", new Object[] {"00"+(id++), "Emily", "Higgins"});

        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
                Cell cell = row.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            FileOutputStream out = new FileOutputStream(new File("WriteExcel.xlsx"));
            workbook.write(out);
            out.close();

            System.out.println("WriteExcel.xlsx written successfully on disk.");

        }
        catch (Exception e)
        {
            e.printStackTrace();
        }



    }
}
