
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class main {
    public static void main(String[] args) {
        try {
            // สร้าง object ของ excel
            XSSFWorkbook wb = new XSSFWorkbook();
            // สร้าง sheet
            XSSFSheet sheet = wb.createSheet("กหดก");
            // สร้างแถวแรก การนับแถวเริ่มจาก 0,1,2....
            XSSFRow row = sheet.createRow((short)0);
            // สร้าง cell แรกแล้วใส่ค่าลงไป การนับcellเริ่มจาก 0,1,2....
            XSSFCell cell = row.createCell(0);
            cell.setCellValue("Hello Excel.");

            // path ของไฟล์
            FileOutputStream out = new FileOutputStream("hello.xlsx");
            wb.write(out);
            wb.close();
            out.close();
            System.out.println("Excel created successfully");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
