package test_package;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;

public class excel_creation {

    @Test
    public void f(){
        XSSFWorkbook wb = new XSSFWorkbook();
        File file = new File("C:\\Users\\300965925\\Downloads\\test.xlsx");

//        Random rando = new Random();
//        int rando_int1 = rando.nextInt(10000);
//        int rando_int2 = rando.nextInt(10000);
//        int rando_int3 = rando.nextInt(10000);

        XSSFSheet sh = wb.createSheet("test");

        sh.createRow(0).createCell(0).setCellValue("Email");
        sh.getRow(0).createCell(1).setCellValue("Password");
//        sh.getRow(0).createCell(2).setCellValue("Random Number");

        sh.createRow(1).createCell(0).setCellValue("testemail");
        sh.getRow(1).createCell(1).setCellValue("testpassword");
//        sh.getRow(1).createCell(2).setCellValue(rando_int1);
//
//        sh.createRow(2).createCell(0).setCellValue("user2");
//        sh.getRow(2).createCell(1).setCellValue("pass2");
//        sh.getRow(2).createCell(2).setCellValue(rando_int2);
//
//        sh.createRow(3).createCell(0).setCellValue("user3");
//        sh.getRow(3).createCell(1).setCellValue("pass3");
//        sh.getRow(3).createCell(2).setCellValue(rando_int3);
        try {
            FileOutputStream fos = new FileOutputStream(file);
            wb.write(fos);
            System.out.println("Excel Output:");

            System.out.println(sh.getRow(0).getCell(0).getStringCellValue() + ": " + sh.getRow(1).getCell(0).getStringCellValue());
            System.out.println(sh.getRow(0).getCell(1).getStringCellValue() + ": " + sh.getRow(1).getCell(1).getStringCellValue());
//            System.out.println(sh.getRow(0).getCell(2).getStringCellValue() + ": " + sh.getRow(1).getCell(2));

//            System.out.println("\n" + sh.getRow(0).getCell(0).getStringCellValue() + ": " + sh.getRow(2).getCell(0).getStringCellValue());
//            System.out.println(sh.getRow(0).getCell(1).getStringCellValue() + ": " + sh.getRow(2).getCell(1).getStringCellValue());
//            System.out.println(sh.getRow(0).getCell(2).getStringCellValue() + ": " + sh.getRow(2).getCell(2));
//
//            System.out.println("\n" + sh.getRow(0).getCell(0).getStringCellValue() + ": " + sh.getRow(3).getCell(0).getStringCellValue());
//            System.out.println(sh.getRow(0).getCell(1).getStringCellValue() + ": " + sh.getRow(3).getCell(1).getStringCellValue());
//            System.out.println(sh.getRow(0).getCell(2).getStringCellValue() + ": " + sh.getRow(3).getCell(2));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
