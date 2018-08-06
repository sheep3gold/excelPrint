//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//
///**
// * Created by yx on 2018/8/6.
// */
//public class readExcel {
//    public static void main(String[] args) {
////        //定义表头
////        String[] title = {"序号", "姓名", "年龄"};
//////创建excel工作簿
////        XSSFWorkbook workbook = new XSSFWorkbook();
//////创建工作表sheet
////        XSSFSheet sheet = workbook.createSheet();
//////创建第一行
////        XSSFRow row = sheet.createRow(0);
////        XSSFCell cell = null;
//////插入第一行数据的表头
////        for (int i = 0; i < title.length; i++) {
////            cell = row.createCell(i);
////            cell.setCellValue(title[i]);
////        }
//////写入数据
////        for (int i = 1; i <= 10; i++) {
////            XSSFRow nrow = sheet.createRow(i);
////            XSSFCell ncell = nrow.createCell(0);
////            ncell.setCellValue("" + i);
////            ncell = nrow.createCell(1);
////            ncell.setCellValue("user" + i);
////            ncell = nrow.createCell(2);
////            ncell.setCellValue("24");
////        }
//////创建excel文件
//        File file = new File("D://test1.xlsx");
////        try {
////            file.createNewFile();
////            //将excel写入
////            FileOutputStream stream = new FileOutputStream(file);
////            workbook.write(stream);
////            stream.close();
////        } catch (IOException e) {
////            e.printStackTrace();
////        }
//    try{
//        FileInputStream fis=new FileInputStream(file);
//        XSSFWorkbook wookbook = new XSSFWorkbook(fis);
//        // 创建对Excel工作簿文件的引用
//        XSSFSheet sheet1 = wookbook.getSheetAt(0);
//        XSSFRow row1=sheet1.getRow(0);
//        XSSFCell cell1=row1.getCell(1);
//        String cellvalue=cell1.toString();
//        System.out.print(cellvalue);
//    }catch (IOException e){
//        e.printStackTrace();
//    }
////}
//    }
//}
//
