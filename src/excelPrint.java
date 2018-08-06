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
 * Created by yx on 2018/8/6.
 */
public class excelPrint {
    public static void main(String[] args)throws Exception {
        Scanner input=new Scanner(System.in);
        File file = new File("D://test1.xlsx");
        Map<Integer, String> mymap=new HashMap<Integer, String>();
        FileInputStream fis;
        int key;
        String value="";

        fis=new FileInputStream(file);


        XSSFWorkbook wookbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = wookbook.getSheetAt(0);
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        XSSFRow row=null;
        XSSFCell cell_a=null;
        XSSFCell cell_b=null;

        for (int i = firstRowNum+1; i <= lastRowNum; i++) {
            row = sheet.getRow(i);
            //取得第i行
            cell_a = row.getCell(0);
            cell_a.setCellType(CellType.STRING);
            //取得i行的第二列
            key=Integer.parseInt(cell_a.getStringCellValue().trim());
            //取得i行的第三列
            cell_b = row.getCell(2);
            cell_b.setCellType(CellType.STRING);
            value=cell_b.getStringCellValue().trim();
//            System.out.println(key+" "+value);
            mymap.put(key,value);
        }


//        System.out.print(mymap.get(5));
//        XSSFRow row=sheet.getRow(0);
//        XSSFCell cell=row.getCell(1);
//        String cellvalue=cell.toString();
//        System.out.print(cellvalue);
        // 创建对Excel工作簿文件的引用
        int sumNum=48;
        System.out.print("请输入点名人数");
        int number=input.nextInt();
        Set<Integer> getSet=GetNumber(number,sumNum);
        System.out.println("点名：");
        Iterator it=getSet.iterator();
        while (it.hasNext()){
            int i=(Integer) it.next();
            System.out.println(mymap.get(i));
        }

    }
    public static Set<Integer> GetNumber(int num,int sumNum){
        Random rand = new Random();
        Set<Integer> set=new HashSet<Integer>();
        while (set.size()<num){
            int n=rand.nextInt(sumNum)+1;
            set.add(n);
        }
        return set;
    }
}
