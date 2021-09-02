import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.List;

/**
 * @author faye
 * @create 2021-09-2021/9/2-17:00
 * 用于根据excel表格更改文件夹名称
 */
public class ExcelUtil {

    public static void main(String[] args) throws IOException {
        parseFromExcel("H:\\清源镇中理化实验操作考试学生成绩表.xlsx","H:\\清源镇中5.27理化实验录制视频");
    }
     public static void parseFromExcel(String excelPath,String filePath) throws IOException {
         FileInputStream fis = new FileInputStream(excelPath);
         Workbook workbook = WorkbookFactory.create(fis);
         Sheet sheet = workbook.getSheetAt(0);
        //最后一行的行数
         int lastRow = sheet.getLastRowNum();
         for(int i=1;i<lastRow;i++){
             Row row = sheet.getRow(i);
             //获得学号和分数
             long number =  (long)row.getCell(0).getNumericCellValue();
             int grade = (int)row.getCell(3).getNumericCellValue();
             if(grade != 10){
                 StringBuilder sb = new StringBuilder();
                 String target = sb.append(number).toString();
                 sb.append("_" +grade);
                 String newName = sb.toString();
                 fileReName(filePath,target,newName);
             }
         }

     }

     public static void fileReName(String path,String target,String newName) throws FileNotFoundException {
         File file = new File(path);
         File[] files = file.listFiles();
         for(int i=0;i<files.length;i++){
             String oldName = files[i].getName();
             if(oldName.equals(target)){
                 files[i].renameTo(new File(path + "\\"+ newName));
             }
         }

     }

     public static void attributeFile(String chemistryPath,String physicsPath,String excelPath) throws IOException {
         FileInputStream fis = new FileInputStream(excelPath);
         Workbook workbook = WorkbookFactory.create(fis);
         Sheet sheet = workbook.getSheetAt(0);
         int lastRowNum = sheet.getLastRowNum();
         for(int i=1;i<lastRowNum;i++){
             Row row = sheet.getRow(i);
             String className = row.getCell(4).getStringCellValue();
             String questionName = row.getCell(5).getStringCellValue();
             String targetDictionaryName  = className + questionName;
             //文件的名字
             String fileName = null;
             if(row.getCell(3).getNumericCellValue() < 10){
                 fileName = row.getCell(0).getNumericCellValue() + "_" + row.getCell(3).getNumericCellValue();
             }
             else {
                 StringBuilder sb = new StringBuilder();
                 int num = (int)row.getCell(0).getNumericCellValue();
                 fileName = sb.append(num).toString();
             }
         }
     }





}
