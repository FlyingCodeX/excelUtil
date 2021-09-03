import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.Files;
import java.util.List;

/**
 * @author faye
 * @create 2021-09-2021/9/2-17:00
 * 用于根据excel表格更改文件夹名称
 */
public class ExcelUtil {

    //根据分数给文件夹改名
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
                 fileRename(filePath,target,newName);
             }
         }

     }

     public static void fileRename(String path,String target,String newName) throws FileNotFoundException {
         File file = new File(path);
         File[] files = file.listFiles();
         for(int i=0;i<files.length;i++){
             String oldName = files[i].getName();
             if(oldName.equals(target)){
                 files[i].renameTo(new File(path + "\\"+ newName));
             }
         }

     }

     public static void attributeFile(String chemistryPath,String physicsPath,String excelPath,String srcPath) throws IOException {
         Logger logger = LoggerFactory.getLogger("分配文件日志");

         FileInputStream fis = new FileInputStream(excelPath);
         Workbook workbook = WorkbookFactory.create(fis);
         Sheet sheet = workbook.getSheetAt(0);
         int lastRowNum = sheet.getLastRowNum();

         for(int i=1;i<lastRowNum;i++){
             String fileName = null;
             Row row = sheet.getRow(i);
             String className = row.getCell(4).getStringCellValue();
             int questionName = (int)row.getCell(5).getNumericCellValue();
             String targetDictionaryName  = className + questionName;
             //缺考,无文件
             if((int)row.getCell(3).getNumericCellValue()==-2){
                 logger.warn((long)row.getCell(0).getNumericCellValue() + " " + "该学生缺考");
                 continue;
             }
             else if(row.getCell(3).getNumericCellValue() < 10){
                 fileName = (long)row.getCell(0).getNumericCellValue() + "_" + (int)row.getCell(3).getNumericCellValue();
             }
             else {
                 StringBuilder sb = new StringBuilder();
                 long num = (long)row.getCell(0).getNumericCellValue();
                 String studentName = row.getCell(1).getStringCellValue();
                 fileName = sb.append(num).toString();
             }
             //创建源文件夹和目标文件夹
             File srcFile = new File(srcPath + "\\" + fileName);
             File targetFile = null;
             //如果源文件不存在,则直接跳过
             if(!srcFile.exists()){
                 logger.warn(srcFile.toString() + " " + "文件不存在!");
                 continue;
             }
             //物理和化学文件名的判断
             if(className.equals("物理")){
                 targetFile = new File(physicsPath + "\\" +targetDictionaryName);
             }
             else if (className.equals("化学")) {
                 targetFile = new File(chemistryPath + "\\" + targetDictionaryName);
             }

             FileUtils.copyDirectoryToDirectory(srcFile,targetFile);

             logger.info(srcFile.toString() + " " +"传输完成");
             if(i == lastRowNum-1){
                 logger.info("全部文件传输完成");
             }


         }



     }





}
