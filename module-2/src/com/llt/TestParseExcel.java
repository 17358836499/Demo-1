package com.llt;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import org.apache.poi.ss.usermodel.Workbook;//对应Excel文档
import org.apache.poi.hssf.usermodel.HSSFWorkbook;//对应xls格式的Excel文档
import org.apache.poi.xssf.usermodel.XSSFWorkbook;//对应xlsx格式的Excel文档
import org.apache.poi.ss.usermodel.Sheet;//对应Excel文档中的一个sheet
import org.apache.poi.ss.usermodel.Row;//对应一个sheet中的一行
import org.apache.poi.ss.usermodel.Cell;//对应一个单元格
import org.apache.poi.ss.usermodel.DateUtil;//时间工具

/**
 *     <dependencies>
 *         <dependency>
 *             <groupId>org.apache.poi</groupId>
 *             <artifactId>poi-ooxml</artifactId>
 *             <version>3.17</version>
 *         </dependency>
 *     </dependencies>
 */


public class TestParseExcel {

    public static void main(String[] args) {

        Workbook wb = null;
        Sheet sheet = null;
        Row row = null;
        List<ArrayList<String>> list = null;//用于存储所有的行记录
        String cellData = null;
        Scanner input = new Scanner(System.in);
        System.out.println("请输入要解析的Excel文件路径:");
        String filePath = input.nextLine();
        System.out.println("解析结果如下-------------->");
        wb = readExcel(filePath);
        if(wb != null){
            //用来存放表中数据
            list = new ArrayList<ArrayList<String>>();
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            //获取第一行
            row = sheet.getRow(0);
            //获取最大列数
            int colnum = row.getPhysicalNumberOfCells();
            for (int i = 0; i<rownum; i++) {
                List<String> rowList = new ArrayList<String>();//一个list对应一行
                row = sheet.getRow(i);
                if(row !=null){
                    for (int j=0;j<colnum;j++){
                        cellData = (String) getCellFormatValue(row.getCell(j));
                        rowList.add(cellData);
                    }
                }else{
                    break;
                }
                list.add((ArrayList<String>) rowList);
            }
        }
        //遍历解析出来的list
        int i = 0;
        for (List<String> oneRow : list) {
            System.out.print("["+i+++"] ");
            for (String data: oneRow) {
                System.out.print(data+" ");
            }
            System.out.println();
        }

    }

    //根据不同的文件后缀的excel相应的生成对应的文件对象
    public static Workbook readExcel(String filePath){
        Workbook wb = null;
        if(filePath==null){
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                return wb = new HSSFWorkbook(is);
            }else if(".xlsx".equals(extString)){
                return wb = new XSSFWorkbook(is);
            }else{
                return wb = null;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    //判断读取出来每一行对应每一列的数据类型并且返回其值
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
                case Cell.CELL_TYPE_NUMERIC:{
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA:{
                    //判断cell是否为日期格式
                    if(DateUtil.isCellDateFormatted(cell)){
                        //转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    }else{
                        //数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING:{
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }
}


