import jxl.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;

/**
 *
 * Created by Ericwyn on 2017/2/16.
 */
public class FileUtil {
    public static List<Map<String ,Object>> readExcel(String path){
//        path="/storage/emulated/0/test.xls";
        List<Map<String ,Object>> data=new ArrayList<>();
        try {
            InputStream is = new FileInputStream(path);
            Workbook book = Workbook.getWorkbook(is);
//            int num = book.getNumberOfSheets();
            // 获得第一个工作表对象
            Sheet sheet = book.getSheet(0);
//            String schoolName=sheet.getCell(1,3).getContents(); //获取学院信息
            int Rows = sheet.getRows();         //行
            int Cols = sheet.getColumns();      //列
            //列名，行名，读取时候按照（列，行来进行读取）
            int colsOfSchoolName=   0;
            int rolsOfSchoolName=   0;//学院
            int colsOfClassName=    0;
            int rolsOfclassName=    0;//班级
            int colsOfStuName=      0;
            int rolsOfStuName=      0;//姓名
            int colsOfAwards=       0;
            int rolsOfAwards=       0;//所获奖项
            int colsOfMarks=        0;
            int rolsOfMarks=        0;//分值
            for (int i = 0; i < Cols; ++i) {
                for (int j = 0; j < Rows; ++j) {
                    // getCell(Col"列",Row"行")获得单元格的值
                    if(sheet.getCell(i,j).getContents().equals("学院")){
                        colsOfSchoolName=i;
                        rolsOfSchoolName=j;
                    }
                    if(sheet.getCell(i,j).getContents().equals("班级")){
                        colsOfClassName=i;
                        rolsOfclassName=j;
                    }
                    if(sheet.getCell(i,j).getContents().equals("姓名")){
                        colsOfStuName=i;
                        rolsOfStuName=j;
                    }
                    if(sheet.getCell(i,j).getContents().equals("所获奖项")){
                        colsOfAwards=i;
                        rolsOfAwards=j;
                    }
                    if(sheet.getCell(i,j).getContents().equals("分值")) {
                        colsOfMarks = i;
                        rolsOfMarks = j;
                    }
                }
            }
            //妈的源文件里各种中文英文括号混杂，唉我能怎么办啊，我也很绝望啊
            for(int school=rolsOfSchoolName+1;school<(Rows-rolsOfSchoolName);school++){
                Map<String,Object> map=new HashMap<>();
                map.put("school",sheet.getCell(colsOfSchoolName,school).getContents().replace("(","（").replace(")","）").replace(" ", ""));
                map.put("class",sheet.getCell(colsOfClassName,  school).getContents().replace("(","（").replace(")","）").replace(" ", ""));
                map.put("name",sheet.getCell(colsOfStuName,     school).getContents().replace("(","（").replace(")","）").replace(" ", ""));
                map.put("award",sheet.getCell(colsOfAwards,     school).getContents().replace("(","（").replace(")","）").replace(" ", ""));
                map.put("marks",sheet.getCell(colsOfMarks,      school).getContents().replace("(","（").replace(")","）").replace(" ", ""));
                data.add(map);
            }
            book.close();
        } catch (Exception e) {
            System.out.println("文件读取发生错误");
        }
        return data;
    }



}
