import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * TCASHelp(The comprehensive assessment system Helper)
 * 综合测评系统助手
 * Created by Ericwyn on 2017/2/16.
 */


public class Main {
    public static void main(String args[]) {
        final String EXCEL_PATH="E:\\Chaos\\IntiliJ Java Project\\TCAS\\testExcelAndWord\\护理学院五四表彰获奖项目.xls".replace("\\\\","/");
        List<Map<String ,Object>> list_ofData=FileUtil.readExcel(EXCEL_PATH);
        for(Map map:list_ofData){
            String schoolName=(String )map.get("school");
            String className=(String )map.get("class");
            String stuName=(String )map.get("name");
            String award=(String )map.get("award");
            String marks=(String )map.get("marks");
            System.out.println(schoolName+","+className+","+stuName+","+award+","+marks);
        }
    }






}
