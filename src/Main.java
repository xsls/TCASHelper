
import java.util.*;

/**
 * TCASHelp(The comprehensive assessment system Helper)
 * 综合测评系统助手
 * Created by Ericwyn on 2017/2/16.
 */


public class Main {
    private static final String EXCEL_PATH="E:\\Chaos\\IntiliJ Java Project\\TCAS\\testExcelAndWord\\护理学院五四表彰获奖项目.xls".replace("\\\\","/");

    public static void main(String args[]) {
        List<Map<String ,Object>> list_ofData=FileUtil.readExcel(EXCEL_PATH);
        //先对奖项进行分组
        Set<Map<String,Object>> awardsSet=new HashSet<>();
        Set<String> classSet=new HashSet<>();
        //一个固定的awards序列
        for(Map map:list_ofData){
            Map<String,Object> mapOfSet=new LinkedHashMap<>();
            mapOfSet.put("award",map.get("award"));
            mapOfSet.put("marks",map.get("marks"));
            awardsSet.add(map);
        }
        List<Map<String,Object>> awardsList=new ArrayList<>();
        for(Map map:awardsSet){
            awardsList.add(map);
        }
        //一个固定的class序列
        for(Map map:list_ofData){
            String class_=(String )map.get("class");
            classSet.add(class_);
        }
        List<String> classlist=new ArrayList<>();
        for (String str:classSet){
            classlist.add(str);
        }
        classlist.sort(new Comparator<String>() {
            @Override
            public int compare(String o1, String o2) {
                return o1.compareTo(o2);
            }
        });
        //此数组记录各个奖项具体到每个班级的数量
        int[][] arrOfAwardsNumInAClass=new int[awardsSet.size()][classSet.size()];
        for (int i=0;i<arrOfAwardsNumInAClass.length;i++){
            for(int j=0;j<arrOfAwardsNumInAClass[0].length;j++){
                arrOfAwardsNumInAClass[i][j]=0;
            }
        }
        //小天使说同一个奖项同一个班级最高获奖人数就限定在20以内吧
        String[][][] arrOfAll=new String[awardsSet.size()][classSet.size()][20];
        //String[奖项编号][班级编号][班级里面同一个奖项获奖编号]=学生名字；
        for(Map map:list_ofData){
            String className=(String )map.get("class");
            String stuName=(String )map.get("name");
            String award=(String )map.get("award");
            int awardsNum=awardsNum(awardsList,award);
            int classsNum=classsNum(classlist,className);
            arrOfAll[awardsNum][classsNum][arrOfAwardsNumInAClass[awardsNum][classsNum]++]=stuName;
        }
        int sum=0;
        for(int i=0;i<awardsList.size();i++){
            for(int j=0;j<classlist.size();j++){
                for(int k=0;k<arrOfAwardsNumInAClass[i][j];k++){
                        System.out.println(awardsList.get(i).get("award")+","
                                +classlist.get(j)+","+arrOfAll[i][j][k]+","+
                                awardsList.get(i).get("marks"));
                        sum++;
                }
            }
        }
        System.out.println("总共有"+sum+"行");
//        readAll();
    }

    private static void readAll(){
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

    /**
     * @return 返回奖项所在的一维数组编号
     */
    private static int awardsNum(List<Map<String,Object>> awardsList,String awards){
//        int i=0;
        for(int i=0;i<awardsList.size();i++){
            Map<String,Object> map=awardsList.get(i);
            String str=(String )map.get("award");
            if(str.equals(awards)){
                return i;
            }
        }
//        for(Map map:awardsList){
//            String awardsFlag=(String )map.get("award");
//            if(awards.equals(awardsFlag)){
//                return i;
//            }else {
//                i++;
//            }
//        }
        return 10010;
    }

    /**
     * @return 返回班级所在的二维数组编号
     */
    private static int classsNum(List<String> classList,String class_){
        for(int i=0;i<classList.size();i++){
            if(classList.get(i).equals(class_)){
                return i;
            }
        }
        return 10010;
    }





}
