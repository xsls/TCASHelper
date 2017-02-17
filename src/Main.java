
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * TCASHelp(The comprehensive assessment system Helper)
 * 综合测评系统助手
 * Created by Ericwyn on 2017/2/16.
 */


public class Main {
    private static final String EXCEL_PATH="E:\\Chaos\\IntiliJ Java Project\\TCAS\\testExcelAndWord\\护理学院五四表彰获奖项目.xls".replace("\\\\","/");
    private static String WT_PATH="E:\\Chaos\\IntiliJ Java Project\\TCAS\\wordTemplate\\wt4.doc".replace("\\\\","/");
    private static String OUTPUT_PATH="E:\\Chaos\\IntiliJ Java Project\\TCAS\\output\\".replace("\\\\","/");
    private static final List<Map<String ,Object>> list_ofData=FileUtil.readExcel(EXCEL_PATH);
    private static final HashMap<String,Object> marksMap=getMarksMap(list_ofData);

    public static void main(String args[]) throws Exception{

        //先对奖项进行分组
//        Set<Map<String,Object>> awardsSet=new HashSet<>();
        Set<String > awardsSet=new HashSet<>();

        Set<String> classSet=new HashSet<>();
        //一个固定的awards序列
        for(Map map:list_ofData){
            String awards=(String )map.get("award");
            awardsSet.add(awards);
        }
        List<String > awardsList=new ArrayList<>();
        for(String str:awardsSet){
            awardsList.add(str);
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
        //此数组记录各个奖项具体到每个班级的数量，arr[奖项编号][班级数量]
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

        //搞个数组，存储每个奖项总共有多少个班级，凭借这个数据进行word的分页
        int arrOfClassNumInAwards[]=new int[awardsList.size()];
        for(int i=0;i<arrOfClassNumInAwards.length;i++){
            arrOfClassNumInAwards[i]=0;
        }
        for(int i=0;i<awardsList.size();i++){
            for (int j=0;j<classlist.size();j++){
                if(arrOfAwardsNumInAClass[i][j]!=0){
                    arrOfClassNumInAwards[i]+=1;
                }
            }
        }
        //一个测试
//        for(int i=0;i<awardsList.size();i++){
//            if(arrOfClassNumInAwards[i]!=0){
//                System.out.println(""+i+":="+arrOfClassNumInAwards[i]);
//            }
//
//        }
        //遍历输出
        for(int i=0;i<awardsList.size();i++){
            for(int j=0;j<classlist.size();j++){
                for(int k=0;k<arrOfAwardsNumInAClass[i][j];k++){
                        System.out.println(awardsList.get(i)+","
                                +classlist.get(j)+","+arrOfAll[i][j][k]+","+
                                marksOfAwards(awardsList.get(i)));
                }
            }
        }
        wordOutPut(arrOfAll,arrOfAwardsNumInAClass,arrOfClassNumInAwards);
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
    private static int awardsNum(List<String> awardsList,String awards){
//        int i=0;
        for(int i=0;i<awardsList.size();i++){
            String str=awardsList.get(i);
            if(str.equals(awards)){
                return i;
            }
        }
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

    /**
     * 生成Word
     * @param data  数据 data[奖项编号][班级编号][人员编号]
     * @param awardsNum arr[奖项编号][班级数量]=这个班级获得这个奖项的人数
     * @param classInAwards arr[奖项编号]=这个奖项总共有多少个班级获得
     * @throws Exception IO异常
     */
    public static void wordOutPut(String[][][] data,int[][] awardsNum,int[] classInAwards) throws Exception{
        String out_path=OUTPUT_PATH+"test.doc";
        int num=0;
        for(int i=0;i<data.length;i++){
            if(classInAwards[i]%4==0){

            }
        }
//        for(int i=0;i<data.length;i++){         //外层循环，奖项的数量
//            int ye=0;//这个奖项要生成的页数（每4个班一个页面）
//            int wtOfLast=0;//末页需要应用的模板
//            if(ClassInAwards[i]%4==0){          //刚好有4X个班级，全部套用wt4模板就好了
//                //将这一页的数据打包成一个List<Map>
//                for(int j=)
//                ye=ClassInAwards[i]/4;
//                InputStream is = new FileInputStream(WT_PATH);      //导入模板
//                HWPFDocument doc = new HWPFDocument(is);
//
//                for(int j=0;j<ye;j++){
//                    Range range = doc.getRange();                       //获取模板上面的文字
//                    if(j!=0){
//                        InputStream out = new FileInputStream(out_path);      //导入模板
//                        HWPFDocument docOfOut = new HWPFDocument(is);
//                        Range rangeOld=docOfOut.getRange();
//                    }
//
//                    //开始替换文本
//                    range.replaceText("${appleAmt}", "100.00");
//                    range.replaceText("${bananaAmt}", "200.00");
//                    range.replaceText("${totalAmt}", "300.00");
//                    OutputStream os = new FileOutputStream(out_path);
//                    //把doc输出到输出流中
//                    doc.write(os);
//
//                    //关闭输出流
//                    try {
//                        os.close();
//                    } catch (IOException e) {
//                        e.printStackTrace();
//                    }
//                }
//
//                try {
//                    is.close();
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//
//
////                    Range range = doc.getRange();
//                for(int k=0;k<data[i].length;k++){      //遍历班级
//
////                        if (awardsNum[i][k]!=0){
////                            String wtPath=WT_PATH.replace("wt1","wt"+awardsNum[i][j]);
////                            System.out.println(wtPath);
////                            for (int k=0;k<data[i][j].length;k++){
////
////                            }
////                        }
//                }
//
//
//            }else {                             //最后一页不足4个班级，最后一页要套用相应的模板
//                ye=ClassInAwards[i]/4+1;
//                wtOfLast=ClassInAwards[i]%4;
//                for(int j=0;j<data[i].length;j++){
//                    if (awardsNum[i][j]!=0){
//                        String wtPath=WT_PATH.replace("wt1","wt"+awardsNum[i][j]);
//                        System.out.println(wtPath);
//                        for (int k=0;k<data[i][j].length;k++){
//
//                        }
//                    }
//                }
//            }
//        }
    }

    /**
     * 通过奖项-分值的表，查询奖项对应的分值
     * @param awardsName    奖项的名字
     * @return  返回奖项对应的分值
     */
    private static String marksOfAwards(String awardsName){
        return (String)marksMap.get(awardsName);
    }

    /**
     * 制作一个奖项-分值的Map
     * @param list  数据来源
     * @return  返回的分值表
     */
    private static HashMap<String,Object> getMarksMap(List<Map<String ,Object>> list){
        HashMap<String,Object> map=new HashMap<>();
        for(Map mapOfList:list){
            map.put((String )mapOfList.get("award"),mapOfList.get("marks"));
        }
        return map;
    }


    public static void testWrite(String WT_PATH,String WORD_PATH) throws Exception {
        String templatePath = WT_PATH;
        InputStream is = new FileInputStream(templatePath);
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();

        //开始替换文本
        range.replaceText("${appleAmt}", "100.00");
        range.replaceText("${bananaAmt}", "200.00");
        range.replaceText("${totalAmt}", "300.00");
        OutputStream os = new FileOutputStream("D:\\word\\write.doc");
        //把doc输出到输出流中
        doc.write(os);

        //关闭输入输出流
        try {
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}