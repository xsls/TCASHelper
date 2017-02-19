
import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

/**
 *程序的主UI
 * Created by Ericwyn on 2017/2/19.
 */
public class MainUI{
    public JButton createButtom;
    public JTextField textField_ExcelPath;
    public JButton ExcelPathButton;
    public JTextField textField_outPath;
    public JButton outPathButtom;
    public JPanel rootPanel;
    public JPanel backPanel;
    public JPanel FileChoosePanel;
    public JPanel excel_pathPanel;
    public JLabel Excel_pathTV;
    public JPanel out_pathPanel;
    public JLabel Out_pathTV;
    public JPanel TopText;
    public JLabel TextName;
    private JTextField textField_wtPaht;
    private JButton chooseWtPathButton;
    private JLabel wt_pathTV;
    private JLabel input_schoolName;
    private JTextField schoolNameInput;

    public String EXCEL_PATH="";
    public String OUT_PATH="";
    public String WT_PATH="";
    public String DIR_NAME="";

    private int DIRECTORIES_ONLY=1;
    private int FILES_ONLY=0;

    public MainUI(){

        TextName.setText("综合评测文档生成器");

        input_schoolName.setText("输入学院名称");

        wt_pathTV.setText("模板文件目录");
        chooseWtPathButton.setText("选择目录");

        Out_pathTV.setText("文档生成目录");
        outPathButtom.setText("选择目录");


        Excel_pathTV.setText("选择表格文件");
        ExcelPathButton.setText("选择文件");

        createButtom.setText("生成文件");


        chooseWtPathButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                WT_PATH=choosePath(DIRECTORIES_ONLY)+"wt4.doc";
                if(WT_PATH!=""){
                    textField_wtPaht.setText(WT_PATH);
                }
            }
        });
        ExcelPathButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                EXCEL_PATH=choosePath(FILES_ONLY);

                if(EXCEL_PATH!=""){
                    textField_ExcelPath.setText(EXCEL_PATH);
                }
            }
        });
        outPathButtom.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                OUT_PATH=choosePath(DIRECTORIES_ONLY)+schoolNameInput.getText();
                if(OUT_PATH!=""){
                    textField_outPath.setText(OUT_PATH+"\\");
                }
            }
        });
        createButtom.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                System.out.println(EXCEL_PATH);
                System.out.println(OUT_PATH);
                System.out.println(WT_PATH);
                if (!EXCEL_PATH.equals("") && !OUT_PATH.equals("") && !WT_PATH.equals("")){
                    File file=new File(OUT_PATH);
                    try {
                        if(!file.exists()) {
                            System.out.println(OUT_PATH+"不存在");
                            file.mkdir();
                        } else {
                            System.out.println(OUT_PATH+"目录存在");
                        }
                        new MainActivity(EXCEL_PATH,OUT_PATH+"\\",WT_PATH,schoolNameInput.getText()).mainAct();
                        System.runFinalizersOnExit(true);
                    }catch (Exception ee){
                        System.out.println(ee.toString());
                    }
                    clear();
                }
            }
        });

    }

    private void clear(){
        OUT_PATH="";
        EXCEL_PATH="";
        schoolNameInput.setText("");
        textField_ExcelPath.setText("");
        textField_outPath.setText("");
    }

    public String choosePath(int chooserType) {
        // TODO Auto-generated method stub
        JFileChooser jfc=new JFileChooser();
        jfc.setFileSelectionMode(chooserType );
        jfc.showDialog(new JLabel(), "选择");
        File file=jfc.getSelectedFile();
        if(file.isDirectory()){
            return file.getAbsolutePath()+"\\";
        }
        System.out.println(jfc.getSelectedFile().getName());
        return file.getAbsolutePath();
    }


}
