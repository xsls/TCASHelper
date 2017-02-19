import com.sun.xml.internal.ws.policy.privateutil.PolicyUtils;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

/**
 *
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

    public String EXCEL_PATH="";
    public String OUT_PATH="";

    public MainUI(){
        TextName.setText("综合评测文档生成器");
        Excel_pathTV.setText("表格文件路径");
        Out_pathTV.setText("文件生成目录");
        ExcelPathButton.setText("选择文件");
        outPathButtom.setText("选择目录");
        createButtom.setText("生成文件");
        ExcelPathButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                EXCEL_PATH=choosePath();
                if(EXCEL_PATH!=""){
                    textField_ExcelPath.setText(EXCEL_PATH);
                }
            }
        });
        outPathButtom.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                OUT_PATH=choosePath();
                if(OUT_PATH!=""){
                    textField_outPath.setText(OUT_PATH);
                }
            }
        });
        createButtom.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                if (EXCEL_PATH!="" && OUT_PATH!=""){
                    try {
                        new MainActivity(EXCEL_PATH,OUT_PATH).mainAct();
                    }catch (Exception ee){

                    }
                }
            }
        });

    }




    public String choosePath() {
        // TODO Auto-generated method stub
        JFileChooser jfc=new JFileChooser();
        jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
        jfc.showDialog(new JLabel(), "选择");
        File file=jfc.getSelectedFile();
        if(file.isDirectory()){
            return file.getAbsolutePath();
        }else if(file.isFile()){
            return file.getAbsolutePath();
        }
        System.out.println(jfc.getSelectedFile().getName());
        return "";
    }


}
