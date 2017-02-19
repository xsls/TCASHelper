import com.sun.deploy.uitoolkit.impl.fx.ui.FXMessageDialog;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 *
 * Created by Ericwyn on 2017/2/19.
 */
public class rukou {
    private Dialog dialog;

    public static void main(String[] args) {
        MenuBar menuBar=new MenuBar();

        JFrame frame = new JFrame("综合测评文档生成器");
        frame.setMenuBar(menuBar);
        Menu about=new Menu("Help");
        menuBar.add(about);
        MenuItem helpItem=new MenuItem("Help");
        MenuItem aboutAuthor=new MenuItem("About");

        about.add(helpItem);
        about.add(aboutAuthor);
        MainUI mainUI=new MainUI();
        javax.swing.JPanel panel=mainUI.rootPanel;
        frame.setContentPane(panel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);

        Dialog dialog=new Dialog(frame, "提示信息-self", true);
        dialog.setBounds(400, 200, 350, 150);
        dialog.setLayout(new FlowLayout());//设置弹出对话框的布局为流式布局
        Label lab = new Label();//创建lab标签填写提示内容
        Button okBut = new Button("close");//创建确定按钮
        lab.setText("陈重围做的一个小软件，为了校团委");

        dialog.add(lab);
        dialog.add(okBut);
        dialog.setVisible(false);

        okBut.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                dialog.setVisible(false);
            }
        });
        helpItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

            }
        });

        about.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                dialog.setVisible(true);
            }
        });


    }
}
