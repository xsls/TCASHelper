import javax.swing.*;
import java.awt.*;

/**
 *
 * Created by Ericwyn on 2017/2/19.
 */
public class rukou {
    public static void main(String[] args) {
        JFrame frame = new JFrame("MainUI");
        MainUI mainUI=new MainUI();
        javax.swing.JPanel panel=mainUI.rootPanel;
        frame.setContentPane(panel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }
}
