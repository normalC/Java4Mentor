/*
 * Main.java
 *
 * Created on June 23, 2006, 4:22 PM
 *
 * To change this template, choose Tools | Template Manager
 * and open the template in the editor.
 */

package componentbrowser;
import javax.swing.UIManager;

/**
 *
 * @author jdube
 */
public class Main {
    
    /** Creates a new instance of Main */
    public Main() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch(Exception e) {
            System.out.println("Error setting native LAF: " + e);
        }
        
        ComponentFrame cp = new ComponentFrame();
        cp.setTitle("Component Browser");
        cp.OpenSchematic(args[0]);
        cp.setVisible(true);
    }
    
}
