/*
 * ComponentFrame.java
 *
 * Created on June 23, 2006, 4:54 PM
 */

package componentbrowser;
import javax.swing.ListSelectionModel;
import javax.swing.UIManager;
import javax.swing.event.*;
import javax.swing.table.*;
import ViewDraw.*;
import com.jacob.com.*;

/**
 *
 * @author  jdube
 */
public class ComponentFrame extends javax.swing.JFrame {
    
    /** Creates new form ComponentFrame */
    public ComponentFrame() {
        initComponents();
        ComponentTableModel compmodel = new ComponentTableModel();
        jTable1.setModel(compmodel);
        jTable1.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        ListSelectionModel rowSM = jTable1.getSelectionModel();
        rowSM.addListSelectionListener(new ListSelectionListener() {
            public void valueChanged(ListSelectionEvent e) {
                if (e.getValueIsAdjusting()) return;
                
                ListSelectionModel lsm = (ListSelectionModel)e.getSource();
                if (!lsm.isSelectionEmpty()) {
                    int selectedRow = lsm.getMinSelectionIndex();
                    String uid = jTable1.getModel().getValueAt(selectedRow, 0).toString();
                    vdapp.IVdApp_SelectPath(vdapp.IVdApp_getActiveView().IVdView_GetTopLevelDesignName(), uid, 1, 0, false, false);
                    vdapp.IVdApp_ExecuteCommand("zsel");
                }
            }
        });
    }
    
    public void OpenSchematic(String schname) {
        vdapp = new ViewDraw.CVdApp();
        vdapp.IVdApp_setVisible(true);

        ViewDraw.IVdDocs vddocs = vdapp.IVdApp_getDocuments();
        ViewDraw.IVdDoc vddoc = vddocs.IVdDocs_Open(schname);
        ViewDraw.IVdView vdview = vdapp.IVdApp_getActiveView();
        

    }

    /** This method is called from within the constructor to
     * initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is
     * always regenerated by the Form Editor.
     */
    // <editor-fold defaultstate="collapsed" desc=" Generated Code ">//GEN-BEGIN:initComponents
    private void initComponents() {
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane1.setViewportView(jTable1);

        org.jdesktop.layout.GroupLayout layout = new org.jdesktop.layout.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(org.jdesktop.layout.GroupLayout.LEADING)
            .add(jScrollPane1, org.jdesktop.layout.GroupLayout.DEFAULT_SIZE, 400, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(org.jdesktop.layout.GroupLayout.LEADING)
            .add(jScrollPane1, org.jdesktop.layout.GroupLayout.DEFAULT_SIZE, 300, Short.MAX_VALUE)
        );
        pack();
    }// </editor-fold>//GEN-END:initComponents
    
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ComponentFrame().setVisible(true);
            }
        });
    }
    
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
    private ViewDraw.CVdApp vdapp;

    public final class ComponentTableModel extends AbstractTableModel {
        public int getRowCount() {
            ViewDraw.IVdObjs vdobjs = vdapp.IVdApp_getActiveView().IVdView_Query(ViewDraw.mwVdObjectTypeMask.VDM_COMP, ViewDraw.mwVdAllOrSelected.VD_ALL);
            return vdobjs.IVdObjs_getCount();
        }

        public int getColumnCount() {
            return 4;
        }
        
        public String getColumnName(int columnIndex) {
            String result = "";
            switch (columnIndex){
                case 0: result = "UID"; break;
                case 1: result = "Label"; break;
                case 2: result = "RefDes"; break;
                case 3: result = "Symbol"; break;
            }
            return result;
        }
        
        public boolean isCellEditable(int rowIndex, int columnIndex) {
            // allow labels to be edited
            if (columnIndex == 1)
                return true;
            else
                return false;
        }
        
        public Object getValueAt(int rowIndex, int columnIndex) {
            String result = "";
            try {
                ViewDraw.IVdObjs vdobjs = vdapp.IVdApp_getActiveView().IVdView_Query(ViewDraw.mwVdObjectTypeMask.VDM_COMP, ViewDraw.mwVdAllOrSelected.VD_ALL);
                int count = vdobjs.IVdObjs_getCount();
                Dispatch comp = vdobjs.IVdObjs_Item(rowIndex+1);
                
                switch (columnIndex){
                    case 0: result = Dispatch.call(comp, "UID").toString(); break;
                    case 1: result = Dispatch.call(comp, "GetName", new Variant(0)).toString(); break;
                    case 2: result = Dispatch.call(comp, "Refdes").toString(); break;
                    case 3:
                        Object symblock = Dispatch.call(comp, "SymbolBlock").toObject();
                        result = Dispatch.call(symblock, "GetName", new Variant(0)).toString();
                        break;
                }
            } catch(Exception e) {
                System.out.println("Error retrieving component data: " + rowIndex);
            }
            
            return result;
        }
        
        public void setValueAt(Object value, int row, int col) {
            ViewDraw.IVdObjs vdobjs = vdapp.IVdApp_getActiveView().IVdView_Query(ViewDraw.mwVdObjectTypeMask.VDM_COMP, ViewDraw.mwVdAllOrSelected.VD_ALL);
            int count = vdobjs.IVdObjs_getCount();
            Dispatch comp = vdobjs.IVdObjs_Item(row+1);
            Dispatch label = Dispatch.call(comp, "Label").toDispatch();
            if (label.m_pDispatch != 0)
                Dispatch.put(label, "TextString", value.toString());
            else {
                vdapp.IVdApp_getActiveView().IVdView_Application().IVdApp_ExecuteCommand("label " + value.toString() + " no visible local notinverted");
            }
        }

        
    }






}

