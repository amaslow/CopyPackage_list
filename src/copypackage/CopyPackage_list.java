package copypackage;

import java.awt.Desktop;
import java.awt.Rectangle;
import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.Arrays;
import javax.swing.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.util.Iterator;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileNameExtensionFilter;

public class CopyPackage_list extends javax.swing.JFrame {

    public CopyPackage_list() {

        initComponents();
        jProgressBar1.setStringPainted(true);
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        picturesGroup = new javax.swing.ButtonGroup();
        listLabel = new javax.swing.JLabel();
        listTextField = new javax.swing.JTextField();
        listBrowseButton = new javax.swing.JButton();
        jSeparator1 = new javax.swing.JSeparator();
        listStartButton = new javax.swing.JButton();
        jRowCounterLabel = new javax.swing.JLabel();
        jProgressBar1 = new javax.swing.JProgressBar();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Packages copy");
        setMinimumSize(new java.awt.Dimension(470, 150));
        setPreferredSize(new java.awt.Dimension(410, 150));
        setResizable(false);
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        listLabel.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        listLabel.setText("List of items:");
        getContentPane().add(listLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 10, 80, 30));

        listTextField.setFont(new java.awt.Font("Tahoma", 1, 8)); // NOI18N
        listTextField.setToolTipText("Path to Excel list with items");
        getContentPane().add(listTextField, new org.netbeans.lib.awtextra.AbsoluteConstraints(110, 10, 215, 30));
        listTextField.getDocument().addDocumentListener(new DocumentListener() {
            public void changedUpdate(DocumentEvent e) {
                changed();
            }
            public void removeUpdate(DocumentEvent e) {
                changed();
            }
            public void insertUpdate(DocumentEvent e) {
                changed();
            }
            public void changed() {
                if (!listTextField.getText().equals("")){
                    listStartButton.setEnabled(true);
                }
                else {
                    listStartButton.setEnabled(false);
                }
            }
        });

        listBrowseButton.setText("Browse");
        listBrowseButton.setToolTipText("Click to browse Excel list with items");
        listBrowseButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                listBrowseButtonActionPerformed(evt);
            }
        });
        getContentPane().add(listBrowseButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(330, 10, 70, 30));
        getContentPane().add(jSeparator1, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 50, 400, 10));

        listStartButton.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        listStartButton.setText("START");
        listStartButton.setEnabled(false);
        listStartButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                listStartButtonActionPerformed(evt);
            }
        });
        getContentPane().add(listStartButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 60, 90, 40));
        getContentPane().add(jRowCounterLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(170, 480, -1, -1));
        getContentPane().add(jProgressBar1, new org.netbeans.lib.awtextra.AbsoluteConstraints(110, 80, 300, -1));

        pack();
    }// </editor-fold>//GEN-END:initComponents

    String mainfolder = "X:\\Smartwares - Product Content\\PRODUCTS";

    private void label(JFileChooser dest, File source) throws IOException {
        File subdir = dest.getSelectedFile();
        File output = new File(subdir + "\\" + source.getName());

        if (!subdir.exists()) {
            subdir.mkdir();
        }

        InputStream in = new FileInputStream(source);
        OutputStream out = new FileOutputStream(output);
        // Transfer bytes from in to out
        byte[] buf = new byte[1024];
        int len;
        while ((len = in.read(buf)) > 0) {
            out.write(buf, 0, len);
        }
        in.close();
        out.close();
        int i = 0;
        for (i = 0; i < 300; i++) {
            jProgressBar1.setValue(i);
            jProgressBar1.setName("Working...");
            Rectangle progressRect = jProgressBar1.getBounds();
            progressRect.x = 0;
            progressRect.y = 0;
            jProgressBar1.paintImmediately(progressRect);
        }
    }

    private void listBrowseButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_listBrowseButtonActionPerformed
        JFileChooser list = new JFileChooser();
        list.setDialogTitle("Select excel file with list");
        list.setFileSelectionMode(JFileChooser.FILES_ONLY);
        list.setFileFilter(new FileNameExtensionFilter(".xlsx", ".xls", "xls", "xlsx"));
        list.showOpenDialog(null);
        listTextField.setText(list.getSelectedFile().getPath());
    }//GEN-LAST:event_listBrowseButtonActionPerformed

    private void listStartButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_listStartButtonActionPerformed
        try {
            JFileChooser dest = new JFileChooser();
            dest.setDialogTitle("Select destination folder");
            dest.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            dest.showSaveDialog(null);

            String path = listTextField.getText();
            FileInputStream fis1 = null;
            fis1 = new FileInputStream(path);
            XSSFWorkbook wb = new XSSFWorkbook(fis1);
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell = cellIterator.next();
                String item = cell.getStringCellValue();
                final String sapNodot = item.replace(".", "");

                File dir = new File(mainfolder + "\\" + sapNodot);
                File[] foundFiles = dir.listFiles(new FilenameFilter() {
                    public boolean accept(File dir, String name) {
                        return name.startsWith("Package_" + sapNodot + "_");
                    }
                });
                Arrays.sort(foundFiles);
                File source = foundFiles[foundFiles.length - 1];
                label(dest, source);
            }

            jProgressBar1.setName("Finish");
            File subdir = new File(dest.getSelectedFile() + "\\");
            Desktop desktop = Desktop.getDesktop();
            desktop.open(subdir);

        } catch (FileNotFoundException ex) {
            Logger.getLogger(CopyPackage_list.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(CopyPackage_list.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_listStartButtonActionPerformed

    /**
     *
     * @param args
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(CopyPackage_list.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(CopyPackage_list.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(CopyPackage_list.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(CopyPackage_list.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new CopyPackage_list().setVisible(true);
            }
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JProgressBar jProgressBar1;
    private javax.swing.JLabel jRowCounterLabel;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JButton listBrowseButton;
    private javax.swing.JLabel listLabel;
    private javax.swing.JButton listStartButton;
    private javax.swing.JTextField listTextField;
    private javax.swing.ButtonGroup picturesGroup;
    // End of variables declaration//GEN-END:variables

}
