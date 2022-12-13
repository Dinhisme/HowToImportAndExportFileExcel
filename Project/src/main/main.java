package main;

import entity.Product;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Dinhisme
 */
public class main extends javax.swing.JFrame {

    List<Object[]> listExport = new ArrayList<>();
    
    //Remember to add the library in the folder

    /**
     * Creates new form main
     */
    public main() {
        initComponents();
    }

    public void ImportFileExcel() {
        try {
            JFileChooser fc = new JFileChooser();
            fc.showOpenDialog(null);
            File f = fc.getSelectedFile();
            if (f == null) {                                 //check if they want to choose the file or not
                return;
            }
            String path = f.getAbsoluteFile().toString();    //check if the file is excel or not
            if (!path.contains(".xlsx")) {                   //this is not the excel file
                JOptionPane.showMessageDialog(this, "This is not an excel file !!!");
            } else {                                         //this is the excel file
                DefaultTableModel model = (DefaultTableModel) tblData.getModel();
                model.setRowCount(0);
                FileInputStream fis = new FileInputStream(f);
                XSSFWorkbook wb = new XSSFWorkbook(fis);     //remember to add the library so you can use this!!
                XSSFSheet sheet = wb.getSheetAt(0);          //this thing create a sheet start from the cell you want, right here i want it from 0
                Iterator<Row> rowIter = sheet.iterator();    //this like a loop to push the data from the sheet
                int i = 0;                                   //i dont want to get the first row so i create the count i
                Product item = new Product();
                while (rowIter.hasNext()) {
                    List<Product> list = new ArrayList<>();  //create a list to add data
                    Row row = rowIter.next();
                    Iterator<Cell> cellIter = row.iterator(); //this is a loop to push the data from the row
                    if (i != 0) {                             //pass it if this is the first row
                        while (cellIter.hasNext()) {
                            Cell cell = cellIter.next();
                            switch (cell.getColumnIndex()) {  //every case is a cell
                                case 0:                       //get the data from the first cell of the row
                                    item.setIdProduct(String.valueOf(cell));
                                    break;
                                case 1:                       //get the data from the second cell of the row
                                    item.setProduct(String.valueOf(cell));
                                    break;
                                case 2:
                                    item.setType(String.valueOf(cell));
                                    break;
                                case 3:
                                    item.setBrand(String.valueOf(cell));
                                    break;
                            }
                        }
                        list.add(item);
                        for (Product data : list) {
                            Object[] rowData = {data.getIdProduct(), data.getProduct(), data.getType(), data.getBrand()};
                            model.addRow(rowData);
                            listExport.add(rowData);           //add data to this list so you can export data from the table
                        }
                    }
                    i++;
                }
                fis.close();
                JOptionPane.showMessageDialog(this, "Import success!");
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error!!!");
            System.out.println(ex);
        }
    }

    public void ExportFileExcel() {
        try {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.createSheet("Product");
            XSSFRow row = null;
            Cell cell = null;
            row = sheet.createRow(0);                       //create a sheet at 0

            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue("ID PRODUCT");                //set title data at the first row

            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue("PRODUCT");                   //set title data at the first row

            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue("TYPE");                      //set title data at the first row

            cell = row.createCell(3, CellType.STRING);
            cell.setCellValue("BRAND");                     //set title data at the first row
            int i = 0;
            for (Object[] item : listExport) {

                row = sheet.createRow(i + 1);               //create a sheet at 1 to set the data right place

                cell = row.createCell(0, CellType.STRING);  //create a cell at 0 at the values of row
                cell.setCellValue(String.valueOf(item[0]));

                cell = row.createCell(1, CellType.STRING);
                cell.setCellValue(String.valueOf(item[1]));

                cell = row.createCell(2, CellType.STRING);
                cell.setCellValue(String.valueOf(item[2]));

                cell = row.createCell(3, CellType.STRING);
                cell.setCellValue(String.valueOf(item[3]));
                i++;
            }

            JFileChooser fc = new JFileChooser();
            fc.showOpenDialog(null);
            File f = fc.getSelectedFile();
            String path = f.getAbsoluteFile().toString();
            String file = f.getAbsolutePath();
            if (!path.contains(".xlsx")) {
                file = f.getAbsolutePath() + ".xlsx";
            }
            try {
                FileOutputStream fis = new FileOutputStream(file);
                wb.write(fis);
                fis.close();
                JOptionPane.showMessageDialog(this, "Export success!");
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Error!!!");
                System.out.println(ex);
            }

        } catch (Exception ex) {
            System.out.println(ex);
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jPanel2 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tblData = new javax.swing.JTable();
        btnImport = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        btnExport = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

        jPanel2.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

        tblData.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        tblData.setModel(new javax.swing.table.DefaultTableModel(
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
        tblData.setGridColor(new java.awt.Color(0, 0, 0));
        tblData.setSelectionBackground(new java.awt.Color(0, 0, 0));
        tblData.setShowGrid(true);
        jScrollPane1.setViewportView(tblData);

        btnImport.setText("Import");
        btnImport.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnImportActionPerformed(evt);
            }
        });

        jLabel1.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jLabel1.setText("Some text field or something");

        btnExport.setText("Export");
        btnExport.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExportActionPerformed(evt);
            }
        });

        jLabel2.setFont(new java.awt.Font("VNI-Korin", 0, 18)); // NOI18N
        jLabel2.setText("Dinhisme");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(77, 77, 77)
                .addComponent(jLabel1)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 554, Short.MAX_VALUE)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                        .addGap(0, 0, Short.MAX_VALUE)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                                .addComponent(btnImport, javax.swing.GroupLayout.PREFERRED_SIZE, 83, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(btnExport, javax.swing.GroupLayout.PREFERRED_SIZE, 83, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addComponent(jLabel2, javax.swing.GroupLayout.Alignment.TRAILING))))
                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel2)
                .addGap(37, 37, 37)
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 61, Short.MAX_VALUE)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnImport)
                    .addComponent(btnExport))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 178, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(11, 11, 11)
                .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGap(11, 11, 11))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void btnImportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnImportActionPerformed
        ImportFileExcel();
    }//GEN-LAST:event_btnImportActionPerformed

    private void btnExportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExportActionPerformed
        ExportFileExcel();
    }//GEN-LAST:event_btnExportActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new main().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnExport;
    private javax.swing.JButton btnImport;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable tblData;
    // End of variables declaration//GEN-END:variables
}
