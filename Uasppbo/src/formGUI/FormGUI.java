package formGUI;

import javax.swing.table.DefaultTableModel;
import javax.swing.JOptionPane;

import java.io.*;
import java.io.File;
import java.sql.*;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class FormGUI extends javax.swing.JFrame {
    DefaultTableModel model;
    
    public FormGUI() {
        initComponents();
        
        model = (DefaultTableModel) tblMhs.getModel();
        tblMhs.setModel(model);
        
       showData();
    }
    
    private void showData()
    {
        String nama, nim, prodi, telepon;
        try
        {
            Connection con = DBConnect.GetConnection();
            Statement st = con.createStatement();
            
            String query = "select nim, nama, prodi, telepon from mahasiswa";
            ResultSet rs = st.executeQuery(query);
            
            while (rs.next())
            {
                nama    = rs.getString("nama");
                nim    = rs.getString("nim");
                prodi    = rs.getString("prodi");
                telepon    = rs.getString("telepon");
                
                Object[] data = {nim, nama, prodi, telepon};
                model.addRow(data);
            }
        }
        
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }
   
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel1 = new javax.swing.JLabel();
        lblNama = new javax.swing.JLabel();
        lblNim = new javax.swing.JLabel();
        lblProdi = new javax.swing.JLabel();
        lblTelepon = new javax.swing.JLabel();
        tfNama = new javax.swing.JTextField();
        tfNim = new javax.swing.JTextField();
        tfProdi = new javax.swing.JTextField();
        tfTelepon = new javax.swing.JTextField();
        btnTambah = new javax.swing.JButton();
        btnTampil = new javax.swing.JButton();
        btnSimpan = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        tblMhs = new javax.swing.JTable();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel1.setText("Form Mahasiswa");

        lblNama.setText("Nama");

        lblNim.setText("NIM");

        lblProdi.setText("Prodi");

        lblTelepon.setText("Telepon");

        btnTambah.setText("Tambah");
        btnTambah.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnTambahActionPerformed(evt);
            }
        });

        btnTampil.setText("Tampil");
        btnTampil.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnTampilActionPerformed(evt);
            }
        });

        btnSimpan.setText("Simpan");
        btnSimpan.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSimpanActionPerformed(evt);
            }
        });

        tblMhs.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "NIM", "Nama", "Prodi", "Telepon"
            }
        ));
        jScrollPane1.setViewportView(tblMhs);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(lblNama)
                            .addComponent(lblNim)
                            .addComponent(lblProdi)
                            .addComponent(lblTelepon))
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(tfProdi)
                            .addComponent(tfNim)
                            .addComponent(tfNama)
                            .addComponent(tfTelepon, javax.swing.GroupLayout.Alignment.TRAILING)))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(85, 85, 85)
                                .addComponent(btnTambah)
                                .addGap(37, 37, 37)
                                .addComponent(btnTampil)
                                .addGap(40, 40, 40)
                                .addComponent(btnSimpan))
                            .addGroup(layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
            .addGroup(layout.createSequentialGroup()
                .addGap(186, 186, 186)
                .addComponent(jLabel1)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(9, 9, 9)
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblNama)
                    .addComponent(tfNama, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblNim)
                    .addComponent(tfNim, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblProdi)
                    .addComponent(tfProdi, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblTelepon)
                    .addComponent(tfTelepon, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnTambah)
                    .addComponent(btnTampil)
                    .addComponent(btnSimpan))
                .addGap(18, 18, 18)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btnTambahActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnTambahActionPerformed
        if(tfNim.getText().isEmpty() || tfNama.getText().isEmpty() || tfProdi.getText().isEmpty() || tfTelepon.getText().isEmpty())
        {
            JOptionPane.showMessageDialog(null, "Silahkan lengkapi data !");
        }
        else
        {
            String nama, nim, prodi, telepon;
            nama        = tfNama.getText();
            nim        = tfNim.getText();
            prodi        = tfProdi.getText();
            telepon        = tfTelepon.getText();

            Object[] data = {nim, nama, prodi, telepon};
            model.addRow(data);

            try
            {
                Connection con = DBConnect.GetConnection();
                String query = "insert into mahasiswa (nim, nama, prodi, telepon) values (?,?,?,?)";

                PreparedStatement ps = con.prepareStatement(query);
                ps.setString(1, nim);
                ps.setString(2, nama);
                ps.setString(3, prodi);
                ps.setString(4, telepon);
                ps.executeUpdate();
                ps.close();
                System.out.println("Executed!");

                tfNama.setText("");
                tfNim.setText("");
                tfProdi.setText("");
                tfTelepon.setText("");
                
                JOptionPane.showMessageDialog(null, "Input data berhasil.");
            }

            catch(Exception e)
            {
                e.printStackTrace();
            }
        }
    }//GEN-LAST:event_btnTambahActionPerformed

    private void btnTampilActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnTampilActionPerformed
        DefaultTableModel dltTbl = (DefaultTableModel) tblMhs.getModel();
        dltTbl.setRowCount(0);
        
        showData();
    }//GEN-LAST:event_btnTampilActionPerformed

    private void btnSimpanActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSimpanActionPerformed
        LocalDateTime now = LocalDateTime.now();
        int hour = now.getHour();
        int minute = now.getMinute();
        int second = now.getSecond();
        
        String dir = "...\\Documents\\ppbo_UAS.xls";
//        System.out.printf(dir);
        
        File inputWorkbook = new File(dir);
        WritableWorkbook w;
        
        try
        {
            w = Workbook.createWorkbook(inputWorkbook);
            WritableSheet sheet = w.createSheet("DataMahasiswa", 0);
            for(int i = 0; i < model.getColumnCount(); i++)
            {
                sheet.addCell(new Label(i, 0, (i == 0? "NIM":i == 1? "Nama":i == 2? "Prodi":"telepon")));
            }
            
            for(int baris = 0; baris < model.getRowCount(); baris++)
            {
                for(int kolom = 0; kolom < model.getColumnCount(); kolom++)
                {
                    String data = model.getValueAt(baris, kolom).toString();
                    sheet.addCell(new Label(kolom, baris+1, data));
                }
            }
            w.write();
            w.close();
            
            System.out.println("Simpan Data Sukses !");
        }
        
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }//GEN-LAST:event_btnSimpanActionPerformed

    public static void main(String args[])
    {
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new FormGUI().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnSimpan;
    private javax.swing.JButton btnTambah;
    private javax.swing.JButton btnTampil;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JLabel lblNama;
    private javax.swing.JLabel lblNim;
    private javax.swing.JLabel lblProdi;
    private javax.swing.JLabel lblTelepon;
    private javax.swing.JTable tblMhs;
    private javax.swing.JTextField tfNama;
    private javax.swing.JTextField tfNim;
    private javax.swing.JTextField tfProdi;
    private javax.swing.JTextField tfTelepon;
    // End of variables declaration//GEN-END:variables
}
