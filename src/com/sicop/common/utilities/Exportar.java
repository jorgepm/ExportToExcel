package com.sicop.common.utilities;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import javax.mail.MessagingException;
import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Date;
import jxl.write.Label;
import javax.activation.DataHandler;
import java.util.Properties;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.BodyPart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart; //para enviar imagen adjunta
import javax.mail.internet.MimeBodyPart; //para enviar imagen adjunta
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;

public class Exportar {

  private WritableWorkbook libro;
  private int h = 0;
  //private String url;    
  private String user;
  private String pass;
  private String eMailHost;
  private String exportPath;
  private String eMailUser;
  private String eMailPassword;
  private int eMailPort;
  private boolean eMailUseTLS;
  private boolean eMailAuthentication;

  public Exportar(String sqlUser, String sqlPass, String exportPath, String eMailHost, String eMailUser, String eMailPassword, int eMailPort, boolean eMailUseTLS, boolean eMailAuthentication) {
    this.user = sqlUser;
    this.pass = sqlPass;
    this.exportPath = exportPath;
    this.eMailHost = eMailHost;
    this.eMailUser = eMailUser;
    this.eMailPassword = eMailPassword;
    this.eMailPort = eMailPort;
    this.eMailUseTLS = eMailUseTLS;
    this.eMailAuthentication = eMailAuthentication;
  }

  public void aExcel(String sqlURLConexion, String sentencia, String fileName, String ip) throws FileNotFoundException, IOException, SQLException, WriteException, MessagingException, Exception {
    aExcel(sqlURLConexion, sentencia, fileName, null, ip);
  }

  public void aExcel(String sqlURLConexion, String sentencia, String fileName, String eMail, String ip) throws FileNotFoundException, IOException, SQLException, WriteException, MessagingException, Exception {
    crearConexion();
    OutputStream archivoSalida = new FileOutputStream(exportPath + "\\" + fileName + ".xls");
    libro = Workbook.createWorkbook(archivoSalida);
    
    escribirHoja(sqlURLConexion, sentencia);
    libro.write();
    libro.close();
    archivoSalida.flush();
    archivoSalida.close();
    FilesToZip zipear = new FilesToZip();

    zipear.Zippear(fileName + ".xls", exportPath + "\\" + fileName + ".zip");


    if (eMail != null) {
      enviarMail(fileName, eMail, ip);
    }
    //if(eMail!=null) enviarMail(exportPath + "\\" + fileName + ".zip", eMail);
    //Borrar archivos
    File archivo = new File(exportPath + "\\" + fileName + ".xls");
    File archivoZip = new File(exportPath + "\\" + fileName + ".zip");

    archivo.delete();
    archivoZip.delete();

  }

  private void crearConexion() {
    try {
      DriverManager.registerDriver(new com.microsoft.sqlserver.jdbc.SQLServerDriver());
    } catch (SQLException se) {
      System.out.println("ERROR: " + se.getMessage());
    }
  }

  private void escribirHoja(String url, String sentencia) throws SQLException, WriteException {
    Connection conexion = null;
    Statement query = null;
    int n = 0;
    boolean b = true;
    conexion = DriverManager.getConnection(url, user, pass);
    query = conexion.createStatement();
    query.setQueryTimeout(0);
    query.execute(sentencia);
    do {
      escribirResultset(query.getResultSet(), ++n);
    } while (query.getMoreResults());

    query.close();
    conexion.close();
  }

  private void escribirResultset(ResultSet rs, int n) throws SQLException, WriteException {
    WritableSheet hoja = null;
    ResultSetMetaData mdRS = rs.getMetaData();
    int numColumnas = mdRS.getColumnCount();

    WritableFont formatCampos = new WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD, true);
    WritableCellFormat fCell = new WritableCellFormat(formatCampos);
    fCell.setAlignment(Alignment.CENTRE);
    fCell.setBackground(Colour.GRAY_25);


    int a = 0;
    while (rs.next()) {
      if (a == 0 || a >= 50000) {
        h++;
        hoja = libro.createSheet("Hoja " + h, h - 1);
        hoja.addCell(new Label(0, 0, "Query " + n + (a > 0 ? " - Continuacion" : ""), fCell));
        hoja.mergeCells(0, 0, numColumnas - 1, 0);
        for (int c = 1; c <= numColumnas; c++) {
          hoja.addCell(new Label(c - 1, 1, mdRS.getColumnLabel(c), fCell));
        }
        a = 2;
      }
      for (int i = 1; i <= numColumnas; i++) {   // Recorro las columnas
        hoja.addCell(new Label(i - 1, a, rs.getString(i)));
      }
      a++;
    }
    rs.close();
  }

  private void close() throws IOException, WriteException {
    libro.write();
    libro.close();
  }

  private void enviarMail(String fileName, String correo, String ip) throws MessagingException {
    java.util.Date fecha = new Date();

    Properties datos = new Properties();
    datos.setProperty("mail.smtp.host", ip);
    datos.setProperty("mail.smtp.starttls.enable", "" + eMailUseTLS); //si usa TLS o no
    datos.setProperty("mail.smtp.port", "" + eMailPort);
    datos.setProperty("mail.smtp.user", eMailUser);
    datos.setProperty("mail.smtp.password", eMailPassword);
    datos.setProperty("mail.smtps.auth", "" + eMailAuthentication);
    //datos.put("mail.debug", "true");  //para que nos muestre en detalle el proceso
    //datos.put("mail.smtp.socketFactory.port", "25");

    Session session = Session.getDefaultInstance(datos);
    //session.setDebug(true);

    BodyPart texto = new MimeBodyPart();
    texto.setText("Archivo generado: " + fecha);
    BodyPart archivo = new MimeBodyPart();
    archivo.setDataHandler(new DataHandler(new FileDataSource(exportPath + "\\" + fileName + ".zip")));
    archivo.setFileName(fileName + ".zip");
    MimeMultipart correoEnv = new MimeMultipart();
    correoEnv.addBodyPart(texto);
    correoEnv.addBodyPart(archivo);

    MimeMessage mensaje = new MimeMessage(session);
    mensaje.setFrom(new InternetAddress("alertasmobile@btconsultores.com")); //remitente
    mensaje.addRecipient(Message.RecipientType.TO, new InternetAddress(correo)); //destinatario
    mensaje.setSubject("Archivo generado: " + fecha);
    mensaje.setContent(correoEnv);

    Transport t = session.getTransport("smtp");
    t.connect();
    t.sendMessage(mensaje, mensaje.getAllRecipients());
    t.close();
  }
}
