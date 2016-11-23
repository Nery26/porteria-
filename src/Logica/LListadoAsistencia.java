/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Logica;

import Datos.Conexion;
import static Datos.DListadoAsistencia.consultarUsuaFicha;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author APRENDIZ
 */
public class LListadoAsistencia {
    
    public static void diaActual() throws IOException {
        File fileName = new File("C:\\Users\\CLAUDIA\\Documents\\porteria\\formato.xlsx");
      Date fecha = new Date();
      
      //System.out.print(fecha);
     
     Calendar cal1 = Calendar.getInstance();
      int dia = cal1.get(Calendar.DAY_OF_MONTH);
   
        
            ArrayList<LAprendiz> aprendices = new ArrayList<>();
            ResultSet rs11= consultarUsuaFicha("901620");
            
        try {
            while(rs11.next()){
                LAprendiz aprendiz= new LAprendiz(rs11.getString("documento"), rs11.getString("nombres"), rs11.getString("apellido"));
                aprendices.add(aprendiz);
            }
        } catch (SQLException ex) {
            Logger.getLogger(LListadoAsistencia.class.getName()).log(Level.SEVERE, null, ex);
        }

            
        for (int i = 1; i <= dia ; i++) {
            
            for(LAprendiz aprendiz : aprendices){
           
                Statement st;
                ResultSet rs;
                try {
                    st = Conexion.getConect().createStatement();
                    rs = st.executeQuery("select c1.suma_e,c2.suma_s, SEC_TO_TIME(TIMESTAMPDIFF(SECOND,c1.suma_e,c2.suma_s)) AS TIEMPO_ADENTRO from (select  sec_to_time(sum(time_to_sec(hora_ingreso))) as suma_e from ingreso_salida_usu where estado='adentro' and fecha_ingreso='2016/05/30' and documento='"+aprendiz.getDocumento()+"')  AS c1, (select sec_to_time(sum(time_to_sec(hora_ingreso))) as suma_s From ingreso_salida_usu where estado='afuera' and fecha_ingreso='2016/05/30' and documento='"+aprendiz.getDocumento()+"' )as c2 ");
                    
                          
          while(rs.next()){ 
              
           String tiempoAdentro=(rs.getString("TIEMPO_ADENTRO"));
            String [] ta= tiempoAdentro.split(":");
            int tiempoAdentInt= Integer.parseInt(ta[0]);
            

            if(tiempoAdentInt>=3){
                
            InputStream inp;
               try {
                   inp = new FileInputStream(fileName);
                    XSSFWorkbook wb =  new XSSFWorkbook(inp);
                    XSSFSheet sheet = wb.getSheetAt(0);
                    XSSFRow row = sheet.getRow(12);
                    Cell cell = row.getCell(14);

                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    cell.setCellValue("a");  
                    
               } catch (FileNotFoundException ex) {
                   Logger.getLogger(LListadoAsistencia.class.getName()).log(Level.SEVERE, null, ex);
               }
 
            //Logica.LLeerExcel.LLeerExcel(fileName);
                
            // escriba en el archivo la a
  
             System.out.print("asistio"+" ");

            }
            else{
                
            InputStream inp;
               try {
                   inp = new FileInputStream(fileName);
                    XSSFWorkbook wb =  new XSSFWorkbook(inp);
                    XSSFSheet sheet = wb.getSheetAt(0);
                    XSSFRow row = sheet.getRow(12);
                    Cell cell = row.getCell(14);

                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    cell.setCellValue("x");  
                    
               } catch (FileNotFoundException ex) {
                   Logger.getLogger(LListadoAsistencia.class.getName()).log(Level.SEVERE, null, ex);
               }
             //Logica.LLeerExcel.LLeerExcel1(fileName);
               
                 //escriba en el archivo la x
                System.out.print("no asistio" +" ");
            }

           System.out.print(tiempoAdentInt);
        
           System.out.println( " "+aprendiz.getDocumento()+" "+aprendiz.getNombres()+" "+aprendiz.getApellidos());
            
          } 
                } catch (SQLException ex) {
                    Logger.getLogger(LListadoAsistencia.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }
    }
    
    public static void buscarDocumento(String ficha){
        
        ResultSet rs = null;
        
        if (ficha.equals("")){
        rs = Datos.DListadoAsistencia.consultarUsuaFicha();
        }
        else{
        rs = Datos.DListadoAsistencia.consultarUsuaFicha(ficha);
        
        }
        
         DefaultTableModel modelo = new DefaultTableModel();
         Vista.VConsulta.getTblDocum().setModel(modelo);
         
         modelo.addColumn("Documento");

        try {
            while(rs.next()){
                
                Object [] fila = new Object[1];
                
                fila[0]=(rs.getString(1));

                modelo.addRow(fila);
            }
        } catch (SQLException ex) {
            Logger.getLogger(LListadoAsistencia.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
     
}
