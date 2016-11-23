/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Logica;

import java.io.File;


/**
 *
 * @author APRENDIZ
 */
public class Porteria {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        Datos.Conexion.conexion();
       //Logica.LListadoAsistencia.diaActual();
        //ManejadoraPantalla.abrirGenerarReportes();
        //LCrearCSV.escribirArchivo();
        //Logica.LImportarExel.importarexcel();
      File f = new File("C:\\Users\\CLAUDIA\\Documents\\porteria\\formato.xlsx");
       if (f.exists()) {
        Logica.LLeerExcel1 Object = new LLeerExcel1(f);
     }

    } 
}
