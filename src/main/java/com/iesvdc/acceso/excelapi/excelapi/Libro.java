/*
 * 
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * Esta clase almacena información de libros para generar ficheros de Excel.
 * Un libro se compone de hojas
 * @author Elisa Serrano Blázquez
 */
public class Libro {
    private List <Hoja> hojas;
    private String nombreArchivo;
    
    /**
     * Crea un Libro con el nombre introducido y según elarray de hojas introducidas
     * @param hojas
     * @param nombreArchivo 
     */

    public Libro(List<Hoja> hojas, String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }
    
    /**
     * Crea el Libro con el nombre introducido
     * @param nombreArchivo 
     */

    public Libro(String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = nombreArchivo;
    }
    
    /**
     * Constructor por defecto
     */    

    public Libro() {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }
  
    /**
     * Devuelve el array de Hojas guardadas
     * @return 
     */

    public List<Hoja> getHojas() {
        return hojas;
    }

    /**
     * Modifica el array de hojas por el array introducido
     * @param hojas 
     */
    
    public void setHojas(List<Hoja> hojas) {
        this.hojas = hojas;
    }
    
    /**
     * Devuelve el nombre del Libro
     * @return 
     */
    
    public String getNombreArchivo() {
        return nombreArchivo;
    }
    
    /**
     * Modifica el nombre del Libro según el parámetro introducido
     * @param nombreArchivo 
     */

    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }
    
    /**
     * Añade las Hojas introducidas por parámetro
     * @param hoja 
     */
    
    public void addHoja(Hoja hoja){
        this.hojas.add(hoja);
    }
    
    /**
     * Borra las Hojas que se indiquen por parametro
     * @param index
     * @return
     * @throws ExcelAPIException 
     */
    
    public Hoja removeHoja(int index) throws ExcelAPIException{
        if(index<0 || index>this.hojas.size()){
            throw new ExcelAPIException("Libro::removeHoja(): Posición no válida.");
        }
        return this.hojas.remove(index);
    }
    
    /**
     * devuelve la Hoja según la posicion introducida
     * @param index
     * @return
     * @throws ExcelAPIException 
     */
    
    public Hoja indexHoja(int index) throws ExcelAPIException{
        if(index<0 || index>this.hojas.size()){
            throw new ExcelAPIException("Libro::indexHoja(): Posición no válida.");
        }
        return this.hojas.get(index);
    }
    
    /**
     * Carga las hojas de documento a Hojas del excel
     */
    
    public void load() throws ExcelAPIException{
        File file = new File(getNombreArchivo());
        try {
            FileInputStream in = new FileInputStream(file);
        } catch (IOException ex) {
            throw new ExcelAPIException("Error al cargar el archivo");
        }
        SXSSFWorkbook wb = new SXSSFWorkbook();
        Hoja hoja = wb.createSheet(hoja.getNombre());
        
        
        
    }
    
    /**
     * Carga el archivo introducido por parámetro
     * @param filename
     * @throws ExcelAPIException 
     */
    
    public void load(String filename) throws ExcelAPIException{
        this.nombreArchivo = filename;
        this.load();
    }
    
    /**
     * Guardar el Libro creado
     * @throws ExcelAPIException 
     */
    
    public void save() throws ExcelAPIException{
        SXSSFWorkbook wb = new SXSSFWorkbook();
        for (Hoja hoja : this.hojas) {
            Sheet sh = wb.createSheet(hoja.getNombre());
            for (int i = 0; i < hoja.getFilas(); i++) {
                Row row = sh.createRow(i);
                for (int j = 0; j < hoja.getColumnas(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(hoja.getDato(i, j));                
                }
            }
        }
        
        try {
            FileOutputStream out = new FileOutputStream(this.nombreArchivo);
            wb.write(out);
            out.close();                        
        } catch (IOException ex) {
            throw new ExcelAPIException("Error al guardar el archivo");
        } finally {
            wb.dispose();
        }
    }
    
    public void save(String filename) throws ExcelAPIException{
        this.nombreArchivo = filename;
        this.save();
    }
    
    /**
     * Valida la extension del Libro
     */
    
    private void testExtension(){
        String extension = "/*.xlsx$/";
        nombreArchivo.matches(extension);
    }
}