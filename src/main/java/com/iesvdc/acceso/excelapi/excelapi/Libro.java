/*
 * 
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Esta clase almacena información de libros para generar ficheros de Excel.
 * Un libro se compone de hojas
 * @author Elisa Serrano Blázquez
 */
public class Libro {
    private List <Hoja> hojas;
    private String nombreArchivo;
    
    /**
     * Crea un Libro con el nombre introducido y según el array de hojas introducidas
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
     * Importa una hoja de apachePOI a una de nuestras hojas
     */
       
    public void load() {
        try {
            int numeroFilas=0;
            int numeroColumnas=0;
            FileInputStream file = new FileInputStream(new File(getNombreArchivo()));
            XSSFWorkbook libro = new XSSFWorkbook(file);
            // Iterador de filas
            for (Sheet hojaPOI : libro) {
                Iterator<Row> rowIterator = hojaPOI.iterator();
                while (rowIterator.hasNext()) {
                    numeroFilas++;
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
                        numeroColumnas++;
                    }
                }
                Hoja hoja = new Hoja(hojaPOI.getSheetName(), numeroFilas, numeroColumnas);
                for (int i = 0; i < hoja.getFilas(); i++) {
                    Row filaPOI = hojaPOI.getRow(i);
                    for (int j = 0; j < hoja.getColumnas(); j++) {
                        Cell celdaPOI = filaPOI.createCell(j);
                        //celdaPOI.setCellValue(hoja.getDato(i, j)); 
                        switch (celdaPOI.getCellTypeEnum()) {
                            //si es cadena
                            case STRING:
                                hoja.setDato(celdaPOI.getStringCellValue(),i,j);
                                break;
                            //si es numerico
                            case NUMERIC:
                                hoja.setDato(celdaPOI.getNumericCellValue()+"",i,j);
                                break;
                            //si es una formula
                            case FORMULA:
                                hoja.setDato(celdaPOI.getCellFormula(),i,j);
                                break;
                            case BOOLEAN:
                            //en el caso de Boolean
                                hoja.setDato(celdaPOI.getBooleanCellValue()+"",i,j);
                                break;
                            default:
                                hoja.setDato("", i, j);
			}
                    }
                }
            }
            
        } catch (IOException ex) {
            Logger.getLogger(Libro.class.getName()).log(Level.SEVERE, null, ex);
        }
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