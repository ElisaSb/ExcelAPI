/*
 * 
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
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
     * 
     * @param hojas
     * @param nombreArchivo 
     */
    
    public Libro(List<Hoja> hojas, String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }

    public Libro(String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = nombreArchivo;
    }
    

    public Libro() {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }
  

    public List<Hoja> getHojas() {
        return hojas;
    }

    public void setHojas(List<Hoja> hojas) {
        this.hojas = hojas;
    }
    
    public String getNombreArchivo() {
        return nombreArchivo;
    }

    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }
    
    public void addHoja(Hoja hoja){
        this.hojas.add(hoja);
    }
    
    public Hoja removeHoja(int index) throws ExcelAPIException{
        if(index<0 || index>this.hojas.size()){
            throw new ExcelAPIException("Libro::removeHoja(): Posición no válida.");
        }
        return this.hojas.remove(index);
    }
    
    public Hoja indexHoja(int index) throws ExcelAPIException{
        if(index<0 || index>this.hojas.size()){
            throw new ExcelAPIException("Libro::indexHoja(): Posición no válida.");
        }
        return this.hojas.get(index);
    }
    
    public void load(){
        
    }
    
    public void load(String filename){
        this.nombreArchivo = filename;
        this.load();
        
        //Carlos
        
        /*public void load(){
        File file =  new File(getNombreArchivo());
        FileInputStream fileEntrada = new FileInputStream(file);
        try {
            fileEntrada.close();
        } catch (IOException ex) {
            Logger.getLogger(Libro.class.getName()).log(Level.SEVERE, null, ex);
        }
        SXSSFWorkbook wb = new SXSSFWorkbook(fileEntrada);
        for(int i = 0; i < wb.getSheets(); i++){
            Sheet sh = wb.getSheetAt(i);
            int filas = sh.getRows();
            int columnas = 0;
            for(int j = 0; j < filas; j++){
                if(columnas < sh.getCells()){
                    columnas = sh.getCells();
                }
                Hoja hoja = new Hoja(sh.getName,filas,columnas);
                for(int k = 0; k < filas; k++){
                    Row row = sh.getRow(k);
                    for(int z = 0; z < columnas; z++){
                        Cell cell = sh.getCell(z);
                        switch (cell.getCellType()){
                            case STRING: 
                                hoja.setDato(cell.getStringCellValue(),k,z);
                                break;
                            case NUMERIC:
                                hoja.setDato(cell.getStringCellValue(),k,z);
                                break;
                            case FORMULA:
                                hoja.setDato(cell.getStringCellValue(),k,z);
                                break;
                            case BOOLEAN:
                                hoja.setDato(cell.getStringCellValue(),k,z);
                                break;
                            default:
                                hoja.setDato("",i,j);
                        }
                    }
                }
            }
        }
    }*/
    }
    
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
    
    private void testExtension(){
        String extension = "/*.xlsx$/";
    }
}