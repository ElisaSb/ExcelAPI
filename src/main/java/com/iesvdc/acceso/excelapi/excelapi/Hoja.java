/*

 */
package com.iesvdc.acceso.excelapi.excelapi;

/**
 *  Esta clase almacena información del texto de una hoja de cálculo
 * 
 * @author Elisa Serrano Blazquez
 */
public class Hoja {
    private String [][] datos;
    private String nombre;
    private int nFilas;
    private int nColumnas;
    /**
     * Crea una hoja de cálculo nueva
     */
    public Hoja() {
        this.datos = new String[0][0];
        this.nombre = "";
        this.nFilas = 5;
        this.nColumnas = 5;
    }
    /**
     * Crea una hoja nueva de tamaño nFilas por nColumnas
     * @param nFilas el número de filas
     * @param nColumnas el número de celdas que tiene cada fila 
     */
    public Hoja(int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        nombre="";
    }

    /**
     * Crea una hoja nueva con el nombre elegido y tamaño nFilas por nColumnas
     * @param nombre
     * @param nFilas
     * @param nColumnas 
     */
    
    public Hoja(String nombre, int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nombre = nombre;
        this.nColumnas = nColumnas;
        this.nFilas = nFilas;
    }

    /**
     * Obtiene el dato que se encuentra en la fila y columna indicada
     * @param fila
     * @param columna
     * @return 
     */
    
    public String getDato(int fila, int columna) {
        //TO-DO excepciçon si accedemos a una posicion no válida
        return datos[fila][columna];
        
    }
    
    /**
     * Modifica el dato de la fila y columna indicada con el dato indicado
     * @param dato
     * @param fila
     * @param columna 
     */

    public void setDato(String dato, int fila, int columna) {
        //TO-DO excepciçon si accedemos a una posicion no válida
        this.datos[fila][columna] = dato;
    }
    
    /**
     * Obtiene el nombre de la hoja
     * @return 
     */
    
    public String getNombre(){
        return this.nombre;
    }
    
    /**
     * Modifica el nombre de la hoja, por el que se introduzca
     * @param nombre 
     */
    
    public void setNombre(String nombre){
        this.nombre = nombre;
    }
    
    /**
     * Devueleve el número de filas
     * @return 
     */

    public int getFilas() {
        return nFilas;
    }
    
    /**
     * Devuelve el número de columnas
     * @return 
     */

    public int getColumnas() {
        return nColumnas;
    }
    
    /**
     * Compara las hojas que están guardadas con las hojas introducidas
     * @param hoja
     * @return 
     */
    
    public boolean compare(Hoja hoja){
        boolean iguales = true;
        if(this.nColumnas==hoja.getColumnas() 
                && this.nFilas==hoja.getFilas()
                && this.nombre.equals(hoja.getNombre())){
            for(int i=0;i<this.nFilas;i++){
                for(int j=0;j<this.nColumnas;j++){
                    if(!this.datos[i][j].equals(hoja.getDato(i, j))){
                        iguales=false;
                        break;
                    }
                }
                if(!iguales) break;
            }
        } else {
            iguales = false;
        }
        return iguales;
    }
    
}