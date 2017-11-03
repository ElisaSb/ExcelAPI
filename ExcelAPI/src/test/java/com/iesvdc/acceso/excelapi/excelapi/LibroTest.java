/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.util.ArrayList;
import java.util.List;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author matinal
 */
public class LibroTest {
    
    public LibroTest() {
    }
    
    @BeforeClass
    public static void setUpClass() {
    }
    
    @AfterClass
    public static void tearDownClass() {
    }
    
    @Before
    public void setUp() {
    }
    
    @After
    public void tearDown() {
    }

    /**
     * Test of getHojas method, of class Libro.
     */
    @Test
    public void testGetHojas() {
        System.out.println("getHojas");
        Libro instance = new Libro();
        List<Hoja> expResult = new ArrayList<Hoja>();
        Hoja h1 = new Hoja("hoja1",5,5);
        Hoja h2 = new Hoja("hoja2",10,10);
        for(int i=0;i<6;i++){
            for(int j=0;j<6;j++){
                //cell.setCellValue((char)('A'+j)+" "+(i+1)); 
                h1.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        for(int i=0;i<10;i++){
            for(int j=0;j<6;j++){
                h2.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        expResult.add(h1);
        expResult.add(h2);
        instance.addHoja(h1);
        instance.addHoja(h2);
        List<Hoja> result = instance.getHojas();
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of setHojas method, of class Libro.
     */
    @Test
    public void testSetHojas() {
        System.out.println("setHojas");
        List<Hoja> hojas = null;
        Libro instance = new Libro();
        instance.setHojas(hojas);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of getNombreArchivo method, of class Libro.
     */
    /*@Test
    public void testGetNombreArchivo() {
        System.out.println("getNombreArchivo");
        Libro instance = new Libro();
        String expResult = "nuevo.xslx";
        String result = instance.getNombreArchivo();
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }*/

    /**
     * Test of setNombreArchivo method, of class Libro.
     */
    @Test
    public void testSetNombreArchivo() {
        System.out.println("setNombreArchivo");
        String nombreArchivo = "unNombre.xlsx";
        Libro instance = new Libro();
        instance.setNombreArchivo(nombreArchivo);
        assertEquals(nombreArchivo, instance.getNombreArchivo());
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of addHoja method, of class Libro.
     */
    @Test
    public void testAddHoja() {
        System.out.println("addHoja");
        int filas=20, columnas=30;
        Hoja hoja = new Hoja("pepe",filas,columnas);
        for(int i=0;i<filas;i++){
            for(int j=0;j<columnas;j++){
                hoja.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        Libro instance = new Libro();
        instance.addHoja(hoja);
        try{
            assertEquals(instance.indexHoja(0).compare(hoja), true);
        } catch (ExcelAPIException ex){
            fail("No puedo acceder a la hoja");
        }
        // TODO review the generated test code and remove the default call to fail.
       //fail("The test case is a prototype.");
    }

    /**
     * Test of removeHoja method, of class Libro.
     */
    @Test
    public void testRemoveHoja() throws Exception {
        System.out.println("removeHoja");
        int index = 0;
        Libro instance = new Libro("nuevolibro.xlsx");
        Hoja expResult = new Hoja("hoja1",5,5);
        for(int i=0;i<expResult.getFilas();i++){
            for(int j=0;j<expResult.getColumnas();j++){
                expResult.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        instance.addHoja(expResult);
        Hoja result = instance.removeHoja(index);
        assertEquals("Error en removeHoja", result.compare(expResult), true);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of indexHoja method, of class Libro.
     * Ya esta hecho en el testHoja
     */
    /*
    @Test
    public void testIndexHoja() throws Exception {
        System.out.println("indexHoja");
        int index = 0;
        Libro instance = new Libro();
        Hoja expResult = null;
        Hoja result = instance.indexHoja(index);
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }*/

    /**
     * Test of load method, of class Libro.
     */
    @Test
    public void testLoad_0args() {
        System.out.println("load");
        Libro instance = new Libro();
        instance.load();
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of load method, of class Libro.
     */
    @Test
    public void testLoad_String() {
        System.out.println("load");
        String filename = "";
        Libro instance = new Libro();
        instance.load(filename);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of save method, of class Libro.
     */
    @Test
    public void testSave_0args() throws Exception {
        System.out.println("save");
        Libro instance = new Libro("testNuevo.xlsx");
        Hoja h1 = new Hoja("testNuevo1",6,6);
        Hoja h2 = new Hoja("testNuevo2",10,10);
        for(int i=0;i<6;i++){
            for(int j=0;j<6;j++){
                h1.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        for(int i=0;i<10;i++){
            for(int j=0;j<6;j++){
                h2.setDato((char)('A'+j)+" "+(i+1), i, j);
            }
        }
        instance.addHoja(h1);
        instance.addHoja(h2);
        instance.save();
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of save method, of class Libro.
     */
    @Test
    public void testSave_String() throws Exception {
        System.out.println("save");
        String filename = "nuevo.xlsx";
        Libro instance = new Libro();
        instance.save(filename);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }
    
}
