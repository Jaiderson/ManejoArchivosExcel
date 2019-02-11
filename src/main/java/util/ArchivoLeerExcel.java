package util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.google.common.collect.Lists;

/**
 * Clase para leer un archivo de excel en una ruta dada.  
 *
 * @author Jaider Adriam Serrano Sepulveda.
 */

public class ArchivoLeerExcel {

    private String      ruta;
    private File        archivo;
    private Workbook    libro;
    private CellStyle   headerStyle;
    private List<Sheet> hojas;

    public ArchivoLeerExcel(String ruta) {
       this.ruta = ruta;
    }

    public void abrirArchivo() throws EncryptedDocumentException, InvalidFormatException, FileNotFoundException, IOException {
       this.archivo = new File(ruta);
       this.libro = WorkbookFactory.create(new FileInputStream(archivo));
       this.cargarHojas();
       this.crearEstiloCabeceras();
    }

    private void cargarHojas(){
    	this.hojas = Lists.newArrayList();
    	int tam = this.libro.getNumberOfSheets();
    	
    	if(tam > 0) {
    		for(int x=0; x<tam; x++) {
    			this.hojas.add(libro.getSheetAt(x));
    		}
    	}
    }
    
    private void crearEstiloCabeceras() {
        headerStyle = libro.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);    	
    }

    public List<TablaExcel> getHoja(int pos){
    	Sheet hoja = this.hojas.get(pos);
    	return this.extraerInfoHoja(hoja);
    }
    
    public List<TablaExcel> getHoja(String nombre){
    	Sheet hoja = this.libro.getSheet(nombre);
    	return this.extraerInfoHoja(hoja);
    }

    private List<TablaExcel> extraerInfoHoja(Sheet hoja){
    	List<TablaExcel> result = Lists.newArrayList();
    	TablaExcel fila;
    	
    	int tam = hoja.getLastRowNum();
    	for(int x=0; x<=tam; x++) {
    		fila = new TablaExcel();
    		Row row = hoja.getRow(x);
    		
    		for(int cel=0; cel < row.getLastCellNum(); cel++) {
    			fila.addCampo(row.getCell(cel).toString());
    		}
    		result.add(fila);
    	}
    	return result;    	
    }
    
    //**********************   GETTERS ANS SETTERS   **********************\\
    
	public String getRuta() {
		return ruta;
	}

	public void setRuta(String ruta) {
		this.ruta = ruta;
	}

	public File getArchivo() {
		return archivo;
	}

	public void setArchivo(File archivo) {
		this.archivo = archivo;
	}

	public Workbook getLibro() {
		return libro;
	}

	public void setLibro(Workbook libro) {
		this.libro = libro;
	}

	public CellStyle getHeaderStyle() {
		return headerStyle;
	}

	public void setHeaderStyle(CellStyle headerStyle) {
		this.headerStyle = headerStyle;
	}

	public List<Sheet> getHojas() {
		return hojas;
	}

	public void setHojas(List<Sheet> hojas) {
		this.hojas = hojas;
	}
    
}
