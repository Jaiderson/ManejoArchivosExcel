package util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Lists;

/**
 * Clase para crear un archivo de excel en una ruta dada.
 *
 * @author Jaider Adriam Serrano Sepulveda.
 */
public class ArchivoEscribirExcel {
    
    private final String ruta;
    private File archivo;
    private Workbook libro;
    private CellStyle headerStyle;
    private List<Sheet> hojas;
    
    /**
     * 
     * @param ruta Ruta en la cual se creara el archivo .xlsx debe contener el nombre dado al archivo.
     *             Ejemplo: C:/Tempo/Prueba.xlsx.
     */
    public ArchivoEscribirExcel(String ruta){
        this.ruta = ruta;
        this.hojas = Lists.newArrayList();
    }
    
    /**
     * Metodo para crear las intancias de los objetos File y XSSFWorkbook, ademas se crea el estilo para la cabacera
     * de las hojas que contengan libro de excel.
     * 
     * @return True si los objetos File y Workbook fueron creados correctamente.
     */
    public boolean crear() {
        this.archivo = new File(ruta);
            
        if(this.archivo != null){
            libro = new XSSFWorkbook();
            headerStyle = libro.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);            
            return true;
        }        
        return false;
    }
    
    /**
     * Dado el nombre de la hoja crea una hoja nueva y la agrega a la lista de hojas.
     * 
     * @param nombre Nombre de la nueva hoja de calculo que se agregarara al libro, 
     *        este nombre debe ser diferente al de las hojas existentes.
     * @return Hoja de calculo creada o null si no se pudo crear.
     */
    public Sheet crearHoja(String nombre){
        if(libro != null && nombre != null && !nombre.isEmpty() && !exist(nombre)){
            Sheet hoja = libro.createSheet(nombre);
            this.hojas.add(hoja);
            return hoja;
        }
        return null;
    }
    
    /**
     * 
     * @param name Nombre de la nueva hoja a crear para verificar si existe alguna otra hoja con este mismo nombre.
     * @return True si ya existe otra hoja con el mimo nombre.
     */
    private boolean exist(String name){
    	if(!this.hojas.isEmpty()){
            if (this.hojas.stream().anyMatch((hoja) -> (hoja.getSheetName().equals(name)))) {
                return true;
            }
    	}
   	return false;
    }
    
    /**
     * 
     * 
     * @param nombre Nombre de la nueva hoja de calculo que se agregarara al libro, 
     *        este nombre debe ser diferente al de las hojas existentes.
     * @param cabeceras Lista con los nombre de los campos (Columnas) que trae la consulta SQL.       
     * @return Hoja de calculo creada o null si no se pudo crear.
     */
    public Sheet crearHoja(String nombre, List<String> cabeceras){
        Sheet hoja = null;
        if(libro != null && nombre != null && !nombre.isEmpty() && !exist(nombre)){
            hoja = libro.createSheet(nombre);
            
            if(!cabeceras.isEmpty()){
                Row fila = hoja.createRow(0);        
                for(int i = 0; i < cabeceras.size(); i++) {
                    Cell celda = fila.createCell(i);
                    celda.setCellStyle(headerStyle); 
                    celda.setCellValue(cabeceras.get(i));
                }                
            }
            this.hojas.add(hoja);
            return hoja;
        }
        return hoja;
    }
    
    /**
     * 
     * @param nombre de la hoja de calculo a obtener.
     * @return Hoja de calculo o null si no existe una hoja con ese nombre.
     */
    public Sheet getHoja(String nombre){
    	if(libro != null){
           return libro.getSheet(nombre);
        }
        return null;
    }
    
    /**
     * 
     * @param index Posicion de la hoja a obtener mayor a 1.
     * @return Hoja de calculo o null si no existe una hoja con ese indice.
     */
    public Sheet getHoja(int index){
        if(libro != null){
           return libro.getSheetAt(index);
        }
        return null;        
    }
    
    /**
     * Dada una lista de datos estos se agregan a una fila de la hoja de excel dada.
     * 
     * @param hoja Hoja de excel a la cual se va a agregar la nueva fila.
     * @param pos indice de fila en el cual se escribiran los items.
     * @param items datos a agregar a la fila.
     */
    public void crearFila(Sheet hoja, int pos, List<String> items){
        if(hoja != null && items != null && !items.isEmpty()){
            Row fila = hoja.createRow(pos);
            for(int i = 0; i < items.size(); i++) {
                Cell celda = fila.createCell(i);
                celda.setCellValue(items.get(i));        
            }
        }
    }
    
    /**
     * 
     * @return Estilo de la cabecera de las hojas del archivo de excel.
     */
    public CellStyle getHeaderStyle(){
        return headerStyle;
    }
    
    /**
     * 
     * @param style Nuevo estilo que llevaran las columnas de las cabeceras del archivo de excel.
     */
    public void setHeaderStyle(CellStyle style){
        this.headerStyle = style;
    }

    /**
     * 
     * @return Lista de hojas del archivo.
     */
    public List<Sheet> getHojas() {
        return hojas;
    }

    /**
     * 
     * @param hojas Nueva lista de hojas del archivo de excel.
     */
    public void setHojas(List<Sheet> hojas) {
        this.hojas = hojas;
    }
    
    /**
     * Metodo el cual guarda los cambios realizados al archivo de excel.
     * 
     * @throws IOException Excepcion de manejo de archivos.
     */
    public void save() throws IOException{
        archivo.createNewFile();
    	FileOutputStream salida = new FileOutputStream(archivo);
        libro.write(salida);
        libro.close();            
    }

    @Override
    public String toString() {
        return "ArchivoExcel [ruta=" + ruta + ", archivo=" + archivo + ", libro=" + libro + ", headerStyle="+ headerStyle + ", hojas=" + hojas + "]";
    }

}//Fin clase ArchivoExcel.