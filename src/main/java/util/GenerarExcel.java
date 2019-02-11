package util;

import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;

public class GenerarExcel {
	
    private ArchivoEscribirExcel libro;
    /**
     * Clase para generar el resultado de un sentencia SQL a un archivo de excel (.xlsx).
     * 
     * @author Jaider Adriam Serrano Sepulveda.
     * @param ruta Ruta donde se creara el archivo de excel. Ejm. C:\Tempo\Resultado_SQL.xlsx
     */
    public GenerarExcel(String ruta){
        this.libro = new ArchivoEscribirExcel(ruta);
        this.libro.crear();
    }
    /**
     * 
     * @param matriz Matriz con la informacion a escribir en una hoja de excel.
     * @param conCabecera True si la lista de resultados trae los nombres de las conlumnas resultado de la consulta SQL.
     * @param nomHoja Nombre que se le dara a la nueva hoja creada, este nombre no lo debe tener otra hoja existente.
     * @throws IOException Exepcion de manejo de archivos.
     */
    public void generarExcel(List<TablaExcel> matriz, boolean conCabecera, String nomHoja) throws IOException{
        Sheet hoja;
        int numFila = 0;

        if(conCabecera){
                hoja = libro.crearHoja(nomHoja, matriz.get(0).getCampos());
                matriz.remove(0);
                numFila++;
        }
        else{
                hoja = libro.crearHoja(nomHoja);			
        }

        for(TablaExcel valores : matriz)
            libro.crearFila(hoja, numFila++, valores.getCampos());
    }
    
    public void guardarExcel() throws IOException {
    	libro.save();
    }

//**********************   GETTERS ANS SETTERS   **********************\\	
    /**
     * 
     * @return Libro de excel.
     */
    public ArchivoEscribirExcel getLibro() {
        return libro;
    }

    /**
     * 
     * @param libro Nuevo libro de excel.
     */
    public void setLibro(ArchivoEscribirExcel libro) {
        this.libro = libro;
    }

    @Override
    public String toString() {
        return "GenerarSQLExcel [libro=" + libro + " ]";
    }

}//Fin clase GenerarSQLExcel.
