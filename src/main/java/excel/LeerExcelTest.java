package excel;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import util.ArchivoLeerExcel;
import util.TablaExcel;

public class LeerExcelTest {

	public static void main(String[] args) {
       ArchivoLeerExcel archivo = new ArchivoLeerExcel("C:\\Tempo\\Excel\\Test 01.xlsx");
       try {
           archivo.abrirArchivo();
           for(int x=0; x < archivo.getHojas().size(); x++) {
        	   imprimir(archivo.getHoja(x), archivo.getHojas().get(x).getSheetName());
           }
           
       }
       catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
              System.err.println("Error abriendo archivo de excel  --> "+e.getMessage());
       }

    }
	
	public static void imprimir(List<TablaExcel> matriz, String nombreHoja) {
		System.err.println("Imprimiento contenido de la hoja: "+nombreHoja);
		System.out.println();
		
		for(TablaExcel fila : matriz) {
			String line = "";
			for(String dato : fila.getCampos()) {
				line += dato +" - ";
			}
			System.out.println(line.substring(0, line.length()-3));
		}
		System.out.println();
		System.out.println();
	}

}
