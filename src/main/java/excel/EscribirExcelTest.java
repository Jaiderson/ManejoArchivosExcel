package excel;

import java.io.IOException;
import java.util.List;

import com.google.common.collect.Lists;

import util.GenerarExcel;
import util.TablaExcel;

public class EscribirExcelTest {

	public static void main(String[] args) {
		GenerarExcel genExcel = new GenerarExcel("C:\\Tempo\\Excel\\Test 01.xlsx");
		try {
			genExcel.generarExcel(datosHoja1(), true,  "Directorio 1");
			genExcel.generarExcel(datosHoja2(), false, "Directorio 2");
			genExcel.generarExcel(datosHoja1(), true,  "Directorio 3");
			genExcel.generarExcel(datosHoja2(), false, "Directorio 4");
			genExcel.guardarExcel();
		} catch (IOException e) {
			System.err.println("Error creando archivo de excel --> "+e.getMessage());
		}

	}
	
	private static List<TablaExcel> datosHoja1(){
		List<TablaExcel> hoja1 = Lists.newArrayList();
		TablaExcel cabecera = new TablaExcel();
		cabecera.addCampo("DNI");
		cabecera.addCampo("NOMBRES");
		cabecera.addCampo("APELLIDOS");
		cabecera.addCampo("CELULAR");
		hoja1.add(cabecera);
		
		TablaExcel item = null;
		for(int x=0; x < 50; x++) {
			item = new TablaExcel();
			item.addCampo("123456"+x);
			item.addCampo("NOMBRE "+x);
			item.addCampo("APELLIDOS "+x);
			item.addCampo("31088554"+x);
			hoja1.add(item);			
		}

		return hoja1;
	}

	private static List<TablaExcel> datosHoja2(){
		List<TablaExcel> hoja2 = Lists.newArrayList();
		TablaExcel cabecera = new TablaExcel();
		cabecera.addCampo("DNI 2");
		cabecera.addCampo("NOMBRES 2");
		cabecera.addCampo("APELLIDOS 2");
		cabecera.addCampo("CELULAR 2");
		hoja2.add(cabecera);
		
		TablaExcel item = null;
		for(int x=0; x < 50; x++) {
			item = new TablaExcel();
			item.addCampo("987654"+x);
			item.addCampo("NOMBRE 2 "+x);
			item.addCampo("APELLIDOS 2 "+x);
			item.addCampo("31588550"+x);
			hoja2.add(item);			
		}

		return hoja2;
	}

}
