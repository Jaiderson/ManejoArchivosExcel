package util;

import com.google.common.collect.Lists;
import java.util.List;
/**
 * Clase la cual representa el contenido de la infrmacion de una hoja de excel. 
 * Es decir reresenta la matriz de datos de una hoja de excel.
 *
 * @author Jaider Adriam Serrano Sepulveda.
 */
public class TablaExcel {
	
    private List<String> campos;

    public TablaExcel(){
        campos = Lists.newArrayList();
    }
    /**
     * 
     * @param valor Nuevo valor a agregar en la lista de campos.
     */
    public void addCampo(String valor){
        campos.add(valor);
    }
    
    public void addCampos(List<String> campos){
        if(campos == null || campos.isEmpty()){
            return;
        }
        campos.addAll(campos);
    }
    /**
     * 
     * @return Lista con los valores de los campos.
     */
    public List<String> getCampos() {
        return campos;
    }
    /**
     * 
     * @param campos Nueva lista de los valores de los campos.
     */
    public void setCampos(List<String> campos) {
        this.campos = campos;
    }

    public String getContenido(String separador){
        String result = "";
        if(campos.isEmpty()){
            return result;
        }
        
        result = campos.stream().map((campo) -> campo + separador).reduce(result, String::concat);
        return result;
    }
    
    @Override
    public String toString() {
        return "TableWraper [campos=" + campos + "]";
    }
    
    public String getValores(){        
        if(campos.isEmpty()){
            return "";
        }
        String result = "[";
        for(int x=0; x<campos.size(); x++){
            result += " "+campos.get(x)+" - ";
        }
        return result.substring(0, result.length() - 3)+" ]";
    }
	
}//Fin clase TableWraper.
