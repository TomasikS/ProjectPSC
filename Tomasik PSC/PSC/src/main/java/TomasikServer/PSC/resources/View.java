package TomasikServer.PSC.resources;

import java.util.List;

import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.PathParam;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;

import TomasikServer.PSC.db.Database;
 

@Path("/posta")

public class View {
	@GET
    @Path("/city/{city}")
    @Produces(MediaType.APPLICATION_JSON)
    public String getPSC(@PathParam("city")String city){

        List<String> ZipList = new Database().getPSC(city);

        boolean b = false;
        String result = "[";
        for (String temp : ZipList) {
            if (b == true) {
                result += ',';
            } else
                b = true;
            result += "\"" + temp + "\"";
        }
        result += "]";
        return  result;
    }

    @GET
    @Path("/psc/{psc}")
    @Produces(MediaType.APPLICATION_JSON)
    public String getCity(@PathParam("psc")String p){
        List<String> CityList = new Database().getCity(p);

        boolean b = false;
        String result = "[";
        for (String temp : CityList) {
            if (b == true) {
                result += ',';
            } else
                b = true;
            result += "\"" + temp + "\"";
        }
        result += "]";
        return  result;
    }

}
