package TomasikServer.PSC;

 
import java.util.EnumSet;

import javax.servlet.DispatcherType;
import javax.servlet.FilterRegistration;

import org.eclipse.jetty.servlets.CrossOriginFilter;

 
import TomasikServer.PSC.resources.View;
import io.dropwizard.Application;
import io.dropwizard.setup.Bootstrap;
import io.dropwizard.setup.Environment;

public class ServerApplication extends Application<ServerConfiguration> {

	public static void main(final String[] args) throws Exception {
		
		String applicationPort = "9090";
        String adminPort = "9091";

        System.setProperty("dw.server.applicationConnectors[0].port", applicationPort);
        System.setProperty("dw.server.adminConnectors[0].port", adminPort);
        new ServerApplication().run(args);
    }


    @Override
    public String getName() {
        return "ProjectPSCobci";
    }

    @Override
    public void initialize(final Bootstrap<ServerConfiguration> bootstrap) {
        // TODO: application initialization
    }

    @Override
    public void run(final ServerConfiguration configuration,
                    final Environment environment) {

        environment.jersey().register(
                new View()
        );
        
   
        
        
        
      
        final FilterRegistration.Dynamic cors =
                environment.servlets().addFilter("CORS", CrossOriginFilter.class);

// Configure CORS parameters
        cors.setInitParameter("allowedOrigins", "*");
        cors.setInitParameter("allowedHeaders", "X-Requested-With,Content-Type,Accept,Origin");
        cors.setInitParameter("allowedMethods", "OPTIONS,GET,PUT,POST,DELETE,HEAD");

// Add URL mapping
        cors.addMappingForUrlPatterns(EnumSet.allOf(DispatcherType.class), true, "/*");
    }
}
