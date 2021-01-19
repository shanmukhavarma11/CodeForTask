import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.LogManager;

public class Log4jexample {
	private static org.apache.log4j.Logger log=LogManager.getLogger(Log4jexample.class);
public static void main(String args[]) {
	BasicConfigurator.configure();
	log.info("thank");
	
}

}
