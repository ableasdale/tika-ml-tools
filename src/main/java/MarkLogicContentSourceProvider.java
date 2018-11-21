import java.lang.invoke.MethodHandles;
import java.net.URI;
import java.net.URISyntaxException;

import com.marklogic.xcc.ContentSource;
import com.marklogic.xcc.ContentSourceFactory;
import com.marklogic.xcc.exceptions.XccConfigException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * Base Provider Class for a MarkLogic Content Source Factory - a singleton
 * provider which should only be instantiated once; each individual connection
 * to the XML Server should use this ContentSourceFactory
 *
 * As this is a singleton, use getInstance() to access the class
 *
 * @author Alex Bleasdale
 *
 */
public class MarkLogicContentSourceProvider {

    private static Logger LOG = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass());

    private ContentSource cs;

    private MarkLogicContentSourceProvider() {
        LOG.info("Creating the MarkLogic ContentSourceFactory provider");
        String[] hosts = Util.getConfiguration().getStringArray("host");
        try {
            URI uri = new URI(generateXdbcConnectionUri(hosts[0]));
            cs = ContentSourceFactory
                    .newContentSource(uri);
        } catch (URISyntaxException e) {
            LOG.error("URISyntaxException encountered: ",e);
        } catch (XccConfigException e) {
            LOG.error("XccConfigException encountered: ",e);
        }
    }

    // TODO - config file for other values!
    private String generateXdbcConnectionUri(String hostname) {
        StringBuilder sb = new StringBuilder();
        sb.append("xdbc://").append("q").append(":")
                .append("q").append("@")
                .append(hostname).append(":")
                .append("8000");
        //LOG.info("Conn: " + sb.toString());
        return sb.toString();
    }

    private static class MarkLogicContentSourceProviderHolder {
        private static final MarkLogicContentSourceProvider INSTANCE = new MarkLogicContentSourceProvider();
    }

    public static MarkLogicContentSourceProvider getInstance() {
        return MarkLogicContentSourceProviderHolder.INSTANCE;
    }

    public ContentSource getContentSource() {
        return cs;
    }

}