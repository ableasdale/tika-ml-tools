package com.marklogic.support;

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
        LOG.debug("Creating the MarkLogic ContentSourceFactory provider");
        String host = Util.getConfiguration().getString("host");
        String username = Util.getConfiguration().getString("username");
        String password = Util.getConfiguration().getString("password");
        int port = Util.getConfiguration().getInt("port");
        String contentbase = Util.getConfiguration().getString("contentbase");


        try {
            URI uri = new URI(generateXdbcConnectionUri(host, username, password, port, contentbase));
            cs = ContentSourceFactory
                    .newContentSource(uri);
        } catch (URISyntaxException e) {
            LOG.error("URISyntaxException encountered: ",e);
        } catch (XccConfigException e) {
            LOG.error("XccConfigException encountered: ",e);
        }
    }

    private String generateXdbcConnectionUri(String hostname, String username, String password, int port, String contentbase) {
        StringBuilder sb = new StringBuilder();
        sb.append("xdbc://").append(username).append(":")
                .append(password).append("@")
                .append(hostname).append(":")
                .append(port).append("/").append(contentbase);
        LOG.debug(String.format("XCC Connection: %s", sb.toString()));
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