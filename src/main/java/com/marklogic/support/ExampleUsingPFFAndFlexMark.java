package com.marklogic.support;

import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
import com.vladsch.flexmark.ast.Node;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.options.MutableDataSet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.lang.invoke.MethodHandles;
import java.util.Vector;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

public class ExampleUsingPFFAndFlexMark {

    private static Logger LOG = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass().getCanonicalName());

    private static ExecutorService es =
            new ThreadPoolExecutor(16, 16, 0L, TimeUnit.MILLISECONDS,
                    new ArrayBlockingQueue<Runnable>(999999));

    public static void main(String[] args) {
        new ExampleUsingPFFAndFlexMark(Util.getConfiguration().getString("pstfile"));

        // Stop the thread pool
        es.shutdown();
        // Drain the queue
        while (!es.isTerminated()) {
            try {
                es.awaitTermination(72, TimeUnit.HOURS);
            } catch (InterruptedException e) {
                LOG.error("Exception caught: ", e);
            }
        }
    }

    public ExampleUsingPFFAndFlexMark(String filename) {
        try {
            PSTFile pstFile = new PSTFile(filename);
            LOG.info(pstFile.getMessageStore().getDisplayName());
            processFolder(pstFile.getRootFolder());
        } catch (Exception e) {
            LOG.error("Exception encountered",e);
        }
    }
    int depth = -1;

    public void processFolder(PSTFolder folder) throws PSTException, IOException {
        depth++;
        // the root folder doesn't have a display name
        if (depth > 0) {
            printDepth();
            LOG.info("Folder: "+folder.getDisplayName());
        }

        // go through the folders...
        if (folder.hasSubfolders()) {
            Vector<PSTFolder> childFolders = folder.getSubFolders();
            for (PSTFolder childFolder : childFolders) {
                processFolder(childFolder);
            }
        }

        // and now the emails for this folder
        if (folder.getContentCount() > 0) {
            depth++;
            PSTMessage email = (PSTMessage) folder.getNextChild();
            while (email != null) {
                printDepth();
                es.submit(new EmailFileProcessor(email));
                //LOG.info("Subj: "+email.getSubject());
                //System.out.println(email.getBody());
                MutableDataSet options = new MutableDataSet();
                Parser parser = Parser.builder(options).build();
                Node document = parser.parse(email.getBody());
                HtmlRenderer renderer = HtmlRenderer.builder(options).build();
                //LOG.info(renderer.render(document));

                //Node document = parser.parse(email.getBody());
                //String html = renderer.render(document);  // "<p>This is <em>Sparta</em></p>\n"
                //System.out.println(html);

                //System.out.println(email.getBodyHTML());
                email = (PSTMessage) folder.getNextChild();
            }
            depth--;
        }
        depth--;
    }

    public void printDepth() {
        for (int x = 0; x < depth - 1; x++) {
            System.out.print(" | ");
        }
        System.out.print(" |- ");
    }
}
