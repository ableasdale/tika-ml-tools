package com.marklogic.support;

import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
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
            System.out.println(pstFile.getMessageStore().getDisplayName());
            processFolder(pstFile.getRootFolder());
        } catch (Exception e) {
            LOG.error("Exception encountered", e);
        }
    }

    int depth = -1;


    public void processFolder(PSTFolder folder) throws IOException {
        depth++;
        // the root folder doesn't have a display name
        if (depth > 0) {
            printDepth();
            System.out.println(String.format("Processing Folder: %s (%d)", folder.getDisplayName(), folder.getContentCount()));
        }

        // go through the folders...
        if (folder.hasSubfolders()) {
            Vector<PSTFolder> childFolders = null;
            try {
                childFolders = folder.getSubFolders();
            } catch (PSTException e) {
                LOG.error("PST Exception: ", e);
            }
            for (PSTFolder childFolder : childFolders) {
                processFolder(childFolder);
            }
        }

        // and now the emails for this folder
        if (folder.getContentCount() > 0) {
            depth++;
            PSTMessage email = null;
            email = processPstMessage(folder, email);
            while (email != null) {
                // TODO - pass in depth / full folder heirarchy / path?
                email = processPstMessage(folder, email);
            }
            depth--;
        }
        depth--;
    }

    private PSTMessage processPstMessage(PSTFolder folder, PSTMessage email) throws IOException {
        try {
            email = (PSTMessage) folder.getNextChild();
            if (email != null) {
                String body = email.getBody();
                es.submit(new EmailFileProcessor(email, body, folder.getDisplayName()));
            }
        } catch (PSTException e) {
            LOG.error("PST Exception: ", e);
        }
        return email;
    }

    public void printDepth() {
        for (int x = 0; x < depth - 1; x++) {
            System.out.print(" | ");
        }
        System.out.print(" |- ");
    }
}
