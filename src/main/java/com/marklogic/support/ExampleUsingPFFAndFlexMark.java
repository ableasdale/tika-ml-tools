package com.marklogic.support;

import com.google.common.base.Strings;
import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.lang.invoke.MethodHandles;
import java.time.Instant;
import java.util.*;
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

    @SuppressWarnings("unchecked")
    public void processFolder(PSTFolder folder) throws IOException {
        depth++;
        // the root folder doesn't have a display name
        if (depth > 0) {
            printDepth();
            LOG.info("Folder: "+folder.getDisplayName());
        }

        // go through the folders...
        if (folder.hasSubfolders()) {
            Vector<PSTFolder> childFolders = null;
            try {
                childFolders = folder.getSubFolders();
            } catch (PSTException e) {
                LOG.error("PST Exception 1: ",e);
            }
            for (PSTFolder childFolder : childFolders) {
                processFolder(childFolder);
            }
        }

        // and now the emails for this folder
        if (folder.getContentCount() > 0) {
            depth++;
            PSTMessage email = null;
            email = getPstMessage(folder, email);
            while (email != null) {
                //printDepth();
                // TODO - pass in foldername and depth?
                LOG.info(email.getSubject());
                email = getPstMessage(folder, email);
//                String s = Optional.ofNullable(email.getBody()).orElse("");
//                if(Strings.isNullOrEmpty(email.getBody()))
//                    if (s.trim().isEmpty()) {
//                        LOG.info("no body");
//                    }
                try {
                    if(email != null) {
                        //LOG.info(email.toString());
                        Map<String, String> msg = readAMsg(email);
                    }


                }
                catch(ArrayIndexOutOfBoundsException e){
                    System.out.println("Warning: ArrayIndexOutOfBoundsException");
                    System.out.println("Warning: This message will not be reconciled. \n" +
                            "Check PST extraction process or Archiving process are correct");
                } catch (PSTException e) {
                    e.printStackTrace();
                }

                email = getPstMessage(folder, email);
            }
            depth--;
        }
        depth--;
    }

    private PSTMessage getPstMessage(PSTFolder folder, PSTMessage email) throws IOException {
        try {
            email = (PSTMessage) folder.getNextChild();
            PSTMessage email2 = email;
            es.submit(new EmailFileProcessor(email2));
        } catch (PSTException e) {
            LOG.error("PST Exception 2: ",e);
        }
        return email;
    }

    public void printDepth() {
        for (int x = 0; x < depth - 1; x++) {
            System.out.print(" | ");
        }
        System.out.print(" |- ");
    }
    private Map<String, String> readAMsg(PSTMessage email) throws PSTException, IOException {
        HashMap<String, String> msg = new HashMap<>();
        //LOG.info(email.toString());
        //LOG.info(email.getOriginalSubject());
        //LOG.info("subj"+email.getSubject());
        msg.put("Subject", email.getSubject());
        //LOG.info(email.getBody());
//                this.printDepth();
        if (email.hasAttachments()) {
//                    for (int att = 0; att < email.getNumberOfAttachments();
//                    att++){
//                        System.out.println("Attachment number: " + att + 1);
//                        PSTAttachment attachment = email.getAttachment(att);
//                        this.printDepth();
//                        System.out.println(
//                                attachment.getLongFilename() + '-' +
//                                attachment.getSize() + '-' + attachment
//                                .getFilesize() + '-' +
//                        attachment.getAttachmentContentDisposition());
//                    }
            msg.put("Attachments", Integer.toString(email
                    .getNumberOfAttachments()));
        } else {
            msg.put("Attachments", "0");
        }

        msg.put("Sender", email.getSentRepresentingEmailAddress
                ());
//                this.printDepth();
        msg.put("From", email.getSentRepresentingName());
        msg.put("RcvdRepresentingEmailAddress",
                email.getRcvdRepresentingEmailAddress());
//                this.printDepth();
        msg.put("To", email.getDisplayTo());
//                System.out.println("To: " + email.getDisplayTo());

//                this.printDepth();
        msg.put("CC", email.getDisplayCC());
//                System.out.println("CC: " + email.getDisplayCC());


        msg.put("BCC", email.getDisplayBCC());
        String body = email.getBody();
        msg.put("Contents", body);
//                this.printDepth();
//                System.out.println("Contents: " + body);
//                this.printDepth();
        msg.put("NoOfRecipients", Integer.toString(email
                .getNumberOfRecipients()));
//                this.printDepth();
        // getClientSubmitTime converts date to local timezone
        // We want UTC Time
        msg.put("ClientSubmitUTCTime", dateToUTC(email
                .getClientSubmitTime()).toString());
//                this.printDepth();

        msg.put("ClientSubmitLocalTime", email
                .getClientSubmitTime().toString());

        msg.put("MessageDeliveryTime", email
                .getMessageDeliveryTime().toString());
//                msg.put("Conversation",
//                        this.mapConversation(body).toString());
//                this.printDepth();
        return msg;
    }
    private Instant dateToUTC(Date date){
        return new Date(date.getTime()).toInstant();
    }
}
