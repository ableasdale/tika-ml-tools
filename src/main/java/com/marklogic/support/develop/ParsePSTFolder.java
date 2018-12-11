package com.marklogic.support.develop;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.invoke.MethodHandles;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.concurrent.*;


public class ParsePSTFolder {


    private static Logger LOG = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass().getCanonicalName());
    private static ExecutorService es =
            new ThreadPoolExecutor(16, 16, 0L, TimeUnit.MILLISECONDS,
                    new ArrayBlockingQueue<Runnable>(999999));

            //Executors.newFixedThreadPool(5);

    public static void main(String[] args) throws Exception {
        //cs = ContentSourceFactory.newContentSource(new URI("xcc://q:q@localhost:8000/Emails"));
        // Walk a top level directory
        Files.walk(Paths.get("directory\\here"))
                .filter(Files::isRegularFile)
                .forEach(path -> es.submit(new PSTFileProcessor(path)));

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



    /* TODO - adding attachments would be quite a bit of extra work
    private static void parseMailAttachments(XHTMLContentHandler xhtml, PSTMessage email,
                                      final Metadata mailMetadata,
                                      EmbeddedDocumentExtractor embeddedExtractor)
            throws TikaException {
        int numberOfAttachments = email.getNumberOfAttachments();
        for (int i = 0; i < numberOfAttachments; i++) {
            try {
                PSTAttachment attach = email.getAttachment(i);

                // Get the filename; both long and short filenames can be used for attachments
                String filename = attach.getLongFilename();
                if (filename.isEmpty()) {
                    filename = attach.getFilename();
                }

                xhtml.element("p", filename);

                Metadata attachMeta = new Metadata();
                //attachMeta.set(TikaCoreProperties.RESOURCE_NAME_KEY, filename);
                //attachMeta.set(TikaCoreProperties.EMBEDDED_RELATIONSHIP_ID, filename);
                AttributesImpl attributes = new AttributesImpl();
                attributes.addAttribute("", "class", "class", "CDATA", "embedded");
                attributes.addAttribute("", "id", "id", "CDATA", filename);
                xhtml.startElement("div", attributes);
                if (embeddedExtractor.shouldParseEmbedded(attachMeta)) {
                    TikaInputStream tis = null;
                    try {
                        tis = TikaInputStream.get(attach.getFileInputStream());
                    } catch (NullPointerException e) {//TIKA-2488
                        EmbeddedDocumentUtil.recordEmbeddedStreamException(e, mailMetadata);
                        continue;
                    }

                    try {
                        embeddedExtractor.parseEmbedded(tis, xhtml, attachMeta, true);
                    } finally {

                        tis.close();
                    }
                }
                xhtml.endElement("div");

            } catch (Exception e) {
                throw new TikaException("Unable to unpack document stream", e);
            }
        }
    }


    private static AttributesImpl createAttribute(String attName, String attValue) {
        AttributesImpl attributes = new AttributesImpl();
        attributes.addAttribute("", attName, attName, "CDATA", attValue);
        return attributes;
    }


    public Set<MediaType> getSupportedTypes(ParseContext context) {
        return SUPPORTED_TYPES;
    }*/
}
