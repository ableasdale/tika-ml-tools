package com.marklogic.support;

import com.google.common.base.CharMatcher;
import com.marklogic.xcc.Content;
import com.marklogic.xcc.ContentFactory;
import com.marklogic.xcc.Session;
import com.marklogic.xcc.exceptions.RequestException;
import com.pff.PSTException;
import com.pff.PSTMessage;
import com.vladsch.flexmark.ast.Node;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.options.MutableDataSet;
import nu.xom.Document;
import nu.xom.Element;
import org.apache.tika.metadata.*;
import org.apache.tika.parser.microsoft.OutlookExtractor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.invoke.MethodHandles;

import static java.lang.String.valueOf;
import static java.nio.charset.StandardCharsets.UTF_8;

public class EmailFileProcessor implements Runnable {

    private static Logger LOG = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass().getCanonicalName());
    private static String DC_NS = "http://purl.org/dc/elements/1.1/";
    private static String MS_MAPI_NS = "http://schemas.microsoft.com/mapi";
    private static String MSG_NS = "URN:IANA:message:rfc822:";

    private PSTMessage email;
    private String body;
    private String folderName;

    EmailFileProcessor(PSTMessage email, String body, String folderName) {
        this.email = email;
        this.body = body;
        this.folderName = folderName;
    }

    @Override
    public void run() {
        processEmail(email, body, folderName);
    }

    private static void processEmail(PSTMessage pstMail, String body, String folderName) {
        Element root = new Element("Email");
        root.addNamespaceDeclaration("dc", DC_NS);
        root.addNamespaceDeclaration("meta", MS_MAPI_NS);
        root.addNamespaceDeclaration("message", MSG_NS);
        root.addNamespaceDeclaration("xhtml", "http://www.w3.org/1999/xhtml");

        Element meta = new Element("Metadata");
        addNamespacedElement(meta, TikaCoreProperties.IDENTIFIER, pstMail.getInternetMessageId(), DC_NS);
        addNamespacedElement(meta, TikaCoreProperties.TITLE, pstMail.getSubject(), DC_NS);
        addElement(meta, Metadata.MESSAGE_FROM, pstMail.getSenderName());
        addNamespacedElement(meta, TikaCoreProperties.CREATOR, pstMail.getSentRepresentingName(), DC_NS);
        addNamespacedElement(meta, TikaCoreProperties.CREATED, pstMail.getCreationTime().toString(), DC_NS);
        addNamespacedElement(meta, TikaCoreProperties.MODIFIED, pstMail.getLastModificationTime().toString(), DC_NS);
        addNamespacedElement(meta, TikaCoreProperties.COMMENTS, pstMail.getComment(), DC_NS);
        addElement(meta, Metadata.MESSAGE_FROM, pstMail.getSenderName());
        addElement(meta, "folder", folderName);
        addElement(meta, "descriptorNodeId", valueOf(pstMail.getDescriptorNodeId()));
        addElement(meta, "senderEmailAddress", pstMail.getSenderEmailAddress());
        addElement(meta, "recipients", pstMail.getRecipientsString());
        addElement(meta, "displayTo", pstMail.getDisplayTo());
        addElement(meta, "displayCC", pstMail.getDisplayCC());
        addElement(meta, "displayBCC", pstMail.getDisplayBCC());
        addElement(meta, "importance", valueOf(pstMail.getImportance()));
        addElement(meta, "priority", valueOf(pstMail.getPriority()));
        addElement(meta, "flagged", valueOf(pstMail.isFlagged()));
        addNamespacedElement(meta, Office.MAPI_MESSAGE_CLASS, OutlookExtractor.getMessageClass(pstMail.getMessageClass()), MS_MAPI_NS);
        addNamespacedElement(meta, Message.MESSAGE_FROM_EMAIL, pstMail.getSenderEmailAddress(), MSG_NS);
        addNamespacedElement(meta, Office.MAPI_FROM_REPRESENTING_EMAIL, pstMail.getSentRepresentingEmailAddress(), MS_MAPI_NS);
        addNamespacedElement(meta, Message.MESSAGE_FROM_NAME, pstMail.getSenderName(), MSG_NS);
        addNamespacedElement(meta, Office.MAPI_FROM_REPRESENTING_NAME, pstMail.getSentRepresentingName(), MS_MAPI_NS);

/* TODO - is this killing us?
        MutableDataSet options = new MutableDataSet();
        com.vladsch.flexmark.parser.Parser parser = Parser.builder(options).build();
        Node document = parser.parse(pstMail.getBody());
        HtmlRenderer renderer = HtmlRenderer.builder(options).build();


        Element HTMLbody = new Element("HTMLBody");
        HTMLbody.appendChild(CharMatcher.JAVA_ISO_CONTROL.removeFrom(renderer.render(document))); */


        try {
            //String s = pstMail.getBodyHTML();
            // String t = pstMail.getBodyPrefix();
            // LOG.info("type:"+pstMail.getNativeBodyType());
            // LOG.info("size "+pstMail.getMessageSize());
            // LOG.info("class "+pstMail.getMessageClass());
            // LOG.info(t);
            //String u = pstMail.getBody();
            //LOG.info("body?"+u.length());
            //LOG.info(pstMail.getBody());
            /*if(pstMail.getBody() != null) {
                byte[] mailContent = pstMail.getBody().getBytes();
            }*/
        } catch (Exception e) {
            LOG.info("No mail content", e);
        }

        /* TODO? Add email body as Base64 encoded data

        Element HTMLbody = null;
        addElement(meta, "HTMLContentMD5", DigestUtils.md5Hex(HTMLMessage));
        */

        root.appendChild(meta);
        Element ebody = new Element("Body");
        ebody.appendChild(CharMatcher.JAVA_ISO_CONTROL.removeFrom(body));
        root.appendChild(ebody);
        //root.appendChild(HTMLbody);

        Document doc = new Document(root);

        Session s = MarkLogicContentSourceProvider.getInstance().getContentSource().newSession();
        String docUri = createDocUriFromId(pstMail.getInternetMessageId());
        Content c = ContentFactory.newContent(docUri, doc.toXML(), null);

        try {
            s.insertContent(c);
        } catch (RequestException e) {
            LOG.error("MarkLogic Request Exception encountered", e);
        }
        s.close();


        /* TODO - adding this would be quite a bit of extra work:
        //add recipient details
        try {
            for (int i = 0; i < pstMail.getNumberOfRecipients(); i++) {
                PSTRecipient recipient = pstMail.getRecipient(i);
                switch (OutlookExtractor.RECIPIENT_TYPE.getTypeFromVal(recipient.getRecipientType())) {
                    case TO:
                        OutlookExtractor.addEvenIfNull(Message.MESSAGE_TO_DISPLAY_NAME,
                                recipient.getDisplayName(), mailMetadata);
                        OutlookExtractor.addEvenIfNull(Message.MESSAGE_TO_EMAIL,
                                recipient.getEmailAddress(), mailMetadata);
                        break;
                    case CC:
                        OutlookExtractor.addEvenIfNull(Message.MESSAGE_CC_DISPLAY_NAME,
                                recipient.getDisplayName(), mailMetadata);
                        OutlookExtractor.addEvenIfNull(Message.MESSAGE_CC_EMAIL,
                                recipient.getEmailAddress(), mailMetadata);
                        break;
                    case BCC:
                        OutlookExtractor.addEvenIfNull(Message.MESSAGE_BCC_DISPLAY_NAME,
                                recipient.getDisplayName(), mailMetadata);
                        OutlookExtractor.addEvenIfNull(Message.MESSAGE_BCC_EMAIL,
                                recipient.getEmailAddress(), mailMetadata);
                        break;
                    default:
                        //do we want to handle unspecified or unknown?
                        break;
                }
            }
        } catch (PSTException e) {
            LOG.error("PSTException",e);
        } */
        //we may want to experiment with working with the bodyHTML.
        //However, because we can't get the raw bytes, we _could_ wind up sending
        //a UTF-8 byte representation of the html that has a conflicting metaheader
        //that causes the HTMLParser to get the encoding wrong.  Better if we could get
        //the underlying bytes from the pstMail object...


        // System.out.println(pstMail.getBody());
        //mailMetadata.set(TikaCoreProperties.CONTENT_TYPE_OVERRIDE, MediaType.TEXT_PLAIN.toString());
        //embeddedExtractor.parseEmbedded(new ByteArrayInputStream(mailContent), handler, mailMetadata, true);
    }

    private static String createDocUriFromId(String id) {
        id = id.replace("<", "");
        id = id.replace(">", "");
        return "/" + CharMatcher.JAVA_ISO_CONTROL.removeFrom(id) + ".xml";
    }

    private static void addElement(Element root, String elementName, String content) {
        Element e = new Element(elementName);
        e.appendChild(CharMatcher.JAVA_ISO_CONTROL.removeFrom(content));
        root.appendChild(e);
    }

    private static void addNamespacedElement(Element root, Property elementName, String content, String namespaceUri) {
        Element e = new Element(elementName.getName(), namespaceUri);
        e.appendChild(CharMatcher.JAVA_ISO_CONTROL.removeFrom(content));
        root.appendChild(e);
    }




/*
    private static void processPSTFile(Path path) {
        LOG.info(String.format("Processing PST File: %s", path));
        TikaInputStream in = null;
        try {
            FileInputStream fis = new FileInputStream(path.toFile());
            in = TikaInputStream.get(fis);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Metadata metadata = new Metadata();
        ContentHandler handler = new ToXMLContentHandler();
        ParseContext context = new ParseContext();
        //XMLContent
        XHTMLContentHandler xml = new XHTMLContentHandler(handler, metadata);
        EmbeddedDocumentExtractor embeddedExtractor = EmbeddedDocumentUtil.getEmbeddedDocumentExtractor(context);
        metadata.set(Metadata.CONTENT_TYPE, OutlookPSTParser.MS_OUTLOOK_PST_MIMETYPE.toString());

        PSTFile pstFile = null;
        try {
            pstFile = new PSTFile(in.getFile().getPath());
            metadata.set(Metadata.CONTENT_LENGTH, valueOf(pstFile.getFileHandle().length()));
            boolean isValid = pstFile.getFileHandle().getFD().valid();
            metadata.set("isValid", valueOf(isValid));
            if (isValid) {
                //Vector<PSTFolder> folders = pstFile.getRootFolder().getSubFolders();
                parseFolder(xml, pstFile.getRootFolder(), embeddedExtractor);
            }
        } catch (Exception e) {
            LOG.error("Exception Caught: ", e);
        } finally {
            if (pstFile != null && pstFile.getFileHandle() != null) {
                try {
                    pstFile.getFileHandle().close();
                    //CloseUtils.close(fileStream);
                    in.close();
                    in = null;
                } catch (IOException e) {
                    LOG.error("IO Exception", e);
                }
            }
        }
    } */


    /*
    private static void parseFolder(XHTMLContentHandler handler, PSTFolder pstFolder, EmbeddedDocumentExtractor embeddedExtractor)
            throws Exception {
        if (pstFolder.getContentCount() > 0) {
            PSTMessage pstMail = (PSTMessage) pstFolder.getNextChild();
            while (pstMail != null) {
                final Metadata mailMetadata = new Metadata();
                //parse attachments first so that stream exceptions
                //in attachments can make it into mailMetadata.
                //RecursiveParserWrapper copies the metadata and thereby prevents
                //modifications to mailMetadata from making it into the
                //metadata objects cached by the RecursiveParserWrapper
                //parseMailAttachments(handler, pstMail, mailMetadata, embeddedExtractor);
                //es.submit(new com.marklogic.support.develop.PSTFileProcessor(pstMail));
                parserMailItem(handler, pstMail, mailMetadata, embeddedExtractor);
                pstMail = (PSTMessage) pstFolder.getNextChild();
            }
        }
        if (pstFolder.hasSubfolders()) {
            for (PSTFolder pstSubFolder : pstFolder.getSubFolders()) {
                parseFolder(handler, pstSubFolder, embeddedExtractor);
            }
        }
    }*/

}
