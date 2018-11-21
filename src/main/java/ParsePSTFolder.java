import com.google.common.base.CharMatcher;
import com.marklogic.xcc.*;
import com.marklogic.xcc.exceptions.RequestException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
import nu.xom.Document;
import nu.xom.Element;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.tika.exception.TikaException;
import org.apache.tika.extractor.EmbeddedDocumentExtractor;
import org.apache.tika.extractor.EmbeddedDocumentUtil;
import org.apache.tika.io.TikaInputStream;
import org.apache.tika.metadata.*;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.mbox.OutlookPSTParser;
import org.apache.tika.parser.microsoft.OutlookExtractor;
import org.apache.tika.sax.ToXMLContentHandler;
import org.apache.tika.sax.XHTMLContentHandler;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.invoke.MethodHandles;
import java.lang.reflect.Field;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.UUID;

import static java.lang.String.valueOf;
import static java.nio.charset.StandardCharsets.UTF_8;

public class ParsePSTFolder {

    private static String DC_NS = "http://purl.org/dc/elements/1.1/";
    private static String MS_MAPI_NS = "http://schemas.microsoft.com/mapi";
    private static String MSG_NS = "URN:IANA:message:rfc822:";
    private static ContentSource cs;
    private static Logger LOG = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass().getCanonicalName());

    public static void main(String[] args) throws Exception {
        cs = ContentSourceFactory.newContentSource(new URI("xcc://q:q@localhost:8000/Emails"));
        // Walk a directory
        Files.walk(Paths.get("E:\\\\RevisedEDRMv1_Complete"))
                .filter(Files::isRegularFile)
                .forEach(path -> processPSTFile(path));
    }

    private static void processPSTFile(Path path) {
        LOG.info("Processing PST File: "+path);
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
            try {
                throw new TikaException(e.getMessage(), e);
            } catch (TikaException e1) {
                e1.printStackTrace();
            }
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
    }

    /* I don't think this is necessary
    public static void cleanUp(PSTFile docIn){
        try {
            //The OPCPackage object always refers the temporary
            //file created by the Apache POI
            OPCPackage pkg = docIn.getPackage();

            //Apache POI is not providing any getter method for
            //"oroginalPackagePath" where
            //it'll store the location of the temporary file
            //that is generated. So, Using reflection to get
            // the temporary file name along with its path
            Field f = OPCPackage.class.getDeclaredField("originalPackagePath");
            f.setAccessible(true);
            String path = (String)f.get(pkg);

            //Delete the file
            File file = new File(path);
            file.delete();

        } catch (Exception e) {
            System.out.println("Exception occurred while cleanup");
        }
    } */


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
                parserMailItem(handler, pstMail, mailMetadata, embeddedExtractor);
                pstMail = (PSTMessage) pstFolder.getNextChild();
            }
        }
        if (pstFolder.hasSubfolders()) {
            for (PSTFolder pstSubFolder : pstFolder.getSubFolders()) {
                parseFolder(handler, pstSubFolder, embeddedExtractor);
            }
        }
    }

    private static void parserMailItem(XHTMLContentHandler handler, PSTMessage pstMail, Metadata mailMetadata,
                                       EmbeddedDocumentExtractor embeddedExtractor) throws SAXException, IOException {

        Element root = new Element("Email");
        root.addNamespaceDeclaration("dc", DC_NS);
        root.addNamespaceDeclaration("meta", MS_MAPI_NS);
        root.addNamespaceDeclaration("message", MSG_NS);

        Element meta = new Element("Metadata");
        addNamespacedElement(meta, TikaCoreProperties.IDENTIFIER, pstMail.getInternetMessageId(), DC_NS);
        addNamespacedElement(meta, TikaCoreProperties.TITLE, pstMail.getSubject(), DC_NS);
        addElement(meta, Metadata.MESSAGE_FROM, pstMail.getSenderName());
        addNamespacedElement(meta, TikaCoreProperties.CREATOR, pstMail.getSentRepresentingName(), DC_NS);
        addNamespacedElement(meta, TikaCoreProperties.CREATED, pstMail.getCreationTime().toString(), DC_NS);
        addNamespacedElement(meta, TikaCoreProperties.MODIFIED, pstMail.getLastModificationTime().toString(), DC_NS);
        addNamespacedElement(meta, TikaCoreProperties.COMMENTS, pstMail.getComment(), DC_NS);
        addElement(meta, Metadata.MESSAGE_FROM, pstMail.getSenderName());
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
        root.appendChild(meta);

        /* Add email body as Base64 encoded data
        Element body = new Element("BodyAsBinary");
        body.appendChild(Base64.getEncoder().encodeToString(pstMail.getBody().getBytes(UTF_8)));
        root.appendChild(body); */

        Element body = new Element("Body");
        // TODO - parse out the XHTML content for this too..
        body.appendChild(CharMatcher.JAVA_ISO_CONTROL.removeFrom(new String(pstMail.getBody().getBytes(UTF_8))));
        root.appendChild(body);

        Document doc = new Document(root);

        // TODO - create a Thread (and pool) to do this work to speed up loading
        Session s = cs.newSession();
        String docUri = createDocUriFromId(pstMail.getInternetMessageId());
        Content c = ContentFactory.newContent(docUri, doc.toXML(), null);

        try {
            s.insertContent(c);
        } catch (RequestException e) {
            e.printStackTrace();
        }
        s.close();
        //LOG.debug(doc.toXML());

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
        id = id.replace("<","");
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
