import com.google.common.base.CharMatcher;
import com.marklogic.xcc.Content;
import com.marklogic.xcc.ContentFactory;
import com.marklogic.xcc.Session;
import com.marklogic.xcc.exceptions.RequestException;
import com.pff.PSTMessage;
import nu.xom.Document;
import nu.xom.Element;
import org.apache.commons.codec.digest.DigestUtils;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.*;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.parser.RecursiveParserWrapper;
import org.apache.tika.parser.microsoft.OutlookExtractor;
import org.apache.tika.sax.BasicContentHandlerFactory;
import org.apache.tika.sax.ContentHandlerFactory;
import org.apache.tika.sax.RecursiveParserWrapperHandler;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.lang.invoke.MethodHandles;

import static java.lang.String.valueOf;
import static java.nio.charset.StandardCharsets.UTF_8;

public class PSTProcessor implements Runnable {

    private static Logger LOG = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass().getCanonicalName());
    private static String DC_NS = "http://purl.org/dc/elements/1.1/";
    private static String MS_MAPI_NS = "http://schemas.microsoft.com/mapi";
    private static String MSG_NS = "URN:IANA:message:rfc822:";

    PSTMessage pstMail;

    PSTProcessor(PSTMessage pstMail) {
        //LOG.debug(String.format("Working on: %s", uri));
        this.pstMail = pstMail;
    }

    @Override
    public void run() {
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

        /* TODO? Add email body as Base64 encoded data
        Element body = new Element("BodyAsBinary");
        body.appendChild(Base64.getEncoder().encodeToString(pstMail.getBody().getBytes(UTF_8)));
        root.appendChild(body); */

//            RecursiveParserWrapper parser = new RecursiveParserWrapper(base,
//                    new BasicContentHandlerFactory(BasicContentHandlerFactory.HANDLER_TYPE.BODY, -1));
//            ctx.set(org.apache.tika.parser.Parser.class, parser);
//
//            Metadata metadata = new Metadata();
//            ContentHandler h = new BodyContentHandler(handler);
//            base.parse(new ByteArrayInputStream(mailContent), h, metadata, ctx);
            /*
            ContentHandler h2 = handler;
            ContentHandler textHandler = new BodyContentHandler(h2);
*/
        //parser.parse(new ByteArrayInputStream(mailContent), new BoilerpipeContentHandler(textHandler), metadata);
//            parser.reset();
//            LOG.info(h.toString());
        byte[] mailContent = pstMail.getBody().getBytes(UTF_8);
        Element HTMLbody = null;
        try {
            LOG.debug("****************** START Parsing Message Body **************");
            LOG.debug("ID: "+pstMail.getInternetMessageId());
            Parser p = new AutoDetectParser();
            ContentHandlerFactory factory = new BasicContentHandlerFactory(
                    BasicContentHandlerFactory.HANDLER_TYPE.XML, -1);
            RecursiveParserWrapper wrapper = new RecursiveParserWrapper(p);
            Metadata metadata = new Metadata();
            //metadata.set(TikaCoreProperties.RESOURCE_NAME_KEY, "test_recursive_embedded.docx");
            ParseContext context = new ParseContext();
            RecursiveParserWrapperHandler h2 = new RecursiveParserWrapperHandler(factory, -1);
            h2.startDocument();
            wrapper.parse(new ByteArrayInputStream(mailContent), h2, metadata, context);
            h2.endDocument();

            String HTMLMessage = h2.getMetadataList().get(0).get("X-TIKA:content");
            addElement(meta, "HTMLContentMD5", DigestUtils.md5Hex(HTMLMessage));

            HTMLbody = new Element("HTMLBody");
            HTMLbody.appendChild(CharMatcher.JAVA_ISO_CONTROL.removeFrom(new String(HTMLMessage)));
            LOG.debug("****************** END Parsing Message Body **************");
        } catch (TikaException | SAXException | IOException e) {
            LOG.error("TikaException / SAXException / IOException: ",e);
        } finally {
            // TODO Does anything need to be closed?
        }

        root.appendChild(meta);
        Element body = new Element("Body");
        body.appendChild(CharMatcher.JAVA_ISO_CONTROL.removeFrom(new String(mailContent)));
        root.appendChild(body);
        root.appendChild(HTMLbody);

        Document doc = new Document(root);

        // TODO - create a Thread (and pool) to do this work to speed up loading
        Session s = MarkLogicContentSourceProvider.getInstance().getContentSource().newSession();
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
}
