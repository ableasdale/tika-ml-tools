import com.pff.*;
import org.apache.tika.exception.TikaException;
import org.apache.tika.extractor.EmbeddedDocumentExtractor;
import org.apache.tika.extractor.EmbeddedDocumentUtil;
import org.apache.tika.io.TikaInputStream;
import org.apache.tika.metadata.Message;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.metadata.Office;
import org.apache.tika.metadata.TikaCoreProperties;
import org.apache.tika.mime.MediaType;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.mbox.OutlookPSTParser;
import org.apache.tika.parser.microsoft.OutlookExtractor;
import org.apache.tika.sax.ToXMLContentHandler;
import org.apache.tika.sax.XHTMLContentHandler;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Vector;

import static java.lang.String.valueOf;
import static java.nio.charset.StandardCharsets.UTF_8;

public class ParsePSTFolder {
   // MediaType MS_OUTLOOK_PST_MIMETYPE = MediaType.application("vnd.ms-outlook-pst");
//    Set<MediaType> SUPPORTED_TYPES = singleton(MS_OUTLOOK_PST_MIMETYPE);

    public static void main(String[] args) throws Exception {

        Metadata metadata = new Metadata();
        ContentHandler handler = new ToXMLContentHandler();

        ParseContext context = new ParseContext();
        EmbeddedDocumentExtractor embeddedExtractor = EmbeddedDocumentUtil.getEmbeddedDocumentExtractor(context);

        metadata.set(Metadata.CONTENT_TYPE, OutlookPSTParser.MS_OUTLOOK_PST_MIMETYPE.toString());

        //XMLContent
        XHTMLContentHandler xml = new XHTMLContentHandler(handler, metadata);
        
        TikaInputStream in = TikaInputStream.get(new FileInputStream("your.pst"));
        PSTFile pstFile = null;
        try {
            pstFile = new PSTFile(in.getFile().getPath());
            metadata.set(Metadata.CONTENT_LENGTH, valueOf(pstFile.getFileHandle().length()));
            boolean isValid = pstFile.getFileHandle().getFD().valid();
            metadata.set("isValid", valueOf(isValid));
            if (isValid) {
                System.out.println(pstFile.getRootFolder().getEmailAddress());
                System.out.println(pstFile.getRootFolder().getSubFolderCount());
                System.out.println(pstFile.getRootFolder().getSubFolderCount());
                System.out.println(pstFile.getRootFolder().getItemsString());
                System.out.println(pstFile.getRootFolder().getAssociateContentCount());
                System.out.println(pstFile.getRootFolder().getContentCount());

                Vector<PSTFolder> folders = pstFile.getRootFolder().getSubFolders();
                for (PSTFolder f : folders){
                    System.out.println("subfolder");
                    System.out.println(f.getDisplayName());
                    System.out.println(f.getEmailAddress());
                    System.out.println(f.getContentCount());
                    System.out.println(f.getSubFolderCount());
                    System.out.println(f.getUnreadCount());
                    System.out.println(f.getCreationTime());
                    System.out.println(f.getItemsString());
                    System.out.println(f.getAssociateContentCount());
                    System.out.println(f.getContainerClass());
                }
               parseFolder(xml, pstFile.getRootFolder(), embeddedExtractor);
            }
        } catch (Exception e) {
            throw new TikaException(e.getMessage(), e);
        } finally {
            if (pstFile != null && pstFile.getFileHandle() != null) {
                try {
                    pstFile.getFileHandle().close();
                } catch (IOException e) {
                    //swallow closing exception
                }
            }
        }
    }

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
       // mailMetadata.set(TikaCoreProperties.RESOURCE_NAME_KEY, pstMail.getInternetMessageId());
       // mailMetadata.set(TikaCoreProperties.EMBEDDED_RELATIONSHIP_ID, pstMail.getInternetMessageId());
        System.out.println(pstMail.getInternetMessageId() +" | "+ pstMail.getSenderName() +" | "+ pstMail.getSubject()+" | "+ pstMail.getCreationTime());
        mailMetadata.set(TikaCoreProperties.IDENTIFIER, pstMail.getInternetMessageId());
        mailMetadata.set(TikaCoreProperties.TITLE, pstMail.getSubject());
        mailMetadata.set(Metadata.MESSAGE_FROM, pstMail.getSenderName());
        mailMetadata.set(TikaCoreProperties.CREATOR, pstMail.getSenderName());
        mailMetadata.set(TikaCoreProperties.CREATED, pstMail.getCreationTime());
        mailMetadata.set(TikaCoreProperties.MODIFIED, pstMail.getLastModificationTime());
        mailMetadata.set(TikaCoreProperties.COMMENTS, pstMail.getComment());
        mailMetadata.set("descriptorNodeId", valueOf(pstMail.getDescriptorNodeId()));
        mailMetadata.set("senderEmailAddress", pstMail.getSenderEmailAddress());
        mailMetadata.set("recipients", pstMail.getRecipientsString());
        mailMetadata.set("displayTo", pstMail.getDisplayTo());
        mailMetadata.set("displayCC", pstMail.getDisplayCC());
        mailMetadata.set("displayBCC", pstMail.getDisplayBCC());
        mailMetadata.set("importance", valueOf(pstMail.getImportance()));
        mailMetadata.set("priority", valueOf(pstMail.getPriority()));
        mailMetadata.set("flagged", valueOf(pstMail.isFlagged()));
        mailMetadata.set(Office.MAPI_MESSAGE_CLASS,
                OutlookExtractor.getMessageClass(pstMail.getMessageClass()));

        mailMetadata.set(Message.MESSAGE_FROM_EMAIL, pstMail.getSenderEmailAddress());

        mailMetadata.set(Office.MAPI_FROM_REPRESENTING_EMAIL,
                pstMail.getSentRepresentingEmailAddress());

        mailMetadata.set(Message.MESSAGE_FROM_NAME, pstMail.getSenderName());
        mailMetadata.set(Office.MAPI_FROM_REPRESENTING_NAME, pstMail.getSentRepresentingName());

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
            //swallow
        }
        //we may want to experiment with working with the bodyHTML.
        //However, because we can't get the raw bytes, we _could_ wind up sending
        //a UTF-8 byte representation of the html that has a conflicting metaheader
        //that causes the HTMLParser to get the encoding wrong.  Better if we could get
        //the underlying bytes from the pstMail object...

        byte[] mailContent = pstMail.getBody().getBytes(UTF_8);

       // System.out.println(pstMail.getBody());
        mailMetadata.set(TikaCoreProperties.CONTENT_TYPE_OVERRIDE, MediaType.TEXT_PLAIN.toString());
        //embeddedExtractor.parseEmbedded(new ByteArrayInputStream(mailContent), handler, mailMetadata, true);
    }

    /*
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
