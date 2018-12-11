package develop;

import org.apache.tika.exception.TikaException;
import org.apache.tika.extractor.EmbeddedDocumentExtractor;
import org.apache.tika.extractor.ParsingEmbeddedDocumentExtractor;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.sax.ToHTMLContentHandler;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class ExamplePSTExtraction {
    public static void main(String[] args) {

        Parser pstParser = new AutoDetectParser();
        Metadata metadata = new Metadata();
        ContentHandler handler = new ToHTMLContentHandler();

        ParseContext context = new ParseContext();
        EmbeddedTrackingExtrator trackingExtrator = new EmbeddedTrackingExtrator(context);
        context.set(EmbeddedDocumentExtractor.class, trackingExtrator);
        context.set(Parser.class, new AutoDetectParser());

        try {
            pstParser.parse(new FileInputStream("E:\\RevisedEDRMv1_Complete\\andrea_ring\\andrea_ring_000_1_1.pst"), handler, metadata, context);
            //handler.
            String output = handler.toString();
            System.out.println(output);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (TikaException e) {
            e.printStackTrace();
        }
    }

    private static class EmbeddedTrackingExtrator extends ParsingEmbeddedDocumentExtractor {
        List<Metadata> trackingMetadata = new ArrayList<Metadata>();

        public EmbeddedTrackingExtrator(ParseContext context) {
            super(context);
        }

        @Override
        public  boolean shouldParseEmbedded(Metadata metadata) {
            return true;
        }

        @Override
        public  void parseEmbedded(InputStream stream, ContentHandler handler, Metadata metadata, boolean outputHtml) throws SAXException, IOException {
            this.trackingMetadata.add(metadata);
            super.parseEmbedded(stream, handler, metadata, outputHtml);
        }
    }
}
