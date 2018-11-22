import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
import com.vladsch.flexmark.ast.Node;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.options.MutableDataSet;


import java.util.Vector;

public class ExampleUsingPFFAndFlexMark {

    public static void main(String[] args)
    {
        new ExampleUsingPFFAndCommonMark("path\\to\\your.pst");
    }

    public ExampleUsingPFFAndFlexMark(String filename) {
        try {
            PSTFile pstFile = new PSTFile(filename);
            System.out.println(pstFile.getMessageStore().getDisplayName());
            processFolder(pstFile.getRootFolder());
        } catch (Exception err) {
            err.printStackTrace();
        }
    }

    int depth = -1;
    public void processFolder(PSTFolder folder)
            throws PSTException, java.io.IOException
    {
        depth++;
        // the root folder doesn't have a display name
        if (depth > 0) {
            printDepth();
            System.out.println(folder.getDisplayName());
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
            PSTMessage email = (PSTMessage)folder.getNextChild();
            while (email != null) {
                printDepth();
                System.out.println("Email: "+email.getSubject());
                //System.out.println(email.getBody());
                MutableDataSet options = new MutableDataSet();
                Parser parser = Parser.builder(options).build();
               // Node document = parser.parse(email.getBody());
               HtmlRenderer renderer = HtmlRenderer.builder(options).build();
               // System.out.println(renderer.render(document));

                Node document = parser.parse(email.getBody());
                String html = renderer.render(document);  // "<p>This is <em>Sparta</em></p>\n"
                System.out.println(html);

                //System.out.println(email.getBodyHTML());
                email = (PSTMessage)folder.getNextChild();
            }
            depth--;
        }
        depth--;
    }

    public void printDepth() {
        for (int x = 0; x < depth-1; x++) {
            System.out.print(" | ");
        }
        System.out.print(" |- ");
    }
}
