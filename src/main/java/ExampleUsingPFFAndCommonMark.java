import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
import org.commonmark.node.Node;
import org.commonmark.parser.Parser;
import org.commonmark.renderer.html.HtmlRenderer;

import java.util.Vector;

public class ExampleUsingPFFAndCommonMark {
    public static void main(String[] args)
    {
        new ExampleUsingPFFAndCommonMark("path\\to\\your.pst");
    }

    public ExampleUsingPFFAndCommonMark(String filename) {
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

                Parser parser = Parser.builder().build();
                Node document = parser.parse(email.getBody());
                HtmlRenderer renderer = HtmlRenderer.builder().build();
                System.out.println(renderer.render(document));

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