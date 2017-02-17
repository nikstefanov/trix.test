
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
 
public class HelloMavenCentral {
 
    public static void main(String[] args) throws Exception {
 
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
 
        wordMLPackage.getMainDocumentPart()
            .addStyledParagraphOfText("Title", "Hello Maven Central");
 
        wordMLPackage.getMainDocumentPart().addParagraphOfText("from docx4j!");
 
        // Now save it
        wordMLPackage.save(new java.io.File(System.getProperty("user.dir") + "/helloMavenCentral.docx") );
 
    }
}
