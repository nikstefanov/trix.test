package trix_test_docgen;

//import org.docx4j.model.datastorage.CustomXmlDataStorage;
//import org.docx4j.openpackaging.parts.CustomXmlDataStoragePart;
import org.docx4j.model.datastorage.BindingHandler;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
/**
 * This class tries out some lines of code found in Getting Started Document of
 * docx4, Text substitution via data bound content controls paragraph.
 * @author User
 *
 */
public class Test {

  public static void main(String[] args) throws Exception{
    WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(
        new java.io.File(
            "C:\\Users\\User\\Documents\\Development\\DataBinding\\OpenDoPE\\"
            + "invoice.docx" ));           
   MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
    
//    CustomXmlDataStoragePart customXmlDataStoragePart =
//      wordMLPackage.getCustomXmlDataStorageParts().get(1);
//    CustomXmlDataStorage customXmlDataStorage =
//      customXmlDataStoragePart.getData();
    
   BindingHandler.applyBindings(documentPart);
    wordMLPackage.save(new java.io.File(
        "C:\\Users\\User\\Documents\\Development\\DataBinding\\OpenDoPE\\"
            + "invoiceApplied.docx" )); 

  }

}
