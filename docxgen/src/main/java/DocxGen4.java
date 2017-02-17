import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.bind.JAXBElement;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;


public class DocxGen4{	
	
	MainDocumentPart documentPart;
	String startTag = "<%";
	String endTag   = "%>";
	Pattern p;
	
	public DocxGen4() throws Exception{	
		p = Pattern.compile(startTag+"(.*?)"+endTag);
	}
	
	public void filesManip() throws Exception{
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(
				new java.io.File(
						Thread.currentThread().getContextClassLoader()
						.getResource("TemplateDocx.docx").toURI()));
		documentPart = wordMLPackage.getMainDocumentPart();
		iteratePara();
        wordMLPackage.save(new java.io.File(
        		System.getProperty("user.dir") + "/FilledDocx.docx") );
	}
	
	@SuppressWarnings("unchecked")
	public void iteratePara() throws Exception{		
		String xpathAllPara = "//w:p";
		List<Object> listAllPara = documentPart.getJAXBNodesViaXPath(xpathAllPara, false);
		Text cText;String cTextVal;
		int startTagInd,endTagInd;
		ArrayList<Text> textList;
		String xpathString;
		Matcher m;int mode;
		for(Object oPara:listAllPara){
			mode = 0;
			xpathString = "";textList = new ArrayList<Text>();
			for(Object oRun : ((P)oPara).getContent()){
				if (oRun instanceof R){			
					for(Object oText : ((R)oRun).getContent()){						
						if (oText instanceof JAXBElement){
							cText = ((JAXBElement<Text>)oText).getValue();						
							cTextVal = cText.getValue();
							
							if (mode ==0 && cTextVal.indexOf(startTag)>=0){
								mode = 2;
							}
						}
					}
				}
			}			
		}
	}
	
	public static void main(String[] args) throws Exception{
		DocxGen2 D = new DocxGen2();
		D.filesManip();
	}
}
