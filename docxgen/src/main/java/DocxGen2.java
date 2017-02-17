import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.bind.JAXBElement;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathFactory;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.xml.sax.InputSource;


public class DocxGen2 {	
	
	MainDocumentPart documentPart;
	String startTag = "<%";
	String endTag   = "%>";
	Pattern p;
	XPath xpath;
	InputSource inputSource;
	
	public DocxGen2() throws Exception{	
		p = Pattern.compile(startTag+"(.*?)"+endTag);
		xpath = XPathFactory.newInstance().newXPath();
		inputSource = new InputSource(Thread.currentThread().getContextClassLoader()
				.getResource("Data.xml").toURI().toString());
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
		ArrayList<Text> textList = new ArrayList<Text>();
		String xpathString ;
		Matcher m;
		for(Object oPara:listAllPara){
			xpathString = "";textList = new ArrayList<Text>();
			for(Object oRun : ((P)oPara).getContent()){
				if (oRun instanceof R){			
					for(Object oText : ((R)oRun).getContent()){						
						if (oText instanceof JAXBElement){
							cText = ((JAXBElement<Text>)oText).getValue();						
							cTextVal = cText.getValue();
							startTagInd = cTextVal.indexOf(startTag,0);
							endTagInd = cTextVal.indexOf(endTag,0);
							if (startTagInd == -1 && endTagInd == -1){
								if (textList.size()>0){
									textList.add(cText);
									xpathString += cTextVal;
								}
							}else if (startTagInd == -1 ||
									endTagInd >= 0 && endTagInd < startTagInd){
								if (textList.size()>0){
									xpathString += cTextVal.substring(0,endTagInd);
									textList.get(0).setValue(
											textList.get(0).getValue().substring(0,
													textList.get(0).getValue().lastIndexOf(startTag)) +
													xpath.evaluate(xpathString, inputSource));
									for (int tLind = 1; tLind < textList.size() ; tLind++)
										textList.get(tLind).setValue("");
									cText.setValue(cTextVal.substring(endTagInd + endTag.length()));
									cTextVal = cText.getValue();
									xpathString = "";textList.clear();
								}
							}							
							m = p.matcher(cTextVal);
							while(m.find()){								
								cText.setValue(cTextVal.replace(startTag+m.group(1)+endTag,
										xpath.evaluate(m.group(1), inputSource)));
								cTextVal = cText.getValue();
								xpathString = "";textList.clear();
								m = p.matcher(cTextVal);
							}
							if ((startTagInd = cTextVal.lastIndexOf(startTag))>=0){
								textList.add(cText);
								xpathString = cTextVal.substring(startTagInd + startTag.length());
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
