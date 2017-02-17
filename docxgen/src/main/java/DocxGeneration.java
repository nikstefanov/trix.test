import java.io.StringWriter;
import java.util.List;
import java.util.ArrayList;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

import javax.xml.bind.JAXBElement;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.docx4j.TextUtils;


public class DocxGeneration {
	
	enum flowMode {NORMAL,TAG}
	Pattern p;
	MainDocumentPart documentPart;
	ArrayList<String>  xpathList;
	String startTag = "<%";
	String endTag   = "%>";
	
	public DocxGeneration() throws Exception{
		p = Pattern.compile(startTag+"(.*?)"+endTag);
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(
				new java.io.File(
						Thread.currentThread().getContextClassLoader()
						.getResource("TemplateDocx.docx").toURI()));
		documentPart = wordMLPackage.getMainDocumentPart();
	}

	/**
	 * @param args
	 */	
	public static void main(String[] args) throws Exception{
		DocxGeneration D = new DocxGeneration();
		D.iteratePara();
	}
	
	@SuppressWarnings("unchecked")
	public void iteratePara() throws Exception{		
		String xpathAllPara = "//w:p";
		List<Object> listAllPara = documentPart.getJAXBNodesViaXPath(xpathAllPara, false);
		Text cText;String cTextVal;
		int startTagInd,endTagInd,flowMode = 0;	// 0 - normal flow, 1 - tag mode
		for(Object oPara:listAllPara){
			xpathListUpdate(oPara);
			if (xpathList.size()>0){
			for(Object oRun : ((P)oPara).getContent()){
				if (oRun instanceof R){			
					for(Object oText : ((R)oRun).getContent()){						
						if (oText instanceof JAXBElement){
							cText = ((JAXBElement<Text>)oText).getValue();						
							cTextVal = cText.getValue();							
							startTagInd=cTextVal.indexOf(startTag);
							endTagInd = cTextVal.indexOf(endTag);
//							if (flowMode == 0 && )	
							}							
						}
					}
				}
			}			
		}
	}
	
	/**
	 *  Puts all the xpaths from the current paragraph in <code>xpathList</code> list.
	 *  
	 *  @param the current paragraph object
	 */
	private void xpathListUpdate(Object oPara) throws Exception{
		xpathList.clear();
		StringWriter swPara = new StringWriter();
		TextUtils.extractText(oPara,swPara);
		
		Matcher m = p.matcher(swPara.getBuffer());
		while(m.find())
			xpathList.add(m.group(1));
//		for(String A:xpathList) System.out.println(A);
	}
		

	
	
	 

}
