import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.bind.JAXBElement;

import org.docx4j.TextUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;


public class DocxGen3 {

	MainDocumentPart documentPart;
	String startTag = "<%";
	String endTag   = "%>";
	Pattern p;
	ArrayList<String>  xpathList;
	StringBuilder paraStripped;
	
	public DocxGen3() throws Exception{	
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
		int mode;int endTagInd;
		for(Object oPara:listAllPara){
			xpathListUpdate(oPara); mode = 0;
			for(Object oRun : ((P)oPara).getContent()){
				if (oRun instanceof R){			
					for(Object oText : ((R)oRun).getContent()){						
						if (oText instanceof JAXBElement){
							cText = ((JAXBElement<Text>)oText).getValue();						
							cTextVal = cText.getValue();
							if (mode==1){
								endTagInd = cTextVal.indexOf('>',cIndex);
								if (endTagInd==-1) cText.setValue("");
								else 
							}else{
							while(!paraStripped.toString().startsWith(cTextVal)){
								int cIndex = 0;
								for (; cTextVal.charAt(cIndex) == paraStripped.charAt(cIndex); cIndex++);
								endTagInd = cTextVal.indexOf('>',cIndex);
								if (endTagInd == -1){
									cText.setValue(cTextVal.substring(0,cIndex) + "AlaBala");
									mode = 1;
								}else {
									cText.setValue(cTextVal.substring(0,cIndex) +
											"AlaBala" + cTextVal.substring(endTagInd));
								}
									paraStripped.replace(cIndex,cIndex+1,"AlaBala");
							}
							paraStripped.delete(0,cTextVal.length());
						}
						}
					}
				}
			}			
		}
	}
	
	private void xpathListUpdate(Object oPara) throws Exception{
		xpathList.clear();
		StringWriter swPara = new StringWriter();
		TextUtils.extractText(oPara,swPara);
		
		Matcher m = p.matcher(swPara.getBuffer());
		while(m.find())
			xpathList.add(m.group(1));
		paraStripped = new StringBuilder(m.replaceAll("#"));
//		for(String A:xpathList) System.out.println(A);
	}
	
	public static void main(String[] args) throws Exception{
		DocxGen3 D = new DocxGen3();
		D.filesManip();
	}
}
