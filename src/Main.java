import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Map;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlString;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;

import com.google.cloud.translate.Detection;
import com.google.cloud.translate.Translate;
import com.google.cloud.translate.Translate.TranslateOption;
import com.google.cloud.translate.TranslateOptions;
import com.google.cloud.translate.Translation;

public class Main {

	public static final String API_KEY_ENV_NAME = "GOOGLE_API_KEY";
	static Translate translate;

	public static void main(String[] args) throws IOException {
		translate = TranslateOptions.getDefaultInstance().getService();
		

		

		XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream("MicrosoftPowerPoint.pptx"));
		for (XSLFSlide slide : slideShow.getSlides()) {
			CTSlide ctSlide = slide.getXmlObject();
			XmlObject[] allText = ctSlide.selectPath(
					"declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' " + ".//a:t");
			for (int i = 0; i < allText.length; i++) {
				if (allText[i] instanceof XmlString) {
					XmlString xmlString = (XmlString) allText[i];
					String text = xmlString.getStringValue();
					

					// Google Translate
					
					String result = translateText(text);
					
					xmlString.setStringValue(result);
//					}
				}
			}
		}

		FileOutputStream out = new FileOutputStream("MicrosoftPowerPointChanged.pptx");
		slideShow.write(out);
		slideShow.close();
		out.close();
	}

	private static String translateText(String text) {
		
		if (text.replace(" ", "").isEmpty()) {
			System.out.println("not translated");
			return text;
		}
		Detection detection = translate.detect(text);
		//String detectedLanguage = detection.getLanguage();
		Translation translation = translate.translate(text, TranslateOption.sourceLanguage("de"),
				TranslateOption.targetLanguage("en"));
		System.out.println(text + ":-translated-:"+translation.getTranslatedText());
		return translation.getTranslatedText();

	}

	
}
