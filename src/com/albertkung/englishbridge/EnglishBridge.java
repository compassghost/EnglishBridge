package com.albertkung.englishbridge;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map.Entry;

import javax.swing.JFileChooser;
import javax.xml.bind.JAXBElement;

import org.apache.commons.io.FileUtils;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ContentAccessor;

public class EnglishBridge {
	static final org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
	static final String COMMENT_NAME = "EnglishBridge";

	HashMap<String, String> britishToAmerican = new HashMap<>();
	HashMap<String, String> americanToBritish = new HashMap<>();
	HashMap<String, String> dictionary;
	private boolean toBritish = true;
	private File copied;
	private MainDocumentPart documentPart;

	public EnglishBridge() {
		try {
			createDictionary("dictionary.csv");
			JFileChooser fc = new JFileChooser();
			int returnVal = fc.showOpenDialog(null);
			if (returnVal == JFileChooser.APPROVE_OPTION) {
				File file = fc.getSelectedFile();
				scanDocument(file);
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void scanDocument(File file) throws Exception {
		if (toBritish)
			dictionary = americanToBritish;
		else
			dictionary = britishToAmerican;
		copied = new File(file.getAbsolutePath().replace(".docx",
				"_swapped.docx"));
		FileUtils.copyFile(file, copied);

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
				.load(copied);
		documentPart = wordMLPackage.getMainDocumentPart();
		replaceText(documentPart);
		wordMLPackage.save(copied);
	}

	private void replaceText(ContentAccessor c) throws Exception {
		for (Object p : c.getContent()) {
			if (p instanceof ContentAccessor)
				replaceText((ContentAccessor) p);

			else if (p instanceof JAXBElement) {
				Object v = ((JAXBElement) p).getValue();

				if (v instanceof ContentAccessor)
					replaceText((ContentAccessor) v);

				else if (v instanceof org.docx4j.wml.Text) {
					org.docx4j.wml.Text t = (org.docx4j.wml.Text) v;
					System.out.println(t.getValue());
					replaceParams(t);
				}
			}
		}
	}

	private void replaceParams(org.docx4j.wml.Text t)
			throws InvalidFormatException {
		String text = t.getValue();
		if (text != null) {
			t.setSpace("preserve"); // needed?
			for (Entry<String, String> entry : dictionary.entrySet()) {
				String key = entry.getKey();
				String value = entry.getValue();
				int searchParam = 0;
				while (text.indexOf(key, searchParam) >= 0) {
					searchParam = text.indexOf(key, searchParam);
					System.out.println("Found " + key);
					int searchTerm = text.indexOf(key, searchParam);
					if (text.indexOf(value, searchParam) == searchTerm) {
						// likely an extended word, i.e. furor vs
						// furore
					} else if (text.indexOf(value, searchTerm - value.length()
							+ key.length()) == searchTerm) {
					} else {
						for (int i = 0; i < key.length(); i++) {
							int endString = searchParam + key.length() + 1;
							text = text.substring(0, searchParam).toString()
									+ value
									+ (endString < text.length() ? text
											.substring(
													searchParam + key.length()
															+ 1).toString()
											: "");
						}
					}
					searchParam = searchParam + value.length();
				}
			}
			System.out.println("New text: " + text);
			t.setValue(text);
		}
	}

	public void createDictionary(String fileLocation) throws Exception {
		System.out.println(Paths.get(".").toAbsolutePath().normalize()
				.toString());
		britishToAmerican.clear();
		americanToBritish.clear();
		BufferedReader br = new BufferedReader(new FileReader(fileLocation));
		try {
			String line = br.readLine();
			while (line != null) {
				String[] words = line.split(",");
				britishToAmerican.put(words[0].trim(), words[1].trim());
				americanToBritish.put(words[1].trim(), words[0].trim());
				line = br.readLine();
			}
		} finally {
			br.close();
		}
	}
}
