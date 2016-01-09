package com.ksls;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Copyright {

	private static final int maxRowCount = 50;
	private static final int maxFileCount = 80;
	private static final String templateFile = "F:\\template.docx";
	
	private static int fileCount = 0;
	private static String basePath = "C:\\Users\\bin.shen\\git\\sdbb2\\";
	private static String[] folders = new String[] { 
		"application\\controllers", 
		"application\\models",
		"application\\views", 
		"application\\core", 
		"application\\helpers", 
		"application\\hooks",
		"application\\libraries", 
		"assets", 
		"dwz", 
		"system" 
	};

	public static void main(String[] args) {

		try {
			String filePath = "F:\\报备系统源码.docx";
			createFile(filePath);

			OPCPackage pack = POIXMLDocument.openPackage(filePath);
			XWPFDocument word = new XWPFDocument(pack);
			XWPFParagraph paragraph = word.getParagraphs().get(17);
			paragraph.setIndentationLeft(0);
			paragraph.setIndentationHanging(0);
			paragraph.setAlignment(ParagraphAlignment.LEFT);

			XWPFRun run = paragraph.insertNewRun(0);

			run.setFontSize(10);
			run.setBold(false);

			for (int i = 0; i < folders.length; i++) {
				if (fileCount > maxFileCount)
					break;

				String path = basePath + folders[i];
				readDirectories(run, path);
			}

			String tempFile = "F:\\temp.docx";
			File newFile = new File(tempFile);
			FileOutputStream fos = new FileOutputStream(newFile);
			newFile.deleteOnExit();
			word.write(fos);
			fos.flush();
			fos.close();
			pack.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void createFile(String filePath) {
		try {
			int byteread = 0;
			File oldfile = new File(templateFile);
			if (oldfile.exists()) {
				InputStream inStream = new FileInputStream(templateFile);
				FileOutputStream fs = new FileOutputStream(filePath);
				byte[] buffer = new byte[1444];
				while ((byteread = inStream.read(buffer)) != -1) {
					fs.write(buffer, 0, byteread);
				}
				inStream.close();
				fs.close();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void readDirectories(XWPFRun run, String path) throws Exception {

		File file = new File(path);
		File[] tempList = file.listFiles();
		System.out.println(path);
		for (int i = 0; i < tempList.length; i++) {
			if (tempList[i].isFile()) {
				String extension = getExtensionName(tempList[i].getName());
				if ("php".equalsIgnoreCase(extension) || "html".equalsIgnoreCase(extension) || "js".equalsIgnoreCase(extension)) {
					fileCount++;
					if (fileCount > maxFileCount) break;

					readFile(run, tempList[i].getPath());
				}
			} else if (tempList[i].isDirectory()) {
				readDirectories(run, tempList[i].getPath());
			}
		}
	}

	private static void readFile(XWPFRun run, String filePath) throws Exception {

		File file = new File(filePath);
		run.setText(fileCount + " - " + file.getName());
		run.addBreak();

		BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(file), "UTF-8"));
		String s;
		int rowCount = 0;
		while ((s = reader.readLine()) != null && rowCount <= maxRowCount) {
			s = s.replaceAll("\t", "    ");
			run.setText(s + "\r\n");
			run.addBreak();
			rowCount++;
		}
		reader.close();
		run.addBreak();
	}

	public static String getExtensionName(String filename) {
		if ((filename != null) && (filename.length() > 0)) {
			int dot = filename.lastIndexOf('.');
			if ((dot > -1) && (dot < (filename.length() - 1))) {
				return filename.substring(dot + 1);
			}
		}
		return filename;
	}
}
