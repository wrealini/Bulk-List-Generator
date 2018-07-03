package core;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.TreeMap;
import java.util.TreeSet;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.parser.image.TiffParser;
import org.apache.tika.parser.jpeg.JpegParser;
import org.apache.tika.sax.BodyContentHandler;
import org.xml.sax.ContentHandler;

import com.google.common.io.Files;

public class Main {

	public static void main(String[] args) {
		try{
			JFileChooser jfc = new JFileChooser();
			jfc.setDialogTitle("Save Index");
			jfc.setToolTipText("Bulk List Generator by Wallace R. Realini III");
			jfc.setCurrentDirectory(new File(System.getProperty("user.dir")));
			jfc.setFileFilter(new FileTypeFilter(".xls", "Microsoft Excel Documents"));
			if(jfc.showSaveDialog(null) != JFileChooser.APPROVE_OPTION)
				return;
			File file = jfc.getSelectedFile();
			file = new File(file.getPath().split("\\.", 2)[0] + ".xls");
			File cd = new File(file.getParentFile().getParent() + "\\Conveyors [Duplicates]");
			TreeMap<String, TreeSet<String>> conveyors = null;
			if(cd.exists())
				conveyors  = new TreeMap<String, TreeSet<String>>();
			File ttd = new File(file.getParentFile().getParent() + "\\Transfer Towers [Duplicates]");
			TreeMap<String, TreeSet<String>> transfertowers = null;
			if(ttd.exists())
				transfertowers  = new TreeMap<String, TreeSet<String>>();
			String prefix = null;
			
			HashMap<String, TreeMap<String, TreeSet<String>>> spmap = new HashMap<String, TreeMap<String, TreeSet<String>>>();
			HashMap<String, Metadata> filemap = new HashMap<String, Metadata>();
			for(File tf: file.getParentFile().listFiles()) {
				if(!tf.getName().endsWith(".tif")) if(!tf.getName().endsWith(".jpg"))
					continue;
				InputStream input = new FileInputStream(tf);
				ContentHandler textHandler = (ContentHandler) new BodyContentHandler();
				Metadata metadata = new Metadata();
				metadata.set(Metadata.RESOURCE_NAME_KEY, tf.getPath());
		
				Parser parser;
				if(tf.getName().endsWith(".tif"))
					parser = new TiffParser();
				else
					parser = new JpegParser();
				parser.parse(input, textHandler, metadata, new ParseContext());
				input.close();
				
				String subject = metadata.get("Windows XP Subject");
				if(subject == null)
					continue;
				TreeMap<String, TreeSet<String>> tm = spmap.get(subject);
				String page = metadata.get("Windows XP Title");
				if(page == null)
					continue;
				if(prefix == null)
					prefix = page.split(" - ", 2)[0];
				if(tm == null)
					tm = new TreeMap<String, TreeSet<String>>();
				TreeSet<String> ts = tm.get(page);
				if(ts == null)
					ts = new TreeSet<String>();
				ts.add(tf.getName());
				tm.put(page, ts);
				spmap.put(subject, tm);
				filemap.put(tf.getName(), metadata);
			}
			
			Workbook wb = new HSSFWorkbook();
			
			for(String subject: spmap.keySet()) {
				Sheet sheet = wb.createSheet(subject);
				
				Row titles = sheet.createRow((short)0);
				titles.createCell(0).setCellValue("Page Number");
				titles.createCell(1).setCellValue("Equipment");
				titles.createCell(2).setCellValue("File Name");
				titles.createCell(3).setCellValue("Author");
				titles.createCell(4).setCellValue("Date");
				titles.createCell(5).setCellValue("Physical Location");
				
				short r = 1;
				TreeMap<String, TreeSet<String>> tm = spmap.get(subject);
				for(String page: tm.keySet()) for(String filename: tm.get(page)) {
					Metadata m = filemap.get(filename);
					String date = m.get("Creation-Date");
					date = date.substring(0, 10);
					date = date.replace('-', '/');
					String keywords = m.get("Windows XP Keywords"), equipment = "", physicallocation = "";
					if(keywords == null)
						;
					else for(String tag: keywords.split(";")) if(tag.startsWith("Roll"))
						physicallocation += tag + "; ";
					else if(tag.contains("-"))
						physicallocation += tag + "; ";
					else {
						equipment += tag + "; ";
						if(conveyors != null) if(tag.startsWith("Conveyor ")) {
							TreeSet<String> temp = conveyors.get(tag);
							if(temp == null)
								temp = new TreeSet<String>();
							temp.add(filename);
							conveyors.put(tag, temp);
						}
						if(transfertowers != null) if(tag.startsWith("Transfer Tower ")) {
							TreeSet<String> temp = transfertowers.get(tag);
							if(temp == null)
								temp = new TreeSet<String>();
							temp.add(filename);
							transfertowers.put(tag, temp);
						}
					}
					if(!equipment.equals(""))
						equipment = equipment.substring(0, equipment.length() - 2);
					if(!physicallocation.equals(""))
						physicallocation = physicallocation.substring(0, physicallocation.length() - 2);
					
					Row row = sheet.createRow(r);
					row.createCell(0).setCellValue(page);
					row.createCell(1).setCellValue(equipment);
					row.createCell(2).setCellFormula("HYPERLINK(\"" + filename + "\", \"" + filename + "\")");
					row.createCell(3).setCellValue(m.get("Windows XP Author"));
					row.createCell(4).setCellValue(date);
					row.createCell(5).setCellValue(physicallocation);
					r++;
				}
			}
			
			if(conveyors != null) for(String tag: conveyors.keySet()) {
				File subdirectory = new File(cd.getPath() + "\\" + tag);
				if(subdirectory.exists()) for(File sdf: subdirectory.listFiles()) {
					Parser parser;
					if(sdf.getName().endsWith(".tif"))
						parser = new TiffParser();
					else if(sdf.getName().endsWith(".jpg"))
						parser = new JpegParser();
					else continue;
					InputStream input = new FileInputStream(sdf);
					ContentHandler textHandler = (ContentHandler) new BodyContentHandler();
					Metadata metadata = new Metadata();
					metadata.set(Metadata.RESOURCE_NAME_KEY, sdf.getPath());
					parser.parse(input, textHandler, metadata, new ParseContext());
					input.close();
					if(metadata.get("Windows XP Title").startsWith(prefix)) try {
						sdf.delete();
					} catch(Exception ex) {}
				} else
					subdirectory.mkdirs();
				for(String fn: conveyors.get(tag)) {
					File copyfrom = new File(file.getParent() + "\\" + fn);
					File copyto = new File(subdirectory.getPath() + "\\" + fn);
					try {
						Files.copy(copyfrom, copyto);
					} catch(Exception ex) {}
				}
			}
			if(transfertowers != null) for(String tag: transfertowers.keySet()) {
				File subdirectory = new File(ttd.getPath() + "\\" + tag);
				if(subdirectory.exists()) for(File sdf: subdirectory.listFiles()) {
					Parser parser;
					if(sdf.getName().endsWith(".tif"))
						parser = new TiffParser();
					else if(sdf.getName().endsWith(".jpg"))
						parser = new JpegParser();
					else continue;
					InputStream input = new FileInputStream(sdf);
					ContentHandler textHandler = (ContentHandler) new BodyContentHandler();
					Metadata metadata = new Metadata();
					metadata.set(Metadata.RESOURCE_NAME_KEY, sdf.getPath());
					parser.parse(input, textHandler, metadata, new ParseContext());
					input.close();
					if(metadata.get("Windows XP Title").startsWith(prefix)) try {
						sdf.delete();
					} catch(Exception ex) {}
				} else
					subdirectory.mkdirs();
				for(String fn: transfertowers.get(tag)) {
					File copyfrom = new File(file.getParent() + "\\" + fn);
					File copyto = new File(subdirectory.getPath() + "\\" + fn);
					try {
						Files.copy(copyfrom, copyto);
					} catch(Exception ex) {}
				}
			}
			
		    FileOutputStream fileOut = new FileOutputStream(file);
		    wb.write(fileOut);
		    fileOut.close();
		    wb.close();
		    
		    Desktop.getDesktop().open(file);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(null, e.getClass().toGenericString());
			e.printStackTrace();
		} catch (Error e) {
			JOptionPane.showMessageDialog(null, e.getClass().toGenericString());
		}
	}
}
