package core;

import java.io.File;
import javax.swing.filechooser.FileFilter;
 
public class FileTypeFilter extends FileFilter {
	//source http://www.codejava.net/java-se/swing/add-file-filter-for-jfilechooser-dialog
    private String extension;
    private String description;
 
    public FileTypeFilter(String extension, String description) {
        this.extension = extension;
        this.description = description;
    }
    
    @Override
    public boolean accept(File f) {
        if (f.isDirectory()) {
            return true;
        }
        return f.getName().endsWith(extension);
    }
 
    public String getDescription() {
        return description + String.format(" (*%s)", extension);
    }
}