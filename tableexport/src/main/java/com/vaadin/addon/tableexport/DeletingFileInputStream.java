package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.Serializable;
import java.util.function.Consumer;

/**
 * This input stream deletes the given file when the InputStream is closed;
 * intended to be used with temporary files.
 * 
 * Code obtained from:
 * http://vaadin.com/forum/-/message_boards/view_message/159583
 * 
 */
public class DeletingFileInputStream extends FileInputStream implements Serializable {

	private static final long serialVersionUID = 3840351665563343001L;

    private File file;
    private Consumer<File> onCloseCallback;

	/**
	 * Instantiates a new deleting file input stream.
	 * 
	 * @param file the file
	 * @param onCloseCallback callback invoked after the stream was closed (just prior to file delete)
	 * @throws FileNotFoundException the file not found exception
	 */
	public DeletingFileInputStream(File file, Consumer<File> onCloseCallback) throws FileNotFoundException {
		super(file);
		this.file = file;
		this.onCloseCallback = onCloseCallback;
	}

	@Override
	public void close() throws IOException {
		super.close();
		if (onCloseCallback != null) {
			onCloseCallback.accept(file);
		}
		file.delete();
	}
}