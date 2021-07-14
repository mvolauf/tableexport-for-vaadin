package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.function.Consumer;
import java.util.function.Supplier;

import com.vaadin.server.StreamResource.StreamSource;

/**
 * StreamSource implementation based on top of temporary file
 */
public class TemporaryFileStreamSource implements StreamSource {

    /** The Constant serialVersionUID. */
    private static final long serialVersionUID = 3801605481686085335L;

    /** The input stream.
     *  Made it transient per: https://github.com/jnash67/tableexport-for-vaadin/issues/28
     */
    private final transient Supplier<File> fileProducer;
    private final transient Consumer<File> onCloseCallback;

    /**
     * Instantiates a new file stream resource.
     * 
     * @param fileProducer the file to download
     */
    public TemporaryFileStreamSource(Supplier<File> fileProducer) {
    	this(fileProducer, null);
    }

    /**
     * Instantiates a new file stream resource.
     * 
     * @param fileProducer the file to download
  	 * @param onCloseCallback callback invoked after the stream was closed (just prior to file delete)
     */
    public TemporaryFileStreamSource(Supplier<File> fileProducer, Consumer<File> onCloseCallback) {
        this.fileProducer = fileProducer;
        this.onCloseCallback = onCloseCallback;
    }

    @Override
    public InputStream getStream() {
        try {
			return new DeletingFileInputStream(fileProducer.get(), onCloseCallback);
		} catch (FileNotFoundException e) {
			return null;
		}
    }
}