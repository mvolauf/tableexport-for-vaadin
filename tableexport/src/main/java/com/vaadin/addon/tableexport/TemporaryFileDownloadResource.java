package com.vaadin.addon.tableexport;

import com.vaadin.server.DownloadStream;
import com.vaadin.server.StreamResource;

import java.io.File;
import java.util.function.Consumer;
import java.util.function.Supplier;

/**
 * The Class TemporaryFileDownloadResource.
 * 
 * Code obtained from: http://vaadin.com/forum/-/message_boards/view_message/159583
 */
public class TemporaryFileDownloadResource extends StreamResource {

    /** The Constant serialVersionUID. */
    private static final long serialVersionUID = 476307190141362413L;

    /** The filename. */
    private final String filename;

    /** The content type. */
    private String contentType;

    /**
     * Instantiates a new temporary file download resource.
     * 
     * @param fileName
     *            the file name
     * @param contentType
     *            the content type
     * @param tempFile
     *            the temp file
     */
    public TemporaryFileDownloadResource(String fileName, String contentType, Supplier<File> tempFile)  {
    	this(fileName, contentType, tempFile, null);
    }

    /**
     * Instantiates a new temporary file download resource.
     * 
     * @param fileName
     *            the file name
     * @param contentType
     *            the content type
     * @param tempFile
     *            the temp file
  	 * @param onCloseCallback callback invoked after the stream was closed (just prior to file delete)
     */
    public TemporaryFileDownloadResource(String fileName, String contentType, 
    		Supplier<File> tempFile, Consumer<File> onCloseCallback)  {
        this(new TemporaryFileStreamSource(tempFile, onCloseCallback), fileName, contentType);
    }

    /**
     * Instantiates a new temporary file download resource.
     * 
     * 
     * @param streamSource 
     * @param fileName
     *            the file name
     * @param contentType
     *            the content type
     * @param tempFile
     *            the temp file
  	 * @param onCloseCallback callback invoked after the stream was closed (just prior to file delete)
     */
    public TemporaryFileDownloadResource(StreamSource streamSource, String fileName, String contentType)  {
        super(streamSource, fileName);
        this.filename = fileName;
        this.contentType = contentType;
    }

    @Override
    public DownloadStream getStream() {
        final DownloadStream stream =
                new DownloadStream(getStreamSource().getStream(), contentType, filename);
        stream.setParameter("Content-Disposition", "attachment;filename=" + filename);
        // This magic incantation should prevent anyone from caching the data
        stream.setParameter("Cache-Control", "private,no-cache,no-store");
        // In theory <=0 disables caching. In practice Chrome, Safari (and, apparently, IE) all
        // ignore <=0. Set to 1s
        stream.setCacheTime(1000);
        return stream;
    }

}
