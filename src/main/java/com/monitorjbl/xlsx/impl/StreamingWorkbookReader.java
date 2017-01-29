package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.StreamingReader.Builder;
import com.monitorjbl.xlsx.exceptions.OpenException;
import com.monitorjbl.xlsx.exceptions.ReadException;
import com.monitorjbl.xlsx.sst.BufferedStringsTable;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static com.monitorjbl.xlsx.XmlUtils.document;
import static com.monitorjbl.xlsx.XmlUtils.searchForNodeList;
import static java.util.Arrays.asList;

public class StreamingWorkbookReader implements Iterable<Sheet>, AutoCloseable {
  private static final Logger log = LoggerFactory.getLogger(StreamingWorkbookReader.class);

  private final List<StreamingSheet> sheets;
  private final List<Map<String, String>> sheetProperties = new ArrayList<Map<String, String>>();
  private final Builder builder;
  private File tmp;
  private File sstCache;
  private OPCPackage pkg;
  private SharedStringsTable sst;
  private boolean use1904Dates = false;

  /**
   * This constructor exists only so the StreamingReader can instantiate
   * a StreamingWorkbook using its own reader implementation. Do not use
   * going forward.
   *
   * @param sst      The SST data for this workbook
   * @param sstCache The backing cache file for the SST data
   * @param pkg      The POI package that should be closed when this workbook is closed
   * @param reader   A single streaming reader instance
   * @param builder  The builder containing all options
   */
  @Deprecated
  public StreamingWorkbookReader(SharedStringsTable sst, File sstCache, OPCPackage pkg, StreamingSheetReader reader, Builder builder) {
    this.sst = sst;
    this.sstCache = sstCache;
    this.pkg = pkg;
    this.sheets = asList(new StreamingSheet(null, reader));
    this.builder = builder;
  }

  public StreamingWorkbookReader(Builder builder) {
    this.sheets = new ArrayList<StreamingSheet>();
    this.builder = builder;
  }

  public StreamingSheetReader first() {
    return sheets.get(0).getReader();
  }

  public void init(InputStream is) {
    File f = null;
    try {
      f = writeInputStreamToFile(is, builder.getBufferSize());
      log.debug("Created temp file [" + f.getAbsolutePath() + "]");

      init(f);
      tmp = f;
    } catch(IOException e) {
      throw new ReadException("Unable to read input stream", e);
    } catch(RuntimeException e) {
      f.delete();
      throw e;
    }
  }

  public void init(File f) {
    try {
      if(builder.getPassword() != null) {
        // Based on: https://poi.apache.org/encryption.html
        POIFSFileSystem poifs = new POIFSFileSystem(f);
        EncryptionInfo info = new EncryptionInfo(poifs);
        Decryptor d = Decryptor.getInstance(info);
        d.verifyPassword(builder.getPassword());
        pkg = OPCPackage.open(d.getDataStream(poifs));
      } else {
        pkg = OPCPackage.open(f);
      }

      XSSFReader reader = new XSSFReader(pkg);
      if(builder.getSstCacheSize() > 0) {
        sstCache = Files.createTempFile("", "").toFile();
        log.debug("Created sst cache file [" + sstCache.getAbsolutePath() + "]");
        sst = BufferedStringsTable.getSharedStringsTable(sstCache, builder.getSstCacheSize(), pkg);
      } else {
        sst = reader.getSharedStringsTable();
      }

      StylesTable styles = reader.getStylesTable();
      NodeList workbookPr = searchForNodeList(document(reader.getWorkbookData()), "/workbook/workbookPr");
      if(workbookPr.getLength() == 1) {
        final Node date1904 = workbookPr.item(0).getAttributes().getNamedItem("date1904");
        if(date1904 != null) {
          use1904Dates = ("1".equals(date1904.getTextContent()));
        }
      }

      loadSheets(reader, sst, styles, builder.getRowCacheSize());
    } catch(IOException e) {
      throw new OpenException("Failed to open file", e);
    } catch(OpenXML4JException e) {
      throw new ReadException("Unable to read workbook", e);
    } catch(XMLStreamException e) {
      throw new ReadException("Unable to read workbook", e);
    } catch(GeneralSecurityException e) {
      throw new ReadException("Unable to read workbook - Decryption failed", e);
    }
  }

  void loadSheets(XSSFReader reader, SharedStringsTable sst, StylesTable stylesTable, int rowCacheSize) throws IOException, InvalidFormatException,
      XMLStreamException {
    lookupSheetNames(reader);
    Iterator<InputStream> iter = reader.getSheetsData();
    int i = 0;
    while(iter.hasNext()) {
      XMLEventReader parser = XMLInputFactory.newInstance().createXMLEventReader(iter.next());
      sheets.add(new StreamingSheet(sheetProperties.get(i++).get("name"), new StreamingSheetReader(sst, stylesTable, parser, use1904Dates, rowCacheSize)));
    }
  }

  void lookupSheetNames(XSSFReader reader) throws IOException, InvalidFormatException {
    sheetProperties.clear();
    NodeList nl = searchForNodeList(document(reader.getWorkbookData()), "/workbook/sheets/sheet");
    for(int i = 0; i < nl.getLength(); i++) {
      Map<String, String> props = new HashMap<String, String>();
      props.put("name", nl.item(i).getAttributes().getNamedItem("name").getTextContent());

      Node state = nl.item(i).getAttributes().getNamedItem("state");
      props.put("state", state == null ? "visible" : state.getTextContent());
      sheetProperties.add(props);
    }
  }

  List<? extends Sheet> getSheets() {
    return sheets;
  }

  public List<Map<String, String>> getSheetProperties() {
    return sheetProperties;
  }

  @Override
  public Iterator<Sheet> iterator() {
    return new StreamingSheetIterator(sheets.iterator());
  }

  @Override
  public void close() {
    try {
      for(StreamingSheet sheet : sheets) {
        sheet.getReader().close();
      }
      pkg.revert();
    } finally {
      if(tmp != null) {
        log.debug("Deleting tmp file [" + tmp.getAbsolutePath() + "]");
        tmp.delete();
      }
      if(sst instanceof BufferedStringsTable) {
        log.debug("Deleting sst cache file [" + tmp.getAbsolutePath() + "]");
        ((BufferedStringsTable) sst).close();
        sstCache.delete();
      }
    }
  }

  static File writeInputStreamToFile(InputStream is, int bufferSize) throws IOException {
    File f = Files.createTempFile("tmp-", ".xlsx").toFile();
    FileOutputStream fos = null;
    try {
      fos = new FileOutputStream(f);
      int read;
      byte[] bytes = new byte[bufferSize];
      while((read = is.read(bytes)) != -1) {
        fos.write(bytes, 0, read);
      }
      is.close();
      fos.close();
    } catch (IOException e) {
    	is.close();
    	fos.close();
    }
    return f;
  }

  static class StreamingSheetIterator implements Iterator<Sheet> {
    private final Iterator<StreamingSheet> iterator;

    public StreamingSheetIterator(Iterator<StreamingSheet> iterator) {
      this.iterator = iterator;
    }

    @Override
    public boolean hasNext() {
      return iterator.hasNext();
    }

    @Override
    public Sheet next() {
      return iterator.next();
    }

    @Override
    public void remove() {
      throw new RuntimeException("NotSupported");
    }
  }
}
