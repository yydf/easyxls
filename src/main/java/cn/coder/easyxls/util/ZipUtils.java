package cn.coder.easyxls.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public final class ZipUtils {
	public static void putEntry(ZipOutputStream zipStream, String entry, String resource) throws IOException {
		InputStream is = ZipUtils.class.getClassLoader().getResourceAsStream(resource);
		byte[] data = new byte[1024];
		zipStream.putNextEntry(new ZipEntry(entry));
		int length = -1;
		while ((length = is.read(data)) != -1) {
			zipStream.write(data, 0, length);
		}
		zipStream.closeEntry();
	}

	public static void putStreamEntry(ZipOutputStream zipStream, String entry, byte[] data) throws IOException {
		zipStream.putNextEntry(new ZipEntry(entry));
		zipStream.write(data, 0, data.length);
		zipStream.closeEntry();
	}

	public static void close(ZipOutputStream zipStream, OutputStream outputStream) {
		if (zipStream != null) {
			try {
				zipStream.close();
			} catch (IOException e) {
			}
		}
		if (outputStream != null) {
			try {
				outputStream.close();
			} catch (IOException e) {
			}
		}
	}

	public static ZipOutputStream createZip(OutputStream outputStream, byte[] workBook, byte[] workBookRels, byte[] app) throws IOException {
		ZipOutputStream zipStream =  new ZipOutputStream(outputStream);
		putEntry(zipStream, "_rels/.rels", "_rels.rels.xml");
		putEntry(zipStream, "docProps/core.xml", "core.xml");
		putEntry(zipStream, "xl/styles.xml", "styles.xml");
		putEntry(zipStream, "xl/theme/theme1.xml", "theme1.xml");
		putEntry(zipStream, "[Content_Types].xml", "Content_Types.xml");

		putStreamEntry(zipStream, "docProps/app.xml", app);
		putStreamEntry(zipStream, "xl/workbook.xml", workBook);
		putStreamEntry(zipStream, "xl/_rels/workbook.xml.rels", workBookRels);
		return zipStream;
	}
}
