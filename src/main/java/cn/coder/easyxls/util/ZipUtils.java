package cn.coder.easyxls.util;

import java.io.IOException;
import java.io.InputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ZipUtils {
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
}
