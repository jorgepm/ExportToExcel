
package com.sicop.common.utilities;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.zip.*;


public class FilesToZip {
	static final int BUFFER = 1024;

        public void Zippear(String pFile, String pZipFile) throws Exception {
		// objetos en memoria
		FileInputStream fis = null;
		FileOutputStream fos = null;
		ZipOutputStream zipos = null;

		// buffer
		byte[] buffer = new byte[BUFFER];
		try {
                        BufferedInputStream origen=null;
			// fichero contenedor del zip
			fos = new FileOutputStream(pZipFile);
			// fichero a comprimir
			fis = new FileInputStream("c:\\exporta\\" + pFile);
			// fichero comprimido
			zipos = new ZipOutputStream(fos);
			ZipEntry zipEntry = new ZipEntry(pFile);
			zipos.putNextEntry(zipEntry);
			int len = 0;
			// zippear
			while ((len = fis.read(buffer, 0, BUFFER)) != -1)
				zipos.write(buffer, 0, len);
			// volcar la memoria al disco
			zipos.flush();
		} catch (Exception e) {
			throw e;
		} finally {
			// cerramos los files
			zipos.close();
			fis.close();
			fos.close();
		}

	}

}
