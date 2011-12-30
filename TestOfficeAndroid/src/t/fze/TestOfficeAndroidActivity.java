package t.fze;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Enumeration;
import java.util.ResourceBundle;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import android.app.Activity;
import android.os.Bundle;
import android.util.Log;
import dalvik.system.DexClassLoader;

public class TestOfficeAndroidActivity extends Activity {
	static final int BUF_SIZE = 8 * 1024;

    /** Called when the activity is first created. */
	private static final String ONARY_DEX_NAME = "onary_dex.jar";	
	private static final String SECONDARY_DEX_NAME = "secondary_dex.jar";
	private static final String TERTIARY_DEX_NAME = "tertiary_dex.jar";
	private static final String FOURTIARY_DEX_NAME = "fourtiary_dex.jar";
	private static final String FIFTHARY_DEX_NAME = "fifthary_dex.jar";
	private static final String SIXTHARY_DEX_NAME = "sixthary_dex.jar";
	private static final String SEVENARY_DEX_NAME = "sevenary_dex.jar";
	private static final String EIGHTARY_DEX_NAME = "eightary_dex.jar";
	private static final String NINEARY_DEX_NAME = "nineary_dex.jar";
	
	private static final ResourceBundle RESOURCE_BUNDLE = new ResourceBundle(){
			
			@Override
			public Enumeration<String> getKeys() {
				ArrayList<String> s = new ArrayList<String>();
				
				return Collections.enumeration(s);
			}

			@Override
			protected Object handleGetObject(String key) {
				// TODO Auto-generated method stub
				return null;
			}
    	};
    
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.main);

        //if the demo xlsx file isn't already copied, copy it from the assets directory to the sdcard
        if(!new File("/sdcard/office/SampleSS.xlsx").exists())
        {
        	copyFilesToPrivate("SampleSS.xlsx");
        }
        copyXSBToSDCard();
        if(!new File("/sdcard/office/style/excelStyle.css").exists())
        {
        	new File("/sdcard/office/style/").mkdirs();
    		File dexInternalStoragePath = new File("/sdcard/office/style/"+"excelStyle.css");
            BufferedInputStream bis = null;
            OutputStream dexWriter = null;

            
            try {
                bis = new BufferedInputStream(getAssets().open("excelStyle.css"));
                dexWriter = new BufferedOutputStream(new FileOutputStream(dexInternalStoragePath));
                byte[] buf = new byte[BUF_SIZE];
                int len;
                while((len = bis.read(buf, 0, BUF_SIZE)) > 0) {
                    dexWriter.write(buf, 0, len);
                }
                dexWriter.close();
                bis.close();
                
            } catch (Exception e) {e.printStackTrace();}
        }
        
        //loading the class that converts xlsx to html
        Class cl0 = loadClass( "org.apache.poi.ss.examples.html.ToHtml");
    
        try {
        	
        	//not really needed, just to make sure we got the class
			Object o = cl0.newInstance();
			//get the method "create" from ToHtml and invoke it with our demo XLSX file
//			cl0.getMethod("create",String.class, java.lang.Appendable.class)
//				.invoke(null, "/sdcard/office/SampleSS.xlsx", 
//						new java.io.PrintWriter(new java.io.FileWriter("/sdcard/office/SampleSS.html")));
//			cl0.getMethod("main", new Class[]{String[].class}).invoke(o, new String[]{"/sdcard/office/SampleSS.xlsx"});
			
			 //rechercher la méthode "main"
	        Class[] paramTypes = new Class[] { String[].class };
	        Method main = cl0.getMethod("main", paramTypes);
	        
	        //invoquer la méthode avec deux arguments
	        //(merci Java pour le cast obligatoire et non intuitif)
	        main.invoke(o, (Object) new String[] {"/sdcard/office/SampleSS.xlsx", "/sdcard/office/SampleSS.html"});
		} catch (InstantiationException e) {
			TestOfficeAndroidActivity.this.getClassLoader().clearAssertionStatus();
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			TestOfficeAndroidActivity.this.getClassLoader().clearAssertionStatus();
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			TestOfficeAndroidActivity.this.getClassLoader().clearAssertionStatus();
		}
        
        
    }
    /**
     * copies 
     */
	private void copyXSBToSDCard() {
		try {
			File dirXSB = new File("/sdcard/office/xsb/");
			//if xsb files are already on the SD card, no need
			if(!dirXSB.exists() || (dirXSB.exists() && dirXSB.list().length < 2528))
			{
				dirXSB.mkdirs();
				Log.d("OfficeInstallUnzipXSB", "part 1...");
				extractTo("xsb1.zip", dirXSB.getAbsolutePath()+"/");
				Log.d("OfficeInstallUnzipXSB", "part 1... done");
				Log.d("OfficeInstallUnzipXSB", "part 2...");
				extractTo("xsb2.zip", dirXSB.getAbsolutePath()+"/");
				Log.d("OfficeInstallUnzipXSB", "part 2... done");
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	  /**
     * Extraire un fichier
     * @param archive fichier zip
     * @param file nom du fichier à extraire
     * @param destPath répertoir de destination du fichier
     * @throws FileNotFoundException
     * @throws IOException
     */
    public void extractTo(String archive, String destPath) throws FileNotFoundException, IOException {
        ZipInputStream zipInputStream = null;
        ZipEntry zipEntry = null;
        byte[] buffer = new byte[2048];
        
        zipInputStream = new ZipInputStream(getAssets().open(archive));
        zipEntry = zipInputStream.getNextEntry();
        while (zipEntry != null) {
            
        	FileOutputStream fileoutputstream = new FileOutputStream(destPath + zipEntry.getName());
                int n;

                while ((n = zipInputStream.read(buffer, 0, 2048)) > -1) {
                    fileoutputstream.write(buffer, 0, n);
                }

                fileoutputstream.close();
                zipInputStream.closeEntry();

            zipEntry = zipInputStream.getNextEntry();
        }
    }
    private Class loadClass(String className) {
    	
    	//copy all the DEX files to the SDcard
    	File onary = copyFilesToPrivate(ONARY_DEX_NAME);
    	File secondary = copyFilesToPrivate(SECONDARY_DEX_NAME);
    	File tertiary = copyFilesToPrivate(TERTIARY_DEX_NAME);
    	File fourtiary = copyFilesToPrivate(FOURTIARY_DEX_NAME);
    	File fifthiary = copyFilesToPrivate(FIFTHARY_DEX_NAME);
    	File sixthiary = copyFilesToPrivate(SIXTHARY_DEX_NAME);
    	File sevenary = copyFilesToPrivate(SEVENARY_DEX_NAME);
    	File eightary = copyFilesToPrivate(EIGHTARY_DEX_NAME);
    	File nineary= copyFilesToPrivate(NINEARY_DEX_NAME);
        
        // Internal storage where the DexClassLoader writes the optimized dex file to
    	new File("/sdcard/office/opti").mkdirs();
        final File optimizedDexOutputPath = new File("/sdcard/office/opti");//getDir("outdex", Context.MODE_PRIVATE);

        //the dexloader is given all the jars that we need and loads them into the activity's classloader
		DexClassLoader cl = new DexClassLoader(onary.getAbsolutePath()+":"+secondary.getAbsolutePath()
												+":"+tertiary.getAbsolutePath()+":"+fourtiary.getAbsolutePath()
												+":"+fifthiary.getAbsolutePath()+":"+sixthiary.getAbsolutePath()
												+":"+sevenary.getAbsolutePath()+":"+eightary.getAbsolutePath()
												+":"+nineary.getAbsolutePath(),
                                               optimizedDexOutputPath.getAbsolutePath(),
                                               null,
                                               TestOfficeAndroidActivity.this.getClassLoader());
        Class libProviderClazz = null;
        try {
            // Load the library.
            libProviderClazz = cl.loadClass(className);
//            Class<?> handler = Class.forName(className, true, cl);
        }
        catch(Exception e)
        {e.printStackTrace();}
            
		return libProviderClazz;
	}
	private File copyFilesToPrivate(String dexName) {
		new File("/sdcard/office").mkdirs();
		File dexInternalStoragePath = new File("/sdcard/office/"+dexName);
        BufferedInputStream bis = null;
        OutputStream dexWriter = null;

        
        try {
            bis = new BufferedInputStream(getAssets().open(dexName));
            dexWriter = new BufferedOutputStream(new FileOutputStream(dexInternalStoragePath));
            byte[] buf = new byte[BUF_SIZE];
            int len;
            while((len = bis.read(buf, 0, BUF_SIZE)) > 0) {
                dexWriter.write(buf, 0, len);
            }
            dexWriter.close();
            bis.close();
            
        } catch (Exception e) {e.printStackTrace();
        return null;}
        Log.d("OfficeDEXCopy", "copied "+dexName+" to /sdcard/office/");
        return dexInternalStoragePath;
	}
	
	

//  files = new String[]{"/sdcard/download/SampleSS.xlsx", "/sdcard/download/SampleSSx.html" };
//  String[] files = new String[]{"/sdcard/download/SampleDoc.doc", "/sdcard/download/SampleDoc.html" };
//  WordToHtmlConverter.main(files);
//  
//  files = new String[]{"/sdcard/download/SampleSS.xls", "/sdcard/download/SampleSS.html" };
//  ExcelToHtmlConverter.main(files);
//	Method main =cl0.getMethod("main", new Class[]{String[].class});//String[].class
//	main.invoke(null, files);
	
//	ToHtml.main(files);

  
//  Class cl = loadClass(SECONDARY_DEX_NAME, "org.apache.xmlbeans.XmlOptions");
//  Class cl1 = loadClass(TERTIARY_DEX_NAME, "org.apache.poi.ss.examples.html.ToHtml");
  
//  PathClassLoader myClassLoader =
//      new dalvik.system.PathClassLoader(
//      		new File(getDir("dex", Context.MODE_PRIVATE), SECONDARY_DEX_NAME).getAbsolutePath(),
//              ClassLoader.getSystemClassLoader());
//  try {
//		Class<?> handler = Class.forName("org.apache.xmlbeans.XmlOptions", true, myClassLoader);
//	} catch (ClassNotFoundException e1) {
//		// TODO Auto-generated catch block
//		e1.printStackTrace();
//	}
}