import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Iterator;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.LibraryLoader;
import com.jacob.com.Variant;

/**
 * @author vicky.thakor
 * @date 28th December, 2013
 * http://sourceforge.net/projects/jacob-project/
 * First Program to understand how to use JACOB library
 */
public class JACOBGettingStarted {

    public static void main(String[] args) {
    	
    	
    	
        /**
         * `System.getProperty("os.arch")`
         * It'll tell us on which platform Java Program is executing. Based on that we'll load respective DLL file.
         * Placed under same folder of program file(.java/.class).
         */
        String libFile = System.getProperty("os.arch").equals("amd64") ? "jacob-1.17-x64.dll" : "jacob-1.17-x86.dll";
        try {
            /* Read DLL file*/
            //InputStream inputStream = jacobExample.class.getResourceAsStream(libFile);
        	InputStream inputStream = new FileInputStream("C:\\Users\\Administrator\\Downloads\\jacob-1.18\\jacob-1.18-x64.dll");
            /**
             *  Step 1: Create temporary file under <%user.home%>\AppData\Local\Temp\jacob.dll 
             *  Step 2: Write contents of `inputStream` to that temporary file.
             */
            File temporaryDll = File.createTempFile("jacob", ".dll");
            FileOutputStream outputStream = new FileOutputStream(temporaryDll);
            byte[] array = new byte[8192];
            for (int i = inputStream.read(array); i != -1; i = inputStream.read(array)) {
                outputStream.write(array, 0, i);
            }
            outputStream.close();
            /**
             * `System.setProperty(LibraryLoader.JACOB_DLL_PATH, temporaryDll.getAbsolutePath());`
             * Set System property same like setting java home path in system.
             * 
             * `LibraryLoader.loadJacobLibrary();`
             * Load JACOB library in current System.
             */
            System.setProperty(LibraryLoader.JACOB_DLL_PATH, temporaryDll.getAbsolutePath());
            LibraryLoader.loadJacobLibrary();
            
            ComThread.InitSTA();

            /**
             * Create ActiveXComponent using CLSID. You can also use program id here.
             * Next line(commented line/compProgramID) shows you how you can create ActiveXComponent using ProgramID.
             */
            //ActiveXComponent compCLSID = new ActiveXComponent("clsid:{00024500-0000-0000-C000-000000000046}");
            ActiveXComponent ocx = new ActiveXComponent("KHOPENAPI.KHOpenAPICtrl.1");
            
            Variant result = ocx.getProperty("GetConnectState");
            System.out.println("ocx getProperty : " + result.toString());
            
            System.out.println("The Library been loaded, and an activeX component been created");
            
            /**
             * This is function/method of Microsoft Excel to use it with COM bridge.
             * Excel methods and its use can be found on
             * http://msdn.microsoft.com/en-us/library/bb179167(v=office.12).aspx
             * 
             * Understand code:
             * 1. Make Excel visible
             * 2. Get workbook of excel object.
             * 3. Open 1test.xls1 file in excel
             */
            /*Dispatch.put(compCLSID, "Visible", new Variant(true));
            Dispatch workbook = compCLSID.getProperty("Workbooks").toDispatch();
            Dispatch.call(workbook, "Open", new Variant("C:\\Users\\Administrator\\Downloads\\lrngShrgGrid20160906.xlsx"));*/

            /* Temporary file will be removed after terminating-closing-ending the application-program */
            temporaryDll.deleteOnExit();
            
            ComThread.Release();
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        
    }
}