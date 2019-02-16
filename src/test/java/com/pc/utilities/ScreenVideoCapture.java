package com.pc.utilities;

import static org.monte.media.FormatKeys.EncodingKey;
import static org.monte.media.FormatKeys.FrameRateKey;
import static org.monte.media.FormatKeys.KeyFrameIntervalKey;
import static org.monte.media.FormatKeys.MIME_AVI;
import static org.monte.media.FormatKeys.MediaTypeKey;
import static org.monte.media.FormatKeys.MimeTypeKey;
import static org.monte.media.VideoFormatKeys.CompressorNameKey;
import static org.monte.media.VideoFormatKeys.DepthKey;
import static org.monte.media.VideoFormatKeys.ENCODING_AVI_TECHSMITH_SCREEN_CAPTURE;
import static org.monte.media.VideoFormatKeys.QualityKey;

import java.awt.AWTException;
import java.awt.GraphicsConfiguration;
import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.Comparator;

import org.apache.commons.io.FileUtils;
import org.monte.media.Format;
import org.monte.media.FormatKeys.MediaType;
import org.monte.media.math.Rational;
import org.monte.screenrecorder.ScreenRecorder;
public class ScreenVideoCapture  {

	private static ScreenRecorder screenRecorder;
	
	public ScreenVideoCapture ()
	{
		
	}

/**
 * 
 * TODO Method to start the video capturing using monte screen recorder.
 *
 * @throws IOException
 * @throws AWTException
 */
public static void startVideoCapture() throws IOException, AWTException
{
	try
	{
		if(HTML.properties.getProperty("VideoCapture").equalsIgnoreCase("true"))
		{
			
			GraphicsDevice gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
			int width = gd.getDisplayMode().getWidth();
			int height = gd.getDisplayMode().getHeight();
			java.awt.Rectangle captureArea=new java.awt.Rectangle(width,height);
	GraphicsConfiguration gc = GraphicsEnvironment
            .getLocalGraphicsEnvironment()
            .getDefaultScreenDevice()
            .getDefaultConfiguration();
	//String path=System.getProperty("user.dir");
	//File directory = new File (".");
	
	File directory = new File (".");
	String sConfigfilespath = directory.getCanonicalPath()+"\\Reports\\Videos\\";
	
	File mediaFolder=new File(sConfigfilespath);
	if(!mediaFolder.exists())
	{
		mediaFolder.mkdirs();
	}       
			screenRecorder = new ScreenRecorder(gc,captureArea,
			       new Format(MediaTypeKey, MediaType.FILE, MimeTypeKey, MIME_AVI),
			       new Format(MediaTypeKey, MediaType.VIDEO, EncodingKey, ENCODING_AVI_TECHSMITH_SCREEN_CAPTURE,
			            CompressorNameKey, ENCODING_AVI_TECHSMITH_SCREEN_CAPTURE,
			            DepthKey, 24, FrameRateKey, Rational.valueOf(15),
			            QualityKey, 1.0f,
			            KeyFrameIntervalKey, 15 * 60),
			       new Format(MediaTypeKey, MediaType.VIDEO, EncodingKey, "black",
			            FrameRateKey, Rational.valueOf(30)),
			      null,mediaFolder);
		
			
			screenRecorder.start();
		
		}
	} 
	catch (Exception exception)
	{		
		exception.printStackTrace();
	}
		
}

/**
 * 
 * TODO Method to stop the video recording.
 *
 * @throws IOException
 * @throws AWTException
 */
public static void stopVideoCapture(String sTestName) throws IOException, AWTException
{
	try
	{
		if(HTML.properties.getProperty("VideoCapture").equalsIgnoreCase("true"))
		{
			screenRecorder.stop();
			//capturedVideoFileName(sTestName);
		}
	} catch (Exception exception)
	{
		
		exception.printStackTrace();
	}
}

/**
 * 
 * TODO Method to rename the captured video with respective test case name.
 * @throws IOException 
 *
 */
public static void capturedVideoFileName(String sTestName) throws IOException
{
	File directory = new File (".");
	String sConfigfilespath = directory.getCanonicalPath()+"\\Reports\\Videos\\";
	
	//String path=System.getProperty("user.dir");
	File mediaFolder=new File(sConfigfilespath);
	
	File[] files = mediaFolder.listFiles();
	File mediaFolder2=new File(sConfigfilespath + sTestName);
	if(!mediaFolder2.exists())
	{
		mediaFolder.mkdirs();
	}
    Arrays.sort( files, new Comparator<Object>()
    {
    public int compare(Object o1, Object o2) {
    	//return new Long(((File)o1).lastModified()).compareTo(new Long(((File)o2).lastModified())); // for ascending order
    	return -1*(new Long(((File)o1).lastModified()).compareTo(new Long(((File)o2).lastModified()))); //for decending order 
    }
    });
    /*for(int i=0;i<files.length;i++)
    {
        System.out.println("i value = "+i+" file name"+files[i].getName());
    }*/
    File oldfile = new File(sConfigfilespath + files[0].getName());
    FileUtils.moveFileToDirectory(oldfile, mediaFolder2, true);
    File newfile = new File(sConfigfilespath + "sTestName"+files[0].getName());
    File newName= new File(sConfigfilespath + "sTestName\\sTestName"+".avi");
   int count=1;
    while(newName.exists())
   {
	   newName=new File(sConfigfilespath + "\\sTestName\\sTestName"+count+".avi");
   count++;
   }
    newfile.renameTo(newName);
    
    /*if(oldfile.renameTo(newfile))
    {
    	System.out.println("latest file has been renamed successfully");
    }
    else
    {
    	System.out.println("renaming of the file has not been successfully"); 
    }*/
}
}

