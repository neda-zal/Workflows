using LoginPI.Engine.ScriptBase;


public class KnowledgeWorker : ScriptBase
{
    void Execute() 
    {
       
     
       
      // Get temp directory location
      var TempDirectory = GetEnvironmentVariable("temp", continueOnError:true);
      var TempDirectoryLE = TempDirectory+"\\LoginEnterprise";
      ShellExecute("cmd /c mkdir -p "+TempDirectoryLE, waitForProcessEnd: true, timeout: 15, forceKillOnExit: false);
      
      // Get HomeDirectory location
      var HomeDirectory = GetEnvironmentVariable("homedrive", continueOnError:true)+GetEnvironmentVariable("homepath", continueOnError:true);
      var HomeDirectoryLE = HomeDirectory+"\\Documents\\LoginEnterprise\\";
      ShellExecute("cmd /c mkdir -p "+HomeDirectoryLE, waitForProcessEnd: true, timeout: 15, forceKillOnExit: false);
      
      // Get outlook pst and prf file ready + registry key
      CopyFile(KnownFiles.OutlookConfiguration, TempDirectory+"\\LoginPI\\Outlook.prf",continueOnError: true,overwrite:true);
      CopyFile(KnownFiles.OutlookData, TempDirectory+"\\LoginPI\\Outlook.pst",continueOnError: true,overwrite:true);
      ShellExecute("cmd /c reg add HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Outlook\\Setup /v importPRF /t REG_SZ /d "+TempDirectory+"\\LoginPI\\Outlook.prf /f", waitForProcessEnd: false, timeout: 15, forceKillOnExit: false);
      
      // Set adobe reader to open seperate windows (no-tabs)
      ShellExecute("cmd /c reg add \"HKCU\\Software\\Adobe\\Acrobat Reader\\DC\\AVGeneral\" /v bSDIMode /t REG_DWORD /d 1 /f", waitForProcessEnd: false, timeout: 15, forceKillOnExit: false);  
      
      // disable powerpoint prompt for hardware acceleration 
      ShellExecute("cmd /c reg add \"HKCU\\Software\\Microsoft\\Office\\16.0\\PowerPoint\\Options\" /v DisableHardwareNotification /t REG_DWORD /d 1 /f", waitForProcessEnd: false, timeout: 15, forceKillOnExit: false);  

      Log(message:"Start outlook and maximize"); 
      // Start outlook and maximize
      StartTimer(name:"AppStart_Outlook");
      var Outlook =  ShellExecute("outlook.exe", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
      var OutlookWindow = FindWindow(title:"Inbox*", timeout:10);
      StopTimer(name:"AppStart_Outlook");
      OutlookWindow.Focus();
      OutlookWindow.Maximize();
      
       Log(message:"Start browser + maximize + read website"); 
      //Start browser + maximize + read website
       StartTimer(name:"AppStart_Browser");
       StartBrowser(useInPrivateBrowsing:false,expectedUrl:"https://www.bbc.com/",timeout:30,continueOnError:true);
       StopTimer(name:"AppStart_Browser");
       var EdgeWindow = FindWindow(title:"BBC - Homepage*", timeout:10);
       Wait(2);
       EdgeWindow.Maximize(); 
       Wait(2);
       EdgeWindow.Type("{HOME}");
       Wait(2); 
       EdgeWindow.Type("{DOWN}{DOWN}");
       Wait(2); 
       EdgeWindow.Type("{DOWN}{DOWN}");
       Wait(2); 
       EdgeWindow.Type("{DOWN}{DOWN}");
       Wait(2);  
       
       
        Log(message:"OutlookAttach1");     
        //OUTLOOK create message (1) with attachment
        // Get files ready (attach to email)
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginPI\\OutLookAttach1.docx",continueOnError: true,overwrite:true);
        Wait(2);
        StartTimer(name:"NewMsgWithAttachment1");
        var OutlookNewMSG1 =  ShellExecute("cmd.exe /c start outlook.exe /a "+TempDirectory+"\\LoginPI\\OutLookAttach1.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false, continueOnError: true);
        var OutlookWindowNewMSG1 = FindWindow(title:"*HTML*", timeout:10);
        StopTimer(name:"NewMsgWithAttachment1");
        Wait(2); 
        OutlookWindowNewMSG1.Focus();
        Wait(2); 
        OutlookWindowNewMSG1.Maximize();
        Wait(2); 
        OutlookWindowNewMSG1.Type("John Dummy");
        Wait(2); 
        OutlookWindowNewMSG1.Type("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}");
        Wait(2); 
        OutlookWindowNewMSG1.Type("The quick brown fox jumps over the lazy dog.{TAB}");
        KeyDown(KeyCode.CTRL);
        OutlookWindowNewMSG1.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookWindowNewMSG1.Close();
        // Clean up
        RemoveFile(path: TempDirectory+"\\LoginPI\\OutLookAttach.docx");
       
Log(message:"Outlook back to main window");        
        //Outlook, back to main window
        OutlookWindow.Focus();
        Wait(2); 
        OutlookWindow.Maximize();
        Wait(2); 
        OutlookWindow.Type("{down}{down}{down}");
        Wait(5);
        OutlookWindow.Type("{down}{down}{down}{down}");
        Wait(5);
        
Log(message:"Outlook new message 2");        
        // Create new message (2)
        KeyDown(KeyCode.CTRL);
        OutlookWindow.Type("n");
        KeyUp(KeyCode.CTRL); 
        var OutlookNewMessageWindow2 = FindWindow(title:"*HTML*", timeout:10, continueOnError: true);
        Wait(2); 
        OutlookNewMessageWindow2.Maximize();
        Wait(2); 
        OutlookNewMessageWindow2.Type("Jane Dummy");
        Wait(2);
        OutlookNewMessageWindow2.Type("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}");
        Wait(2);
        OutlookNewMessageWindow2.Type("Lets grab lunch at the Italian place today.{TAB}");
        KeyDown(KeyCode.CTRL);
        OutlookNewMessageWindow2.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookNewMessageWindow2.Close();
      
      
Log(message:"Outlook back to main window");       
        // Go back to main program and write a new message (3)
       Wait(2); 
       OutlookWindow.Focus();
        OutlookWindow.Maximize();
        Wait(2); 
        KeyDown(KeyCode.CTRL);
        OutlookWindow.Type("n");
        KeyUp(KeyCode.CTRL);

Log(message:"Outlook new message 3"); 
        var OutlookNewMessageWindow3= FindWindow(title:"*HTML*", timeout:10, continueOnError: true);
        Wait(2); 
        OutlookNewMessageWindow3.Maximize();
        Wait(2); 
        OutlookNewMessageWindow3.Type("Joe Sixpack");
        OutlookNewMessageWindow3.Type("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}");
        OutlookNewMessageWindow3.Type("Lets grab lunch at the Dutch place today.{TAB}");
        KeyDown(KeyCode.CTRL);
        OutlookNewMessageWindow3.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookNewMessageWindow3.Close();
      
      
        Wait(seconds:20, showOnScreen:true, onScreenText:"Getting a tripleshot cappucino");

Log(message:"Opening the verge website"); 
        // Open second browser windows to theverge, cannot use startbrowser for the 2nd time.
        try {
        ShellExecute("microsoft-edge:https://www.theverge.com/", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        } catch {}
        var EdgeWindow2 = FindWindow(title : "The Verge*");
        EdgeWindow2.Focus();
    Wait(2);
       EdgeWindow2.Maximize(); 
       EdgeWindow2.Type("{HOME}");
       Wait(2); 
       EdgeWindow2.Type("{DOWN}{DOWN}");
       Wait(2); 
       EdgeWindow2.Type("{DOWN}{DOWN}");
       Wait(2); 
       EdgeWindow2.Type("{DOWN}{DOWN}");
       Wait(2);  
       
       
Log(message:"Outlook back to main window");     
        // Go back to main program and write a new message (4)
        OutlookWindow.Focus();
        Wait(2); 
        KeyDown(KeyCode.CTRL);
        OutlookWindow.Type("n");
        KeyUp(KeyCode.CTRL);
Log(message:"Outlook message 4");         
        var OutlookNewMessageWindow4= FindWindow(title:"*HTML*", timeout:10, continueOnError: true);
        Wait(2); 
        OutlookNewMessageWindow4.Maximize();
        Wait(2); 
        OutlookNewMessageWindow4.Type("Joe Sixpack");
        Wait(2); 
        OutlookNewMessageWindow4.Type("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}");
        Wait(2); 
        OutlookNewMessageWindow4.Type("Lets grab lunch at the Dutch place today.{TAB}");
        KeyDown(KeyCode.CTRL);
        OutlookNewMessageWindow4.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookNewMessageWindow4.Close();
        OutlookWindow.Focus();
  
  
  Log(message:"Winword1"); 
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginEnterprise\\WinWord1.docx",continueOnError: true,overwrite:true);
        ShellExecute("winword.exe " + TempDirectory+"\\LoginEnterprise\\WinWord1.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        
        var WinWord1 = FindWindow(className : "Win32 Window:OpusApp", title : "WinWord1*", processName : "WINWORD");
        WinWord1.Focus().Maximize();
        Wait(2);
        ReadText(20);
        TypeText(10); 
        WinWord1.Type("{CTRL+S}");
        TypeText(5);
        WinWord1.Type("{CTRL+S}");
        TypeText(10);
        WinWord1.Type("{CTRL+S}");
        
        
        Wait(seconds:20, showOnScreen:true, onScreenText:"Talking to a colleague about some awesome stuff");
       
       Log(message:"Winword2"); 
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginEnterprise\\WinWord2.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"Winword2");
        ShellExecute("winword.exe " + TempDirectory+"\\LoginEnterprise\\WinWord2.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        var WinWord2 = FindWindow(className : "Win32 Window:OpusApp", title : "WinWord2*", processName : "WINWORD");
        StopTimer(name:"WinWord2");
        WinWord2.Focus();
        Wait(2);
        WinWord2.Maximize();
        Wait(2);
        ReadText(20);
        TypeText(10); 
        WinWord2.Type("{CTRL+S}");
        TypeText(5);
        WinWord2.Type("{CTRL+S}");
        TypeText(10);
        WinWord2.Type("{CTRL+S}");
  
  Log(message:"PDF1"); 
        CopyFile(KnownFiles.PdfFile, TempDirectory+"\\LoginEnterprise\\PDFDocument1.pdf",continueOnError: true,overwrite:true);
        StartTimer(name:"PDFDocument1");
        ShellExecute(TempDirectory+"\\LoginEnterprise\\PDFDocument1.pdf", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        var PDFReader1 = FindWindow(title : "PDFDocument1*");
        StopTimer(name:"PDFDocument1");
       Wait(2);
       PDFReader1.Focus();
       Wait(2);
       PDFReader1.Maximize();
        TypeText(10); 
        Wait(2);
        WinWord1.Close();
        RemoveFile(path: TempDirectory+"\\LoginEnterprise\\WinWord1.docx", continueOnError: true);
        
        ////////////////////////////////////  SEGMENT 2 /////////////////////////////////////////
        
      Log(message:"focues edge"); 
         EdgeWindow.Focus(); 
         Wait(2);
         EdgeWindow.Type("{home}");
         ReadText(30);
        
        Log(message:"outlook message 4"); 
        //OUTLOOK create message (4) with attachment
        // Get files ready (attach to email)
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginPI\\OutLookAttach4.docx",continueOnError: true,overwrite:true);
        Wait(2);
        StartTimer(name:"NewMsgWithAttachment4");
        var OutlookNewMSG4 =  ShellExecute("cmd.exe /c start outlook.exe /a "+TempDirectory+"\\LoginPI\\OutLookAttach1.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false, continueOnError: true);
        var OutlookWindowNewMSG4 = FindWindow(title:"*HTML*", timeout:10);
        StopTimer(name:"NewMsgWithAttachment4");
        Wait(2); 
        OutlookWindowNewMSG4.Focus();
        Wait(2); 
        OutlookWindowNewMSG4.Maximize();
        Wait(2); 
        OutlookWindowNewMSG4.Type("Deliver to");
        Wait(2); 
        OutlookWindowNewMSG4.Type("{TAB}{TAB}{TAB}{TAB}{TAB}");
        Wait(2); 
        OutlookWindowNewMSG4.Type("The quick brown fox jumps over the lazy dog. Message 4");
        Wait(2); 
        KeyDown(KeyCode.CTRL);
        OutlookWindowNewMSG4.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookWindowNewMSG4.Close();
        // Clean up
        RemoveFile(path: TempDirectory+"\\LoginPI\\OutLookAttach4.docx");  
  
        Log(message:"outlook back to main window"); 
        //Outlook, back to main window
        OutlookWindow.Focus();
        Wait(2);
        OutlookWindow.Maximize();
        Wait(2);
        OutlookWindow.Type("{down}{down}{down}");
        Wait(5);
        OutlookWindow.Type("{down}{down}{down}{down}");
        Wait(5);
        
        
        Log(message:"Outlook message 5"); 
          //OUTLOOK create message (5) with attachment
        // Get files ready (attach to email)
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginPI\\OutLookAttach5.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"NewMsgWithAttachment5");
        var OutlookNewMSG5 =  ShellExecute("cmd.exe /c start outlook.exe /a "+TempDirectory+"\\LoginPI\\OutLookAttach1.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false, continueOnError: true);
        var OutlookWindowNewMSG5 = FindWindow(title:"*HTML*", timeout:10);
        StopTimer(name:"NewMsgWithAttachment5");
        Wait(2);
        OutlookWindowNewMSG5.Focus();
        Wait(2);
        OutlookWindowNewMSG5.Maximize();
        Wait(2);
        OutlookWindowNewMSG5.Type("John Dummy");
        OutlookWindowNewMSG5.Type("{TAB}{TAB}{TAB}{TAB}{TAB}");
        OutlookWindowNewMSG5.Type("The quick brown fox jumps over the lazy dog. Message 5");
        KeyDown(KeyCode.CTRL);
        OutlookWindowNewMSG5.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookWindowNewMSG5.Close();
        // Clean up
        RemoveFile(path: TempDirectory+"\\LoginPI\\OutLookAttach5.docx");        
        
        Log(message:"Edge2"); 
        EdgeWindow2.Focus();
        Wait(2);
        EdgeWindow2.Type("{home}");
        ReadText(30);
        
        Log(message:"PDF22"); 
        CopyFile(KnownFiles.PdfFile, TempDirectory+"\\LoginEnterprise\\PDFDocument2.pdf",continueOnError: true,overwrite:true);
        StartTimer(name:"PDFDocument2");
        ShellExecute(TempDirectory+"\\LoginEnterprise\\PDFDocument2.pdf", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        var PDFReader2 = FindWindow(title : "PDFDocument2*");
        StopTimer(name:"PDFDocument2");
        Wait(2);
        PDFReader2.Focus();
        Wait(2);
        PDFReader2.Maximize();
        Wait(2);
        ReadText(30); 
    
        
        Log(message:"pdf3"); 
        CopyFile(KnownFiles.PdfFile, TempDirectory+"\\LoginEnterprise\\PDFDocument3.pdf",continueOnError: true,overwrite:true);
        StartTimer(name:"PDFDocument3");
        ShellExecute(TempDirectory+"\\LoginEnterprise\\PDFDocument3.pdf", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        var PDFReader3 = FindWindow(title : "PDFDocument3*");
        StopTimer(name:"PDFDocument3");
        Wait(2);
        PDFReader3.Focus();
        Wait(2);
        PDFReader3.Maximize();
        ReadText(40); 
        
        
        Log(message:"Winword3"); 
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginEnterprise\\WinWord3.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"Winword3");
        ShellExecute("winword.exe " + TempDirectory+"\\LoginEnterprise\\WinWord3.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        var WinWord3 = FindWindow(className : "Win32 Window:OpusApp", title : "WinWord3*", processName : "WINWORD");
        StopTimer(name:"WinWord3");
        Wait(2);
        WinWord3.Focus();
        Wait(2);
        WinWord3.Maximize();
        ReadText(20);
        TypeText(10); 
        WinWord3.Type("{CTRL+S}");
        TypeText(5);
        WinWord3.Type("{CTRL+S}");
        TypeText(10);
        WinWord3.Type("{CTRL+S}");
        
        Log(message:"Close pdf3"); 
        PDFReader3.Focus();
        PDFReader3.Close();
        
        
        Log(message:"PowerPoint1"); 
        CopyFile(KnownFiles.PowerPointPresentation, TempDirectory+"\\LoginEnterprise\\PowerPoint1.pptx",continueOnError: true,overwrite:true);
        Wait(2);
        StartTimer(name:"PowerPoint1");
        ShellExecute("powerpnt.exe " + TempDirectory+"\\LoginEnterprise\\PowerPoint1.pptx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        var PowerPoint1 = FindWindow(title : "PowerPoint1*");
        StopTimer(name:"PowerPoint1"); 
        Wait(2);
        PowerPoint1.Focus();
        Wait(2);
        PowerPoint1.Maximize();
        Wait(2);
        Type("{f5}");
        Wait(5);
        ReadText(10);
        PowerPoint1.Type("{esc}");
        
        WinWord3.Focus();
        WinWord3.Close();
        
        PDFReader1.Focus();
        ReadText(34);
        //PDFReader1.Close();
        
        EdgeWindow.Focus();
        
        ////////////////////////////////////  SEGMENT 3 /////////////////////////////////////////
        
        Wait(2);
        EdgeWindow.Type("{HOME}");
        ReadText(30);
        
        
        Log(message:"Outlook message 6"); 
         //OUTLOOK create message (6) with attachment
        // Get files ready (attach to email)
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginPI\\OutLookAttach6.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"NewMsgWithAttachment6");
        var OutlookNewMSG6 =  ShellExecute("cmd.exe /c start outlook.exe /a "+TempDirectory+"\\LoginPI\\OutLookAttach1.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false, continueOnError: true);
        var OutlookWindowNewMSG6 = FindWindow(title:"*HTML*", timeout:10);
        StopTimer(name:"NewMsgWithAttachment6");
        Wait(2);
        OutlookWindowNewMSG6.Focus();
        Wait(2);
        OutlookWindowNewMSG6.Maximize();
        Wait(3);
        OutlookWindowNewMSG6.Type("John Dummy");
        OutlookWindowNewMSG6.Type("{TAB}{TAB}{TAB}{TAB}{TAB}");
        OutlookWindowNewMSG6.Type("The quick brown fox jumps over the lazy dog. Message 6");
        KeyDown(KeyCode.CTRL);
        OutlookWindowNewMSG6.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookWindowNewMSG6.Close();
        // Clean up
        RemoveFile(path: TempDirectory+"\\LoginPI\\OutLookAttach6.docx");  
        
        
   Log(message:"Outlook back to main window");        
        //Outlook, back to main window
        OutlookWindow.Focus();
        Wait(2); 
        OutlookWindow.Maximize();
        Wait(2); 
        OutlookWindow.Type("{down}{down}{down}");
        Wait(5);
        OutlookWindow.Type("{down}{down}{down}{down}");
        Wait(5);
        
        
    Log(message:"Outlook message 7"); 
          //OUTLOOK create message (7) with attachment
        // Get files ready (attach to email)
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginPI\\OutLookAttach7.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"NewMsgWithAttachment7");
        var OutlookNewMSG7 =  ShellExecute("cmd.exe /c start outlook.exe /a "+TempDirectory+"\\LoginPI\\OutLookAttach7.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false, continueOnError: true);
        var OutlookWindowNewMSG7 = FindWindow(title:"*HTML*", timeout:10);
        StopTimer(name:"NewMsgWithAttachment7");
        Wait(2);
        OutlookWindowNewMSG7.Focus();
        Wait(2);
        OutlookWindowNewMSG7.Maximize();
        Wait(3);
        OutlookWindowNewMSG7.Type("Jane Dummy");
        OutlookWindowNewMSG7.Type("{TAB}{TAB}{TAB}{TAB}{TAB}");
        OutlookWindowNewMSG7.Type("The quick brown fox jumps over the lazy dog. Message 7");
        KeyDown(KeyCode.CTRL);
        OutlookWindowNewMSG7.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookWindowNewMSG7.Close();
        // Clean up
        RemoveFile(path: TempDirectory+"\\LoginPI\\OutLookAttach7.docx"); 
        
        EdgeWindow2.Focus();
        Wait(2);
        EdgeWindow2.Type("{HOME}");
        ReadText(60);
        
       
     Log(message:"Microsoft excel actions"); 
        CopyFile(KnownFiles.ExcelSheet, TempDirectory +"\\LoginPI\\Spreadsheet.xlsx",continueOnError: true,overwrite:true);
        ShellExecute("excel.exe " + TempDirectory + "\\LoginPI\\Spreadsheet.xlsx", waitForProcessEnd: false, forceKillOnExit: false);
        var ExcelWindow = FindWindow(className : "Win32 Window:XLMAIN", title : "*Excel", processName : "EXCEL");
        ExcelWindow.Maximize().Focus();
        ExcelWindow.Type("{F5}");
        Wait(2);
        FindWindow(className : "Win32 Window:bosa_sdm_XL9", title : "Go To", processName : "EXCEL").Focus();
        ExcelWindow.Type("A1{enter}{esc}{esc}");
        
        
        ExcelWindow.Maximize().Focus();
        ExcelWindow.Type("{CTRL+S}{ctrl+down}{ctrl+up}");
        ExcelWindow.Type("=(98*323){+}(312*97){enter}{ctrl+down}{ctrl+up}", cpm:100);
        Wait(2);
        
        
        ExcelWindow.Maximize().Focus();
        ExcelWindow.Type("{F5}");
        Wait(2);
        FindWindow(className : "Win32 Window:bosa_sdm_XL9", title : "Go To", processName : "EXCEL").Focus();
        ExcelWindow.Type("C350{enter}{esc}{esc}");
        Wait(2);
        
        
        ReadText(10);
        ExcelWindow.Type("{home}{end}{left}", cpm:200);
        Wait(1);
        ReadText(6);
        
        
        ExcelWindow.Focus();
        ExcelWindow.Type("{CTRL+S}{ctrld+down}{ctrl+up}", cpm:200);
        ExcelWindow.Type("243789897324987238991{enter}");
        ExcelWindow.Type("{CTRL+S}{ctrl+down}{ctrl+up}", cpm:200);
        Wait(2);
        
        
        Wait(2);
        ExcelWindow.Focus();
        ExcelWindow.Type("{CTRL+S}");
        OutlookWindow.Focus();
        
      Log(message:"Outlook back to main window");   
        Wait(2); 
        OutlookWindow.Maximize();
        Wait(2); 
        OutlookWindow.Type("{down}{down}{down}");
        Wait(5);
        OutlookWindow.Type("{down}{down}{down}{down}");
        
        
        EdgeWindow.Focus();
        EdgeWindow.Type("{HOME}");
        ReadText(28);
        
        EdgeWindow2.Focus();
        EdgeWindow2.Type("{HOME}");
        ReadText(13);
        
        
      Log(message:"Freemind actions");   
        CopyFile("\\\\lab-ms01\\VSIShare\\_VSI_Content\\mm-newver\\1.mm", destinationPath: TempDirectory + "\\LoginPI\\UserMindmap.mm",continueOnError:false, overwrite:true); 
        ShellExecute(TempDirectory + "\\LoginPI\\UserMindmap.mm", waitForProcessEnd: false, forceKillOnExit: false);
        var FreemindWindow = FindWindow(className : "Win32 Window:SunAwtFrame", title : "* FreeMind - MindMap Mode *", processName : "javaw");
        FreemindWindow.Focus().Maximize();
        FreemindWindow.Type("{esc}{esc}");
        Wait(1);
        FreemindWindow.Type("{right}{up}{up}{right}{right}", cpm:300);
        Wait(2);
        FreemindWindow.Type("{esc}");
        Wait(1);
        FreemindWindow.Type("{left}{up}{up}{left}{left}{up}", cpm:300);
        Wait(3);
        FreemindWindow.Type("{CTRL+S}");
        
        
        FreemindWindow.Type("{INSERT}");
        TypeText(18);
        FreemindWindow.Type("{CTRL+S}");
        FreemindWindow.Type("{enter}{enter}{esc}");
        
        
        FreemindWindow.Focus().Maximize();
        FreemindWindow.Type("{CTRL+S}");
        FreemindWindow.Type("{esc}{esc}{esc}");
        Wait(1);
        FreemindWindow.Type("{right}{down}{down}{down}{right}{right}{down}{down}{right}{right}{right}{right}{down}", cpm:300);
        Wait(3);
        
        
        FreemindWindow.Type("{INSERT}");
        TypeText(20);
        FreemindWindow.Type("{enter}{enter}{esc}");
        FreemindWindow.Focus().Maximize();
        FreemindWindow.Type("{esc}{esc}{esc}");
        FreemindWindow.Focus().Maximize();
        FreemindWindow.Type("{CTRL+S}");
        Wait(3);
        
        
      Log(message:"WinWord4"); 
        CopyFile(KnownFiles.WordDocument, TempDirectory + "\\LoginPI\\WinWord4.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"Winword4");
        ShellExecute("winword.exe " + TempDirectory + @"\LoginPI\WinWord4.docx", waitForProcessEnd: false, timeout: 30, forceKillOnExit: false);
        var WinWord4 = FindWindow(className : "Win32 Window:OpusApp", title : "WinWord4*", processName : "WINWORD");
	    StopTimer(name:"Winword4");
        WinWord4.Focus().Maximize();
        TypeText(20);
        WinWord4.Type("{CTRL+S}");
        TypeText(20);
        WinWord4.Type("{CTRL+S}");
        TypeText(10);
        WinWord4.Type("{CTRL+S}");
        TypeText(20);
        WinWord4.Type("{CTRL+S}");
        TypeText(30);
        WinWord4.Type("{CTRL+S}");
        WinWord4.Close();
        
        
        FreemindWindow.Focus().Maximize();
        // PDF print
        
        
        ////////////////////////////////////  SEGMENT 4 /////////////////////////////////////////
        
        
        EdgeWindow.Focus();
        EdgeWindow.Type("{HOME}");
        ReadText(18);
        StopBrowser();
        
        
        Log(message:"Outlook message 8"); 
         //OUTLOOK create message (8) with attachment
        // Get files ready (attach to email)
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginPI\\OutLookAttach8.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"NewMsgWithAttachment8");
        var OutlookNewMSG8 =  ShellExecute("cmd.exe /c start outlook.exe /a "+TempDirectory+"\\LoginPI\\OutLookAttach8.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false, continueOnError: true);
        var OutlookWindowNewMSG8 = FindWindow(title:"*HTML*", timeout:10);
        StopTimer(name:"NewMsgWithAttachment8");
        Wait(2);
        OutlookWindowNewMSG8.Focus();
        Wait(2);
        OutlookWindowNewMSG8.Maximize();
        Wait(3);
        OutlookWindowNewMSG8.Type("John Dummy");
        OutlookWindowNewMSG8.Type("{TAB}{TAB}{TAB}{TAB}{TAB}");
        OutlookWindowNewMSG8.Type("The quick brown fox jumps over the lazy dog. Message 8");
        KeyDown(KeyCode.CTRL);
        OutlookWindowNewMSG8.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookWindowNewMSG8.Close();
        // Clean up
        RemoveFile(path: TempDirectory+"\\LoginPI\\OutLookAttach8.docx");  
        
        
   Log(message:"Outlook back to main window");        
        //Outlook, back to main window
        OutlookWindow.Focus();
        Wait(2); 
        OutlookWindow.Maximize();
        Wait(2); 
        OutlookWindow.Type("{down}{down}{down}");
        Wait(5);
        OutlookWindow.Type("{down}{down}{down}{down}");
        Wait(5);
        
        
      Log(message:"Outlook message 9"); 
          //OUTLOOK create message (9) with attachment
        // Get files ready (attach to email)
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginPI\\OutLookAttach9.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"NewMsgWithAttachment9");
        var OutlookNewMSG9 =  ShellExecute("cmd.exe /c start outlook.exe /a "+TempDirectory+"\\LoginPI\\OutLookAttach9.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false, continueOnError: true);
        var OutlookWindowNewMSG9 = FindWindow(title:"*HTML*", timeout:10);
        StopTimer(name:"NewMsgWithAttachment9");
        Wait(2);
        OutlookWindowNewMSG9.Focus();
        Wait(2);
        OutlookWindowNewMSG9.Maximize();
        Wait(3);
        OutlookWindowNewMSG9.Type("Jane Dummy");
        OutlookWindowNewMSG9.Type("{TAB}{TAB}{TAB}{TAB}{TAB}");
        OutlookWindowNewMSG9.Type("The quick brown fox jumps over the lazy dog. Message 9");
        KeyDown(KeyCode.CTRL);
        OutlookWindowNewMSG9.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookWindowNewMSG9.Close();
        // Clean up
        RemoveFile(path: TempDirectory+"\\LoginPI\\OutLookAttach9.docx"); 
        
        
        // Photo viewer actions
        
        
        EdgeWindow2.Focus();
        EdgeWindow2.Type("{HOME}");
        ReadText(20);
        
      Log(message:"MS Edge video"); 
        try {
        ShellExecute("microsoft-edge:https://www.ted.com/talks/john_soluri_the_dark_history_of_bananas", waitForProcessEnd: false);
        } catch {}
        Wait(10);
        var MEVideo = FindWindow(className : "Win32 Window:Chrome_WidgetWin_1", title : "*TED*", processName : "msedge");
        MEVideo.Focus();
        Wait(2);
        //MEVideo.Type("{SPACE}");
        MEVideo.FindControl(title: "Play").Click();
        Wait(2);
        MEVideo.Type("{CTRL+W}");
        
        
        EdgeWindow2.Focus();
        EdgeWindow2.Type("{HOME}");
        ReadText(26);
        EdgeWindow2.Close();
        
        
      Log(message:"Outlook message 10"); 
          //OUTLOOK create message (10) with attachment
        // Get files ready (attach to email)
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginPI\\OutLookAttach10.docx",continueOnError: true,overwrite:true);
        StartTimer(name:"NewMsgWithAttachment10");
        var OutlookNewMSG10 =  ShellExecute("cmd.exe /c start outlook.exe /a "+TempDirectory+"\\LoginPI\\OutLookAttach10.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false, continueOnError: true);
        var OutlookWindowNewMSG10 = FindWindow(title:"*HTML*", timeout:10);
        StopTimer(name:"NewMsgWithAttachment10");
        Wait(2);
        OutlookWindowNewMSG10.Focus();
        Wait(2);
        OutlookWindowNewMSG10.Maximize();
        Wait(3);
        OutlookWindowNewMSG10.Type("Jane Dummy");
        OutlookWindowNewMSG10.Type("{TAB}{TAB}{TAB}{TAB}{TAB}");
        OutlookWindowNewMSG10.Type("The quick brown fox jumps over the lazy dog. Message 10");
        KeyDown(KeyCode.CTRL);
        OutlookWindowNewMSG10.Type("s");
        KeyUp(KeyCode.CTRL);
        Wait(5);
        OutlookWindowNewMSG10.Close();
        // Clean up
        RemoveFile(path: TempDirectory+"\\LoginPI\\OutLookAttach10.docx");  
        
      
      Log(message:"MS Edge video 2"); 
        try {
        ShellExecute("microsoft-edge:https://www.ted.com/talks/john_soluri_the_dark_history_of_bananas", waitForProcessEnd: false);
        } catch {}
        Wait(10);
        var MEVideo2 = FindWindow(className : "Win32 Window:Chrome_WidgetWin_1", title : "*TED*", processName : "msedge");
        MEVideo2.Focus();
        Wait(2);
        MEVideo2.FindControl(title: "Play").Click();
        Wait(2);
        MEVideo2.Type("{CTRL+W}");
        
        
      Log(message:"MS Edge video 3"); 
        try {
        ShellExecute("microsoft-edge:https://www.ted.com/talks/john_soluri_the_dark_history_of_bananas", waitForProcessEnd: false);
        } catch {}
        Wait(10);
        var MEVideo3 = FindWindow(className : "Win32 Window:Chrome_WidgetWin_1", title : "*TED*", processName : "msedge");
        MEVideo3.Focus();
        Wait(2);
        MEVideo3.FindControl(title: "Play").Click();
        Wait(2);
        MEVideo3.Type("{CTRL+W}"); 
        
      
      Log(message:"Winword5"); 
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginEnterprise\\WinWord5.docx",continueOnError: true,overwrite:true);
        ShellExecute("winword.exe " + TempDirectory+"\\LoginEnterprise\\WinWord5.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        
        var WinWord5 = FindWindow(className : "Win32 Window:OpusApp", title : "WinWord5*", processName : "WINWORD");
        WinWord5.Focus().Maximize();
        Wait(2);
        ReadText(24);
        WinWord5.Type("{CTRL+S}");
        WinWord5.Focus();
        
        
     Log(message:"Winword6"); 
        CopyFile(KnownFiles.WordDocument, TempDirectory+"\\LoginEnterprise\\WinWord6.docx",continueOnError: true,overwrite:true);
        ShellExecute("winword.exe " + TempDirectory+"\\LoginEnterprise\\WinWord6.docx", waitForProcessEnd: false, timeout: 10, forceKillOnExit: false);
        
        var WinWord6 = FindWindow(className : "Win32 Window:OpusApp", title : "WinWord6*", processName : "WINWORD");
        WinWord6.Focus().Maximize();
        Wait(2);
        TypeText(24);
        WinWord6.Type("{CTRL+S}");
        WinWord6.Focus();   
        WinWord6.Close();
        
        
        WinWord5.Focus();
        // PDF Print
        WinWord5.Type("{CTRL+S}");
        WinWord5.Close();
        
        
        FreemindWindow.Focus().Maximize();
        FreemindWindow.Type("{esc}{esc}{esc}");
        FreemindWindow.Focus().Maximize();
        
        
        Wait(2);
        FreemindWindow.Close();
        
        Wait(2);
        OutlookWindow.Close();
        Wait(2);
        WinWord2.Close();
        
        
        ExcelWindow.Focus();
        ExcelWindow.Type("{CTRL+DOWN}{CTRL+UP}{CTRL+S}");
        ExcelWindow.Close();
        
        Wait(2);
        PDFReader1.Focus();
        Wait(2);
        PDFReader1.Close();
        
        
        PowerPoint1.Focus();
        Wait(2);
        PowerPoint1.Type("{CTRL+DOWN}{CTRL+UP}{CTRL+S}");
        PowerPoint1.Close();
        
        ////////////////////////////////////  LEFTOVER CLEANUPS /////////////////////////////////////////
       
        Wait(2);
        PDFReader2.Close();
        
        
        Wait(2);
        RemoveFile(path: TempDirectory+"\\LoginEnterprise\\WinWord2.docx", continueOnError: true);
        
        
       
      
       Wait(seconds:60, showOnScreen:true, onScreenText:"Everything should be gone now..");

   }


    
    // Funtions
       //typetext assumes roughly 3 characters per second.. can be improved
       public void TypeText(int Seconds) {
       var chars = " ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz 0123456789  ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz 0123456789  ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz 0123456789  ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz 0123456789  ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz 0123456789 ";
       int RequiredNumberOfChars = Seconds * 3; 
       chars = chars.Substring(0, RequiredNumberOfChars);
       Log(message:"TypeText: " + Seconds + "s resulting in " + RequiredNumberOfChars + " characters being typed");
       Type(chars);
       }

      // Read function, needs to be improved seconds * 1.. i know..
      public void ReadText(int Seconds) {
        int RequiredNumberOfBumps = Seconds * 1;
            for (int i = 0; i < RequiredNumberOfBumps; i++)
            {
            Type("{down}{down}{up}");
            }  
       }


      }
       
    
    
    
