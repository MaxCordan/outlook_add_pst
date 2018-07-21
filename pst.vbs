'On error resume next
    Const ForReading = 1 
    Dim arrFileLines() 
    Dim ObjFSO,objFile,objNet,objOutlook,WshShell
    
    i=0
    
    Set objnet = CreateObject("wscript.network")
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    Set objFile = objFSO.OpenTextFile("c:\users\" & objnet.UserName & "\PSTOUTPUT.txt", ForReading) 
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'For OUTLOOK 2013
    Set objOutlook = CreateObject("Outlook.Application.15")
 
    Do Until objFile.AtEndOfStream 
        Redim Preserve arrFileLines(i) 
        arrFileLines(i) = objFile.ReadLine 
                i = i + 1 
    Loop 
    
    objFile.Close
	Set myNS = objOutlook.GetNamespace("MAPI") 
    For Each strPath in arrFileLines 
		myNS.AddStore strPath	 
	Next  
