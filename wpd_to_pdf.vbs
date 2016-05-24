' wpd_tp_pdf.vbs (VBScript)
' Convert WPD to PDF using MS Word for Windows

' Written by Anthony Cocciolo, 2016


Option Explicit

Dim zSourceDir
Dim oWord, FSO, masterCount, oPPT, oExcel, oBook, ext, folder, stringi, error_log, ofile
Dim orig_name, access_name, pres_name , new_fname
Dim manualNorm


MsgBox ("This script will create PDF files of all WPD files in a given directory, including sub-directories.  The files will have the suffix '_normalized.pdf'.  On the next screen, select the directory where the WPD files are located.")


Set FSO = CreateObject("Scripting.FileSystemObject")

' Need to provide the source directory
zSourceDir = BrowseFolder( "", False )
' zSourceDir = "E:\wpd_to_pdf\test_data"

if NOT fso.FolderExists (zSourceDir) then
	MsgBox ("Folder does not exist.  Quitting...")
	WScript.Quit
else
	MsgBox ("Press OK and the process will begin.  This may take awhile.  You will be notified when the process is complete")
end if


' standardize path
Set folder = fso.GetFolder(zSourceDir)
zSourceDir = folder.Path 


Set oWord = CreateObject("Word.Application")

masterCount = 0
error_log = ""

ConvFiles (zSourceDir)


' recursively normalize files for preservation and access
sub ConvFiles (currentDir)
	Dim oFolder, f, oDoc, subfolders, sf, frontPath
	
	Set oFolder = FSO.GetFolder(currentDir)
	
	' normalize all subfolders recursively
	Set subfolders = oFolder.SubFolders
   	For Each sf in subfolders
   		
   		if  not (sf.path = currentDir) then
    		ConvFiles (sf.Path)
    	end if
    Next
	
	' normalize each file
	for each f in oFolder.Files
	
		' ignore files that are hidden or system files
		if (   (f.attributes and 2) OR (f.attributes AND 4)  ) then
			' do nothing for hidden or system files
		else
	
		
			ext = lcase(FSO.GetExtensionName(f.path))
			
			' only apply to valid files
			if (ext = "wpd" ) then

					
					
					On Error Resume Next
					Set oDoc = oWord.Documents.Open(f.path, , True)
					
				
					if Err.number <> 0 then
						Error_log = Error_log & "Unable to open: " & f.path & " (Description: " & err.description & ")" & vbNewline 
						Set oDoc = Nothing
						
					else
						Set oDoc = oWord.ActiveDocument
					
						
						' create pdf
						new_fname = fso.GetParentFolderName(f.path) & "\" & fso.getbasename(f.name) & "_normalized.pdf"
				
						
						oDoc.SaveAs new_fname, 17
					
						if Err.Number <> 0 then
							Error_log = Error_log & "Error creating: " & new_fname & " (Description: " & err.description & ")" & vbNewline
						else
							access_name = new_fname
						end if
						
						oDoc.Close 0
						Set oDoc = Nothing
					end if
				
					On Error GoTo 0
			
		
			masterCount = masterCount + 1
				
			end if
			
		end if

	next 

end sub




oWord.Quit
Set oWord = Nothing


if error_log = "" then
	MsgBox "Creation of normalized files are completed.  " & masterCount & " files were normalized."
else
	Set ofile = fso.OpenTextFile ("pres_and_access.log", 2, true)
	ofile.writeline error_log
	ofile.close
	

	Set ofile = Nothing
	Set fso = Nothing

	MsgBox masterCount & " Files were attempted to be normalized, but there was an error with one or more.  These are included in the log file: pres_and_access.log"
end if



Function BrowseFolder( myStartLocation, blnSimpleDialog )
' This function generates a Browse Folder dialog
' and returns the selected folder as a string.
'
' Arguments:
' myStartLocation   [string]  start folder for dialog, or "My Computer", or
'                             empty string to open in "Desktop\My Documents"
' blnSimpleDialog   [boolean] if False, an additional text field will be
'                             displayed where the folder can be selected
'                             by typing the fully qualified path
'
' Returns:          [string]  the fully qualified path to the selected folder
'
'
' Function written by Rob van der Woude
' http://www.robvanderwoude.com

    Const MY_COMPUTER   = &H11&
    Const WINDOW_HANDLE = 0 ' Must ALWAYS be 0

    Dim numOptions, objFolder, objFolderItem
    Dim objPath, objShell, strPath, strPrompt

    ' Set the options for the dialog window
    strPrompt = "Select a folder:"
    If blnSimpleDialog = True Then
        numOptions = 0      ' Simple dialog
    Else
        numOptions = &H10&  ' Additional text field to type folder path
    End If
    
    ' Create a Windows Shell object
    Set objShell = CreateObject( "Shell.Application" )

    ' If specified, convert "My Computer" to a valid
    ' path for the Windows Shell's BrowseFolder method
    If UCase( myStartLocation ) = "MY COMPUTER" Then
        Set objFolder = objShell.Namespace( MY_COMPUTER )
        Set objFolderItem = objFolder.Self
        strPath = objFolderItem.Path
    Else
        strPath = myStartLocation
    End If

    Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, _
                                              numOptions, strPath )

    ' Quit if no folder was selected
    If objFolder Is Nothing Then
        BrowseFolder = ""
        Exit Function
    End If

    ' Retrieve the path of the selected folder
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path

    ' Return the path of the selected folder
    BrowseFolder = objPath
End Function


