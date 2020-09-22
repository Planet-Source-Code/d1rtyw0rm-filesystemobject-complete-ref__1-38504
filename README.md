<div align="center">

## FileSystemObject Complete Ref\.


</div>

### Description

Here a description about FSO(File System Object) and howto use it
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[d1rtyw0rm](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/d1rtyw0rm.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/d1rtyw0rm-filesystemobject-complete-ref__1-38504/archive/master.zip)





### Source Code

<link rel="stylesheet" href="http://205.151.63.123/d1rtyw0rm/style.css" type="text/css">
<BR><I>Active X:</I> Microsoft Scripting RunTime<BR><BR>
FSO is there to help coder to work on file and directory.<BR><BR>
The FSO model got 5 object<BR>
<CENTER><B>Drive - Folder - File - FileSystemObject - TextStream</B></CENTER><BR><BR>
<B>Drive Object</B><BR><BR>
<I>Total Size</I> : Disk Dimension<BR><BR>
<I>AvailableSizeFreeSpace</I> : To get Free space of a disk<BR><BR>
<I>DriveLetter</I> : Give the drive letter<BR><BR>
<I>DriveType</I> : Give the type of a drive(fix,net,CD-Rom,etc...)<BR><BR>
<I>SerialNumber</I> : give the serial number of a disk<BR><BR>
<I>FileSystem</I> : NTFS,Fat,Fat32 etc.<BR><BR>
<I>IsReady</I> : return True if drive is ready<BR><BR>
<I>VolumeName</I> : ...<BR><BR><BR>
<U>Exemple of use</U><BR><BR>
<FONT COLOR="NAVY">
Dim fso as new filesystemobject<BR>
dim d as drive<BR>
set d = fso.getdrive("c:") 'Creation of the drive object<BR>
msgbox d.totalsize<BR>
msgbox d.filesystem<BR>
</FONT><BR>
Exemple of what the code will return 13gb, NTFS<BR><HR><BR>
<B>FileSystemObject Object</B><BR><BR>
<I>CreateFolder</I> : Create a directory<BR><BR>
<I>FolderExist</I> : Return True if folder exist<BR><BR>
<I>GetParentFolder</I> : Return parent name directory<BR><BR>
<U>Exemple of use</U><BR><BR>
<FONT COLOR="NAVY">
Dim fso as New FileSystemObject<BR>
fso.CreateFolder("C:\test\x")<BR>
msgbox fso.FolderExist("C:\test")<BR>
msgbox fso.GetParentFolderName("C:\test\bbb")<BR>
</FONT><BR>
This example will first create the folder C:\test\x, after the first msgbox will return true cause we previously create the directory, the last msgbox will return 'C:\test'.<BR><HR><BR>
<B>Folder Object</B><BR><BR>
Allows to manage the repertories<BR><BR>
<I>Delete</I> : Delete a folder<BR><BR>
<I>Move</I> : Move a folder<BR><BR>
<I>Copy</I> : Copy a folder<BR><BR>
<I>Name</I> : Name of the folder<BR><BR>
<U>Exemple of use</U><BR><BR>
<FONT COLOR="NAVY">
dim fso as new filesystemobject<BR>
dim r as folder<BR>
set r = fso.GetFolder("c:\odsource")<BR>
msgbox r.name<BR>
r.copy "D:\"<BR>
r.delete<BR>
</FONT><BR>
First we create the folder object, the msgbox return "odsource", after it copy the folder on D:\ drive, and delete the folder on c:\odsource.<BR><HR><BR>
<B>File Object</B><BR><BR>
Allows to obtain information on a file<BR><BR>
<I>Attributes</I> : Return the file attribute (read only,hidden etc.)<BR><BR>
<I>Copy</I> : Copy a file<BR><BR>
<I>DataCreated</I> : Return creation date<BR><BR>
<I>DateLAstModified</I> : Return date of the last modification<BR><BR>
<I>Delete</I> : Delete a file<BR><BR>
<I>Drive</I> : return the drive where the file is<BR><BR>
<I>Move</I> :allow to move the file<BR><BR>
<I>Name</I> : return the file name<BR><BR>
<I>ParentFolder</I> : return the parent folder name of the file<BR><BR>
<I>Path</I> : Return the full file path(include the file name)<BR><BR>
<I>ShortNAme</I> : return the short name of the file<BR><BR>
<I>Size</I> : File dimention<BR><BR>
<I>Type</I> : return file type<BR><BR>
<U>Exemple of use</U><BR>
<FONT COLOR="NAVY">
dim fsoas new FileSystemObject<BR>
Dim f as File<BR><BR>
Set f = fso.GetFile<BR>("c:\d1rtyw0rm\visualbasicforum.com")<BR>
msgBox f.type<BR>
msgbox f.DateCreated<BR>
msgbox f.Path<BR>
</FONT><BR><BR>
The first msgbox return "COM", second "11-08-2002", third "C:\d1rtyw0rm\visualbasicforum.com"<BR><HR><BR>
<B>TextStream Object</B><BR><BR>
Allow to manage a text file. Binary file cannot be managed by FSO.<BR><BR>
<I>Write</I> : Write a text line without crlf<BR><BR>
<I>Write Line</I> : Write a text line with crlf<BR><BR>
<I>WriteBlankLines</I> : Write a number of crlf<BR><BR>
<I>Close</I> : Close the file<BR><BR>
<I>Read</I> : Reads a specific number of character<BR><BR>
<I>ReadLine</I> : Read a complete line<BR><BR>
</I>ReadAll</I> : Read complete file<BR><BR>
<U>Exemple of use</U><BR>
<FONT COLOR="NAVY">
Dim fso As New FileSystemObject<BR>
Dim ts as TextStream<BR>
Dim s As String<BR>
'Creation of the text file'<BR>
Set ts = fso.OpenTextFile("C:\vbf\dirtyworm.txt", ForWriting,True)<BR>
ts.WriteLine "d1rtyw0rm rul'Z"<BR>
ts.Write "http:\\www.d1rtyw"<BR>
ts.Write "0rm.ca.tc"<BR>
ts.BlankLines(4)<BR>
ts.WriteLine "odsource.com"<BR>
ts.Close<BR><BR>
'Open for adding'<BR><BR>
Set ts = fso.OpenTextFile("c:\odsource\dirtyworm.txt",ForAppending)<BR>
ts.WriteLine "This line will be under odsource.com"<BR>
ts.close<BR><BR>
'File Reading'<BR><BR>
Set ts = fso.OpenTextFile("C:\odsource\dirtyworm.txt", ForReading)<BR>
s = ts.Read(12)<BR>
MsgBox s 'Will print d1rtyw0rm ru'<BR>
s = ts.ReadLine 'Read the rest of the line'<BR>
s = ts.ReadAll<BR><BR>
MsgBox s<BR><BR>
ts.Close<BR>
</FONT><BR>
-d1rtyw0rm

