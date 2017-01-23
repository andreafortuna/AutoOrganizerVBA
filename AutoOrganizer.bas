Attribute VB_Name = "AutoOrganizer"
'-----------------------------------------------
' Auto-Organizer Module for Outlook
'-----------------------------------------------
' @2017 Andrea Fortuna https://andreafortuna.org
'-----------------------------------------------



'--------------- CONFIGURATION -----------------
'--- Archive folder
Private Const FolderPath = "\Opencases"
'--- Stop words(comma separated)
Private Const Stopwords = "RE: ,R: ,FW: ,I: "
'-----------------------------------------------

Function GetFolder(ByVal FolderPath As String) As Outlook.Folder
 Dim TestFolder As Outlook.Folder
 Dim FoldersArray As Variant
 Dim i As Integer
 On Error GoTo GetFolder_Error
 If Left(FolderPath, 2) = "\\" Then
 FolderPath = Right(FolderPath, Len(FolderPath) - 2)
 End If
 FoldersArray = Split(FolderPath, "\")
 Set TestFolder = Application.Session.Folders.Item(FoldersArray(0))
 If Not TestFolder Is Nothing Then
 For i = 1 To UBound(FoldersArray, 1)
 Dim SubFolders As Outlook.Folders
 Set SubFolders = TestFolder.Folders
 Set TestFolder = SubFolders.Item(FoldersArray(i))
 If TestFolder Is Nothing Then
 Set GetFolder = Nothing
 End If
 Next
 End If
 Set GetFolder = TestFolder
 Exit Function
GetFolder_Error:
 Set GetFolder = Nothing
 Exit Function
End Function

Sub AutoOrganize()
    Dim currentExplorer As Explorer
    Set currentExplorer = Application.ActiveExplorer
    Set Selection = currentExplorer.Selection
    Dim newFolder As String
    
    'Stopwords array
    Dim arrStop() As String
    arrStop = Split(Stopwords, ",")
    
    For Each CurrentItem In Selection
        If CurrentItem.Class = olMail Then
            Set currentMail = CurrentItem
            Dim currentSubject As String
            
            currentSubject = currentMail.Subject
            
            'Subject cleanup
            For Each currentWord In arrStop
                currentSubject = Replace(currentSubject, currentWord, "")
            Next
            currentSubject = Trim(currentSubject)
            'confirm
            If newFolder = "" Then newFolder = InputBox("Create new folder", "Auto Organizer", currentSubject)
                                    
            'if cancel...
            If newFolder = "" Then Exit Sub
                                    
            'folders management
            Dim myolApp As Outlook.Application
            Dim myNamespace As Outlook.NameSpace
            Set myolApp = CreateObject("Outlook.Application")
            Set myNamespace = myolApp.GetNamespace("MAPI")

            Dim incidentFolder As Outlook.Folder
            Dim destFolder As Outlook.Folder
                                
            Set incidentFolder = GetFolder("\\" & myNamespace.GetDefaultFolder(olFolderInbox).Parent.Name & FolderPath)
               
            On Error GoTo createFolder
            Set destFolder = incidentFolder.Folders(newFolder)
            currentMail.Move destFolder
            GoTo skipLoop
createFolder:
            incidentFolder.Folders.Add newFolder, olFolderInbox
            Set destFolder = incidentFolder.Folders(newFolder)
            currentMail.Move destFolder
            GoTo skipLoop
        End If
skipLoop:
    Next
End Sub
