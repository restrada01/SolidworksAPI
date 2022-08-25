' Macro for automated date stamping for all drawings in a selected folder
' Created: Rafael Estrada - Aug 24, 2022

' IMPORTANT:
' before running, make sure you go to Tools > References, and then
' Check both Microsoft Scripting Runtime and Microsoft Shell Control and Automation

Option Explicit
Private Const BIF_RETURNONLYFSDIRS As Long = &H1

' Function to get a user-specified path from a popup box
Function BrowseFolder(Optional Caption As String, _
    Optional InitialFolder As String) As String
    ' use shell control to navigate folder popup box
    Dim SH As Shell32.Shell
    Dim F As Shell32.folder

    Set SH = New Shell32.Shell
    Set F = SH.BrowseForFolder(0&, Caption, BIF_RETURNONLYFSDIRS, InitialFolder)
    If Not F Is Nothing Then
        BrowseFolder = F.Items.Item.Path
    End If

End Function

' Function to return filename without filetype extension
Function GetFileNameWithoutExtension(filePath As String) As String
    GetFileNameWithoutExtension = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
End Function

Sub main()
    Dim swApp               As SldWorks.SldWorks
    Dim swModel             As SldWorks.DrawingDoc
    Dim myNote              As SldWorks.Note
    Dim myAnnotation        As SldWorks.Annotation
    Dim Path                As String
    Dim sFileName           As String
    Dim boolstatus          As Boolean
    Dim swTextFormat        As SldWorks.TextFormat
    Dim i                   As Long
    Dim j                   As Long
    Dim boolRet             As Boolean
    Dim swCustPrpMgr        As SldWorks.CustomPropertyManager
    Dim longstatus          As Long
    Dim longwarnings        As Long
    Dim Value               As String
        
    Set swApp = Application.SldWorks
    
    ' select the folder which you want to stamp drawings in
    Path = BrowseFolder()
    If Path = "" Then
        MsgBox "Please select the path and try again"
        End
        Else
            Path = Path & "\"
            ' if 'Date_Stamped' folder isnt created yet, initialize
            If Dir(Path & "Date_Stamped", vbDirectory) = "" Then
                MkDir (Path & "Date_Stamped")
            End If
    End If
    
    ' Loop through all drawings in the specified folder
    sFileName = Dir(Path & "*.slddrw")
    Do Until sFileName = ""
        ' Open drawing and extract sheet names
        Set swModel = swApp.OpenDoc6(Path & sFileName, swDocDRAWING, swOpenDocOptions_Silent, "", longstatus, longwarnings)
        Set swModel = swApp.ActiveDoc
        Dim swDrawingDoc As SldWorks.ModelDoc2
        Dim vSheetName As Variant
        Dim bRet As Boolean
        Set swDrawingDoc = swModel
        vSheetName = swModel.GetSheetNames
    
        ' Loop through all sheets existing in opened drawing
        For i = 0 To UBound(vSheetName)
            bRet = swDrawingDoc.ActivateSheet(vSheetName(i))   ' activate current sheet to print date
            swDrawingDoc.ViewZoomtofit2
            Dim swSheet As Sheet
            Dim vSheetProperties As Variant
            Set swSheet = swDrawingDoc.GetCurrentSheet
            vSheetProperties = swSheet.GetProperties2() ' get sheet properties to later find
        
            Set myNote = swModel.InsertNote(UCase(Format(Now, "DD MMM YYYY")))
                
            ' code below formats and places the text in the desired position
            If Not myNote Is Nothing Then
                    Set myAnnotation = myNote.GetAnnotation
                    If Not myAnnotation Is Nothing Then
                        For j = 0 To myAnnotation.GetTextFormatCount - 1
                            ' set text format to red, bold, and 5mm tall
                            myAnnotation.Color = 255
                            Set swTextFormat = myAnnotation.GetTextFormat(j)
                            swTextFormat.CharHeight = 0.005
                            swTextFormat.Bold = True
                            boolRet = myAnnotation.SetTextFormat(i, False, swTextFormat)
                        Next
                        
                        ' determine paperSize and position text accordingly. For paperSize definition, see
                        ' https://help.solidworks.com/2017/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swDwgPaperSizes_e.html
                        
                        Dim paperSize As Double
                        paperSize = vSheetProperties(0)
                        
                        ' Currently, only B and D-size papers are implemented
                        ' Copy If statement below and adjust position x-coord to add new paper sizes
                        If paperSize = 2 Then
                            boolstatus = myAnnotation.SetPosition(0.297, 0.034, 0) ' B-size Paper positioning
                        End If
                        If paperSize = 4 Then
                            boolstatus = myAnnotation.SetPosition(0.728, 0.034, 0) ' D-size Paper positioning
                        End If
                    End If
            End If
        Next i
    
        ' Save the drawing in "Path\Date_Stamped" folder , which can then be printed through swTaskScheduler
        If swModel.GetType = swDocumentTypes_e.swDocDRAWING Then
            Dim outFolder As String
            If outFolder = "" Then
                    outFolder = Path & "Date_Stamped"
            End If
                
            ' Take off trailing backlash if it exists
            If Right(outFolder, 1) = "\" Then
                outFolder = Left(outFolder, Len(outFolder) - 1)
            End If
                
            ' declare outFilePath as the file saved as a slddrw
            If outFolder <> "" Then
                Dim outFileName As String
                outFileName = GetFileNameWithoutExtension(swModel.GetPathName()) & ".slddrw"
                    
                Dim outFilePath As String
                outFilePath = outFolder & "\" & outFileName
                    
                Dim errs As Long
                Dim warns As Long
                ' Save stamped drawings and catch any exception if folder isn't found or file isn't saved
                If False = swModel.Extension.SaveAs(outFilePath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errs, warns) Then
                    Err.Raise vbError, "", "Failed to export Drawing to " & outFileName
                End If
            End If
        End If
    
        swApp.QuitDoc swModel.GetTitle
        sFileName = Dir
    Loop

End Sub
