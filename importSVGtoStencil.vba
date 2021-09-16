' source files https://github.com/ecceman/affinity
' files must have height/width changed before hand so that they scale correctly in visio
' sed -i.bak -e 's/width="100%"/width="10%"/;s/height="100%"/height="10%"/' *

Function readFile(path As String) As String
    Dim hF As Integer
    hF = FreeFile()
    Open path For Input As #hF
        readFile = Input$(LOF(hF), #hF)
    Close #hF
End Function

Sub Macro1()
  
    Const Switches As String = "/OGN /TW /B" ' Add "/S" for subfolders
    Dim StencilPath As String
    Dim StencilName As String
    Dim StencilPathName As String
    Dim StencilFileName As String
    Dim CommandLine As String
    Dim FolderName As String
    Dim TempName As String
    Dim FSO As New FileSystemObject
    Dim MyFile As Object
    
    FolderName = "D:\Downloads\affinity-master\gray\"
    TempName = FSO.GetSpecialFolder(TemporaryFolder) & "\" & FSO.GetTempName
    CommandLine = "cmd.exe /C dir " & Switches & " """ & FolderName & """ > """ & TempName & """"
    StencilPath = "D:\My Documents\"
    StencilName = "test_import"
    StencilFileName = StencilName & ".vssx"
    StencilPathName = StencilPath & StencilFileName
    
    Call CreateObject("WScript.Shell").Run(CommandLine, 2, True)

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set MyFile = FSO.OpenTextFile(TempName, 1)
    Arr = Split(MyFile.ReadAll, vbNewLine) ' Arr is zero-based array
    
    'Debug.Print (Arr)
    
    For Each Line In Arr
        ' MsgBox (Line)
                
        If Line = "" Then
            Exit For
        End If
        
        Dim split1 As Variant
        Dim split2 As Variant
        Dim LineName As String
        split1 = Split(Line, "sq_")
        split2 = Split(split1(1), ".")
        
        LineName = split2(0)
        
        Dim DiagramServices As Integer
        DiagramServices = ActiveDocument.DiagramServicesEnabled
        ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150
    
        Dim vsoDoc1 As Visio.Document
        Set vsoDoc1 = Application.Documents.Item(StencilPathName)
        Dim UndoScopeID1 As Long
        UndoScopeID1 = Application.BeginUndoScope("New Master")
        Dim vsoMaster1 As Visio.Master
        Set vsoMaster1 = vsoDoc1.Masters.Add
        vsoMaster1.Name = LineName
        vsoMaster1.Prompt = ""
        vsoMaster1.IconSize = visNormal
        vsoMaster1.AlignName = visCenter
        vsoMaster1.MatchByName = False
        vsoMaster1.IconUpdate = visAutomatic
        Application.EndUndoScope UndoScopeID1, True
        
        Application.Documents.Item(StencilPathName).Masters.ItemFromID(vsoMaster1.ID).Open.OpenDrawWindow
    
        Application.Windows.ItemEx(StencilFileName & ":Stencil:" & LineName).Activate
    
        Dim UndoScopeID2 As Long
        UndoScopeID2 = Application.BeginUndoScope("Insert")
        Application.ActiveWindow.Master.Import FolderName & Line
        Application.EndUndoScope UndoScopeID2, True
    
        Application.ActiveWindow.Master.Shapes.ItemFromID(5).OpenSheetWindow
    
        'Application.ActiveWindow.Shape.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = "215.9 mm*0.1"
    
        'Application.ActiveWindow.Shape.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "215.9 mm*0.1"
    
        Application.ActiveWindow.Close
    
        Application.ActiveWindow.Master.Close
    
        'Restore diagram services
        ActiveDocument.DiagramServicesEnabled = DiagramServices

    
        
    Next
    
    Application.Documents.Item(StencilPathName).SaveAs StencilPathName

   Application.ActiveWindow.WindowState = visWSMaximized

End Sub
