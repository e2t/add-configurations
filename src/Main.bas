Attribute VB_Name = "Main"
Option Explicit

Const baseBufferFileName As String = "buffer.txt"

Dim swApp As Object
Dim currentDoc As ModelDoc2
Dim bufferFileName As String

Sub Main()
    Set swApp = Application.SldWorks
    bufferFileName = swApp.GetCurrentMacroPathFolder + "\" + baseBufferFileName
    Set currentDoc = swApp.ActiveDoc
    If Not currentDoc Is Nothing Then
        If currentDoc.GetType <> swDocDRAWING Then
            RestorePreviousText
            MainForm.Show
        End If
    End If
End Sub

Function Run(allText As String)  'hide
    KeepTextForFuture
    AddAllConfigurations MainForm.txtConf.text
    currentDoc.ForceRebuild3 True
End Function

Function ExitApp() 'mask for button
    Unload MainForm
    End
End Function

Sub AddAllConfigurations(text As String)
    Dim i As Variant
    Dim name As String
    Dim descr As String
    
    For Each i In Split(text, vbNewLine)
        SplitNameAndDescr i, name, descr
        If name <> "" Then
            AddConfiguration name, descr
        End If
    Next
End Sub

Sub SplitNameAndDescr(ByVal line As String, ByRef name As String, ByRef descr As String)
    Dim re As RegExp
    Dim m As match
    
    Set re = New RegExp
    re.IgnoreCase = True
    re.Pattern = "([^""]+)""([^""]*)"""
    If re.Test(line) Then
        Set m = re.Execute(line)(0)
        name = Trim(m.SubMatches(0))
        descr = Trim(m.SubMatches(1))
    Else
        name = Trim(line)
        descr = ""
    End If
End Sub

Sub AddConfiguration(name As String, descr As String)
    Dim conf As Configuration
    
    Set conf = currentDoc.ConfigurationManager.AddConfiguration(name, "", "", 0, "", descr)
    If conf Is Nothing Then
        Set conf = currentDoc.GetConfigurationByName(name)
        conf.Description = descr
    End If
    conf.BOMPartNoSource = swBOMPartNumber_ConfigurationName
End Sub

Function RestorePreviousText() 'mask for button
    Dim fso As FileSystemObject
    Dim txtStream As TextStream
    Dim text As String

    text = ""
    Set fso = New FileSystemObject
    On Error GoTo CannotReadFile
    Set txtStream = fso.OpenTextFile(bufferFileName)
    text = txtStream.ReadAll
    txtStream.Close
CannotReadFile:
    MainForm.txtConf.text = text
End Function

Function KeepTextForFuture() 'mask for button
    Dim fso As FileSystemObject
    Dim txtStream As TextStream

    Set fso = New FileSystemObject
    On Error GoTo CannotCreateFile
    Set txtStream = fso.CreateTextFile(bufferFileName)
    On Error Resume Next
    txtStream.Write MainForm.txtConf.text
    txtStream.Close
CannotCreateFile:
End Function
