VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1
Public WithEvents Doc As Word.Document
Attribute Doc.VB_VarHelpID = -1

Public mainTriggered As Boolean
Public completionLength As Integer
Public startPos As Integer
Public tabKey As KeyBinding
Public escKey As KeyBinding
Public idleTime As Date
Private originalSelection(2) As Integer
Public Enabled As Boolean
Public buttonState As Boolean
Public font As font

Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    Debug.Print "Before Save"
    'resetCompletion
End Sub



Function textSelected(sel As Selection) As Boolean
    If sel.start < sel.End Then
        textSelected = True
    Else
        textSelected = False
    End If
End Function

Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    Debug.Print "WindowAvtivate"
    resetCompletion
    Enabled = True
End Sub

Private Sub App_WindowDeactivate(ByVal Doc As Document, ByVal Wn As Window)
    Debug.Print "WindowDeavtivate"
    resetCompletion
    Enabled = False
End Sub


Private Sub App_WindowSelectionChange _
    (ByVal sel As Selection)
    Debug.Print "Selection Changed"
    Enabled = True
    originalSelection(0) = sel.start
    originalSelection(1) = sel.End
    resetCompletion True
End Sub


Public Function resetCompletion(Optional bySelection As Boolean = False)
    On Error GoTo ErrHandler
    Debug.Print "Reseting Completion"

    If IsMissing(bySelection) Or bySelection = False Then
        startPos = Selection.start
    End If
    
    Dim delRange As Range
    Set delRange = Selection.Range.Duplicate
    If Not completionLength = 0 And Not delRange Is Nothing Then
        delRange.SetRange startPos, (startPos + completionLength)
        delRange.font = font
        delRange.Delete
    End If
   
    
    If Not IsMissing(bySelection) And bySelection = True Then
        Selection.SetRange originalSelection(0), originalSelection(1)
    End If
    
    If Not font Is Nothing And textSelected(Selection) Then
        Selection.Range.font = font
    End If
    
    resetParams
    Exit Function
ErrHandler:
    Debug.Print "Error occured while resetting the completion"
    resetParams
End Function

Public Function resetParams()
    MakeRequest "http://localhost:5000/stop", "POST", ""
    
    mainTriggered = False
    originalSelection(0) = Selection.start
    originalSelection(1) = Selection.End
    completionLength = 0
    startPos = Selection.start
    idleTime = Now()
    
    If Not tabKey Is Nothing Then
        tabKey.Clear
    End If
    
    If Not escKey Is Nothing Then
        escKey.Clear
    End If
End Function
Private Sub Doc_ContentControlBeforeStoreUpdate(ByVal ContentControl As ContentControl, Content As String)
    Debug.Print Content
End Sub
