Attribute VB_Name = "Copilot"
Option Explicit

Private Declare PtrSafe Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpString As String, ByVal nSize As Integer) As Long

Dim lastHash As String
Dim originalAddedRange As Range
Dim X As New EventClassModule

Public Sub Register_Event_Handler()
 Set X.App = Word.Application
 Set X.Doc = Word.ActiveDocument
End Sub

Private Sub GetFragmentsAndLastTwoSentences(fragArray, lastWritten)
    Dim Doc As Document
    Set Doc = ActiveDocument
    Dim fragCount As Integer
    fragCount = 0
    Dim Sel As Range
    Set Sel = Selection.Range
    Dim lastTwoSentences As String
    lastWritten = Sel.Paragraphs.Last.Range.Text
    Dim para As Paragraph
    For Each para In Doc.Paragraphs
        Dim frag As Range
        For Each frag In para.Range.Sentences
            If InStr(frag.Text, lastWritten) > 0 Then
                fragCount = fragCount + 1
                ReDim Preserve fragArray(1 To fragCount)
                fragArray(fragCount) = frag.Text
            End If
        Next frag
    Next para
    Dim fullText As String
    fullText = Doc.Range.Text
    Dim lastSentence As String
    Dim lastFullStop As Long
    lastFullStop = InStrRev(fullText, ".", Len(fullText) - 1)
    If lastFullStop > 0 Then
        lastSentence = Mid(fullText, lastFullStop + 1, Len(fullText) - lastFullStop)
        Dim secondLastFullStop As Long
        secondLastFullStop = InStrRev(fullText, ".", lastFullStop - 1)
        If secondLastFullStop > 0 Then
            lastTwoSentences = Mid(fullText, secondLastFullStop + 1, lastFullStop - secondLastFullStop)
        Else
            lastTwoSentences = Mid(fullText, 1, lastFullStop)
        End If
    End If
End Sub

Private Function ConvToBase64String(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
   
   Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Private Function ConvToHexString(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
        
    'Test with empty string input:
    'Hex:   d41d8cd98f00...etc
    'Base-64: 1B2M2Y8Asg...etc
        
    Dim oT As Object, oMD5 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
        
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
 
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oMD5.ComputeHash_2((TextToHash))
 
    If bB64 = True Then
       MD5 = ConvToBase64String(bytes)
    Else
       MD5 = ConvToHexString(bytes)
    End If
        
    Set oT = Nothing
    Set oMD5 = Nothing

End Function

Function GetLastParagraph() As String
    Dim Sel As Range
    Set Sel = Selection.Range
    
End Function

Function MakeRequest(url As String, method As String, body As String) As String
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    
    xhr.Open method, url, False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.send body
    
    If xhr.Status = 200 Then
        MakeRequest = xhr.responseText
    Else
        MakeRequest = ""
    End If
End Function


Function GetUID() As String
    Dim uid As String
    Dim props As Object
    Set props = ActiveDocument.CustomDocumentProperties
    Dim prop As DocumentProperty
    For Each prop In props
        If prop.Name = "copilot_uid" And prop.Type = msoPropertyTypeString Then
            uid = prop.Value
        End If
    Next
    If uid = "" Then
        ' Generate a new UID and store it in the copilot_uid property
        uid = Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss")
        props.Add Name:="copilot_uid", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=uid
    End If
    
    ' Return the UID
    GetUID = uid
End Function
Public Function GetCurrentPageText() As String
    Dim currentPageRange As Range
    Dim pageStart As Long
    Dim pageEnd As Long
    Dim pageText As String
    
    Set currentPageRange = Selection.Range
    
    'Get the start and end of the current page
    currentPageRange.Collapse wdCollapseStart
    pageStart = ActiveDocument.Bookmarks("\Page").start
    currentPageRange.End = pageStart
    currentPageRange.Collapse wdCollapseEnd
    pageEnd = ActiveDocument.Bookmarks("\Page").End
    
    'Set the current page range to the start and end of the current page
    currentPageRange.start = pageStart
    currentPageRange.End = pageEnd
    
    'Get the text of the current page and return it
    pageText = currentPageRange.Text
    GetCurrentPageText = pageText
End Function

Sub changeDetector()
    Dim currentHash As String
    currentHash = MD5(GetCurrentPageText())
    If currentHash = lastHash Then
        Application.OnTime Now() + TimeValue("00:00:01"), "changeDetector"
    Else
        Debug.Print "Change Detected"
        X.resetCompletion
        Application.OnTime Now() + TimeValue("00:00:01"), "TriggerMain"
    End If
    lastHash = currentHash
End Sub
Private Sub Document_Open()
    Register_Event_Handler
    ' Generate a UID for the document
    Dim uid As String
    uid = GetUID()
    
    ' Get the text of the document
    Dim docText As String
    docText = ActiveDocument.Content.Text
    
    ' Create a JSON request body with the UID and document text
    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    json("uid") = uid
    json("text") = docText
    
    Dim requestBody As String
    requestBody = JsonConverter.ConvertToJson(json)
    
    ' Make a request to the server
    Dim response As String
    response = MakeRequest("http://localhost:5000/update-store", "POST", requestBody)
    
    ' Process the response
    If response <> "" Then
        ' Do something with the response
    Else
        ' Handle the error
    End If
    
    ' Set the idle time to the current time
    X.idleTime = Now()
    
    ' Start timer to call ChangeDetector
    Application.Run "changeDetector"
End Sub

Sub TriggerMain()
    ' Check if the user has been idle for 3 seconds
    If Now() - X.idleTime >= TimeValue("00:00:03") And X.enabled Then
        ' Call the Main subroutine
        Main
        Debug.Print "Main triggered"
        Application.Run "changeDetector"
    Else
        ' Schedule TriggerMain to run again in 1 second
        Application.OnTime Now() + TimeValue("00:00:01"), "TriggerMain"
    End If
End Sub


Function GetKeyName(ByVal keyCode As WdKey, keyName) As String
    
    Dim bufSize As Long
    Dim bufPtr As String
    
    bufSize = 256
    bufPtr = String$(bufSize, vbNullChar)
    If GetKeyNameText(keyCode, ByVal bufPtr, bufSize) > 0 Then
        GetKeyName = Left$(bufPtr, InStr(1, bufPtr, vbNullChar) - 1)
        Debug.Print GetKeyName
    Else
        Debug.Print "Unable to get key name"
    End If
End Function



Sub AutoComplete(completion As String)

    Dim newRange As Range
    Set newRange = Selection.Range
    newRange.InsertAfter completion
    X.completionLength = newRange.End - newRange.start
    X.startPos = newRange.start
    
    MakeRequest "http://localhost:5000/start", "POST", ""
    
    ' Set up the key bindings
    Dim kb As keyBindings
    Set kb = Application.keyBindings
    

    Set X.escKey = kb.Add(KeyCategory:=wdKeyCategoryMacro, keyCode:=BuildKeyCode(Arg1:=wdKeyAlt, Arg2:=wdKeyF15), _
    Command:="ClearCompletion")
    ' Add a key binding for the Tab key
    If Not Application.keyBindings.Key(BuildKeyCode(wdKeyTab)) Is Nothing Then
        Application.keyBindings.Key(BuildKeyCode(wdKeyTab)).Clear
    End If
    Set X.tabKey = kb.Add(KeyCategory:=wdKeyCategoryMacro, _
    Command:="Complete", keyCode:=BuildKeyCode(wdKeyTab))
    'If X.addedRange.Paragraphs.Last.ID = Selection.Range.Paragraphs.Last.ID Then
       ' Call waitForChange
    'End If
End Sub

Sub ClearCompletion()
    Debug.Print "Esc key pressed"
    X.resetCompletion
End Sub

Sub Complete()
    Application.keyBindings.Key(BuildKeyCode(wdKeyTab)).Clear
    X.escKey.Clear
    MakeRequest "http://localhost:5000/stop", "POST", ""
End Sub

Public Sub Main()
    ' Get the fragments and last two sentences
    Dim fragArray() As String
    Dim lastWritten As String
    Call GetFragmentsAndLastTwoSentences(fragArray, lastWritten)
    
    ' Generate a UID for the document
    Dim uid As String
    uid = GetUID()
    
    ' Create the request body
    'Dim fragmentsJson As String
    'fragmentsJson = "[ """ & Join(fragArray, """ , """) & """ ]"
    
    Dim jsonBody As Object
    Set jsonBody = CreateObject("Scripting.Dictionary")
    jsonBody("lastWritten") = lastWritten
    jsonBody("uid") = uid
    
    Dim requestBody As String
    requestBody = ConvertToJson(jsonBody)
    
    ' Create and send the HTTP request
    Dim request As Object
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "POST", "http://localhost:5000/completions", False
    request.setRequestHeader "Content-Type", "application/json"
    request.send requestBody

    Dim jsonResponse As Object
    Set jsonResponse = ParseJson(request.responseText)("data")
    ' Display the response in a message box
    AutoComplete (jsonResponse("lastWritten"))
    lastHash = MD5(GetCurrentPageText)
End Sub










