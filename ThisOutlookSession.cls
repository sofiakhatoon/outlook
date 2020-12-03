VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub WriteTextFile2()

Dim myFile As String
myFile = "C:\Users\xyz\Downloads\put.txt"
Open myFile For Append As #1
Write #1, "Sofia", "user", 1000, "pw", 2000
Dim folder As Outlook.MAPIFolder
Dim folderPath As String, currItemType As String
Dim i As Integer

   Dim objOL As Outlook.Application
    Dim objItems As Outlook.Items
    Dim objFolder As Outlook.MAPIFolder
    Dim obj As Object
 
    Set objOL = Outlook.Application
    Set objFolder = objOL.ActiveExplorer.CurrentFolder
    Set objItems = objFolder.Items
 
    For Each obj In objItems
 
     With obj
 
    ' do whatever
        Write #1, ParseTextLinePair(obj.Body, "User ID:")
     
     End With

    Next
 
    Set obj = Nothing
    Set objItems = Nothing
    Set objFolder = Nothing
    Set objOL = Nothing

If TypeName(ActiveExplorer.Selection.Item(1)) = "MailItem" Then
 Set activeMailMessage = ActiveExplorer.Selection.Item(1)

    'Ensure Word Application is open
    Set WordApp = CreateObject("Word.Application")

    'Make Word Application visible
    WordApp.Visible = True

    'strAddr = ParseTextLinePair(activeMailMessage.GetInspector().WordEditor.Range.FormattedText, "User ID:")
strAddr = ParseTextLinePair(activeMailMessage.Body, "User ID:")
    'Copy the formatted text:
  '  Write #1, strAddr

 MsgBox strAddr
 
End If

Close #1
End Sub
Function ParseTextLinePair _
  (strSource As String, strLabel As String)
    Dim intLocLabel As Integer
    Dim intLocCRLF As Integer
    Dim intLenLabel As Integer
    Dim strText As String
    intLocLabel = InStr(strSource, strLabel)
    intLenLabel = Len(strLabel)
        If intLocLabel > 0 Then
        intLocCRLF = InStr(intLocLabel, strSource, vbCrLf)
        If intLocCRLF > 0 Then
            intLocLabel = intLocLabel + intLenLabel
            strText = Mid(strSource, _
                            intLocLabel, _
                            intLocCRLF - intLocLabel)
        Else
            intLocLabel = _
              Mid(strSource, intLocLabel + intLenLabel)
        End If
    End If
    ParseTextLinePair = Trim(strText)
End Function



