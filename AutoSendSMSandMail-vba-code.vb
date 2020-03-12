Option Explicit

Sub Worksheet_Activate()
btnConnTest.BackColor = &H8000000F
btnConnTest.Caption = "Test Connection"

End Sub

Sub Test()
Dim toSendRowsClct As Collection
Set toSendRowsClct = New Collection

Dim AS_stg, TargetSheet As Worksheet
Set AS_stg = ThisWorkbook.Sheets("AutoSend_Settings")
With AS_stg
   
        
    If .Range("B3") = Empty Then 'WORKBOOK NAME
            AS_stg.Activate
            .Range("B3").Activate
            MsgBox "Choose Workbook(Excel file) first in cell B3", vbCritical, "AutoSend Configuration Error"
            Exit Sub
    End If
        
    Dim Wb As Workbook
    On Error Resume Next
    Set Wb = Workbooks(.Range("B3").Value)
    On Error GoTo 0
    If Wb Is Nothing Then
     AS_stg.Activate
        .Range("B3").Activate
        MsgBox "check name of Workbook(Excel file) in cell B3, make sure file is currently open", vbCritical, "AutoSend Configuration Error"
        Exit Sub
    End If
    
        
        ''sprawdza czy jest wybrany arkusz i czy workbook go zawiera
    If .Range("C3") = Empty Then
        AS_stg.Activate
        .Range("C3").Activate
        MsgBox "Specify the sheet name in cell C3", vbCritical, "AutoSend Configuration Error"
        Exit Sub
    End If
        
    If SheetMissing(.Range("C3").Value, Wb) Then
            AS_stg.Activate
            .Range("C3").Activate
            MsgBox "Sheet with the name specified in cell C3 does not exist in B3 Workbook", vbCritical, "AutoSend Configuration Error"
            Exit Sub
    End If
    
    Dim TemplRow, LastRow, CustRow As Long
    Dim TemplName, TargetSheetName As String
    TargetSheetName = .Range("C3").Value
    
   
    Set TargetSheet = Wb.Sheets(TargetSheetName)
    
    
    Dim ContactCol, ContactMode As String
  
    
    If .Range("F3") = Empty Then 'metoda kontaktu
                AS_stg.Activate
                .Range("F3").Activate
                MsgBox "Choose contact mode in cell F3", vbCritical, "AutoSend Configuration Error"
                Exit Sub
            End If
    
    ContactMode = .Range("F3").Value
    
    Select Case ContactMode
        Case "SMS"
            'first check phone connection
        
        
        
        
        
        
        
        
            If .Range("E3") = Empty Then 'kolumna numery telefonow
                AS_stg.Activate
                .Range("E3").Activate
                MsgBox "No column for mobile numbers in E3 cell", vbCritical, "AutoSend Configuration Error"
                Exit Sub
            End If
            ContactCol = .Range("E3").Value
        Case "EMAIL"
            If .Range("d3") = Empty Then 'kolumna adresy email
                AS_stg.Activate
                .Range("d3").Activate
                MsgBox "No column for Email addresses in D3 cell", vbCritical, "AutoSend Configuration Error"
                Exit Sub
            End If
            ContactCol = .Range("D3").Value
        Case Else
            AS_stg.Activate
                .Range("F3").Activate
            MsgBox "Contact mode in cell F3 is invalid. Check for typos - valid values are EMAIL or SMS", vbCritical, "AutoSend Configuration Error"
    End Select
           
End With

'initialize Object that checks for existence of contact phone/email
Dim ContactRqs As Requirements
Set ContactRqs = New Requirements
ContactRqs.compmeth = "AND"
Call ContactRqs.add("=NOT(ISBLANK(" & ContactCol & "#))")

'initialize Object that checks for user-defined conditions
Dim UserRqs As Requirements
Set UserRqs = New Requirements
Call UserRqs.initMe 'prepares requirements from settings sheet

'create report sheet
Dim reportSheet As Worksheet
Dim howmanysheets As Integer
howmanysheets = 0
With ThisWorkbook
        howmanysheets = .Sheets.Count
        Set reportSheet = .Sheets.add(After:=.Sheets(howmanysheets))
        reportSheet.Name = "ASrprt_" & Format(Date, "ddmmmyy") & "sheet" & howmanysheets
        
End With
    
reportSheet.Range("A1").Value = Date
reportSheet.Range("B1").Value = Wb.Name
reportSheet.Range("C1").Value = TargetSheetName

'loop through target sheet from 2 row to the last existing contact detail entry
Dim ToSendRange As Range
Dim ExitMsg As String
With TargetSheet
    LastRow = .Range(ContactCol & "99999").End(xlUp).row
    For CustRow = 2 To LastRow
         If UserRqs.AreMet(TargetSheet, CustRow) And ContactRqs.AreMet(TargetSheet, CustRow) Then
            Call AddToSelection(ToSendRange, .Range("A" & CustRow).EntireRow)
            Call toSendRowsClct.add(CustRow)
            ExitMsg = ExitMsg & CustRow & " chosen to send " & ContactMode & vbNewLine
            .Range("A" & CustRow & ":" & ContactCol & CustRow).Copy Destination:=reportSheet.Range("D" & CustRow)
            reportSheet.Range("A" & CustRow).Value = TemplName
            reportSheet.Range("B" & CustRow).Value = Date
            reportSheet.Range("C" & CustRow).Value = ContactMode
            
        End If
    Next CustRow

.Activate
End With 'TargetSheet





If ContactMode = "SMS" Then
    Call sendSMSToRange(toSendRowsClct, TargetSheet, ContactCol)
ElseIf ContactMode = "EMAIL" Then
    Call sendEMAILToRange(toSendRowsClct, TargetSheet, ContactCol)
End If

'ToSendRange.Select
'Summary.Show vbModeless
'Summary.TextBox1.Text = ExitMsg


'this should be separate Create requirements once
'and save them as series of formulae to be applied to each row
'ALLREQSAremet should only check row against saved series of formulae
reportSheet.Columns.AutoFit
End Sub

Sub sendSMSToRange(ByVal RowsClct As Collection, ByVal TS As Worksheet, ByVal ContactCol As String)
    Dim CustRow, TagPair As Variant
    Dim rplmnt, phoneNo, SMSbody, TagName, TagValue As String
    Dim TagClct, usersClct As Collection
    Set TagClct = GetTagCollection()
    Set usersClct = New Collection
    Dim userEntry() As String
    Dim isTagUsed As Boolean
    
    
    
'   /TODO make usersdata collection that will be passsed to userfrom to perform various calculations without the need
'    /TODO of reading TS again, just collection of data arrays, that will be used based on replacement tags collection
'    /TODO create collection of arrays for each user rediming it



'leave only used Tags

    Dim OutTxt As String
    OutTxt = "|phoneNo"
 

    
    For Each TagPair In TagClct
        
        isTagUsed = InStr(TextBoxSMS.Text, TagPair(0)) > 0
        If isTagUsed Then
        OutTxt = OutTxt & "~" & TagPair(0)
        End If
        
    Next TagPair
    OutTxt = OutTxt & "|"
With TS
    For Each CustRow In RowsClct
           phoneNo = .Range(ContactCol & CustRow).Value
           OutTxt = OutTxt & phoneNo
           ReDim userEntry(0 To TagClct.Count) As String
           userEntry(0) = phoneNo
           Dim userEntryItemInd As Integer
           userEntryItemInd = 0
'add to String that will be send over WIFI
         For Each TagPair In TagClct
            isTagUsed = InStr(TextBoxSMS.Text, TagPair(0)) > 0
            If isTagUsed Then
                userEntryItemInd = userEntryItemInd + 1
                 TagValue = TagPair(1)
                 rplmnt = TS.Range(TagValue & CustRow).Value
                 OutTxt = OutTxt & "~" & rplmnt
                 userEntry(userEntryItemInd) = rplmnt
             End If
             
         Next TagPair
        usersClct.add (userEntry)
        OutTxt = OutTxt & "#"
    Next CustRow
End With 'TS


  OutTxt = OutTxt & "|"
  
  
  'Dim AS_stg As Worksheet
  'Set AS_stg = ThisWorkbook.Sheets("AutoSend_Settings")
Dim summy As Summary
Set summy = New Summary
summy.Show vbModeless
Set summy.UsrsClct = usersClct
Set summy.TagClct = TagClct
summy.CompiledData = OutTxt
summy.SMStempl = TextBoxSMS.Text
summy.tbSMSEdited.Text = TextBoxSMS.Text
summy.txbxIP.Text = TextBoxAndroidIP.Text
summy.txbxPORT.Text = TextBoxAndroidPORT.Text


End Sub


Sub testmini()
Dim OutApp As Object
Dim OutMail As Object


On Error Resume Next
Set OutApp = GetObject("Outlook.Application")
    If Err.Number <> 0 Then
        'Launch a new instance of Out
        Err.Clear
        'On Error GoTo 0
        Set OutApp = CreateObject("Outlook.Application")
        
    End If
 Set OutMail = OutApp.CreateItem(olMailItem)
                With OutMail
                    .To = "example@example.com"
                    .Subject = "Notification About Something"
                    .Body = "Hello Dear xxxx, today is " & DateValue(CStr(Now())) & ". This Email was sent on " & Now() & " You could customize the body to contain value of any cell or more complex calculations, Word Documents and many more."
                   
                    
                    .Display

                End With 'outmail
    
End Sub

Sub sendEMAILToRange(ByVal RowsClct As Collection, ByVal TS As Worksheet, ByVal ContactCol As String)
     Dim AS_stg As Worksheet
     Dim TemplRow As Integer
     Dim TemplName, DocLoc As String
     
     
     Set AS_stg = ThisWorkbook.Sheets("AutoSend_Settings")
     With AS_stg
     
     If .Range("A4") = Empty Then '=MATCH's formula cell - if smth wrong show user selection field in A3
            AS_stg.Activate
            .Range("A3").Activate
            MsgBox "Choose template in cell A3", vbCritical, "AutoSend Configuration Error"
            Exit Sub
    End If
 ' //dimming, moving template stuff into emaill proc
    TemplRow = .Range("A4").Value 'Get template path row
    TemplName = .Range("A3").Value 'Get template doc Name
    DocLoc = .Range("B" & TemplRow).Value 'full path to the file using row MATCH'ed in A4
     If Dir(DocLoc) = "" Then
         AS_stg.Activate
            .Range("B" & TemplRow).Activate
            MsgBox "I cannot find file under the specified path (B7)", vbCritical, "AutoSend Configuration Error"
     Exit Sub
     End If
    
    
    End With
    
   
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim OutApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    Dim TagClct As Collection
    Set TagClct = GetTagCollection()
    
    'remove unneeded tags from collection
    
    
    
    
    
   ' GoTo jumptotests
    
    'open word template
    On Error Resume Next 'if Word is already Running
    'Set WordApp = New Word.Application
    Set WordApp = GetObject("Word.Application")
    If Err.Number <> 0 Then
       ' MsgBox Err.Number & " " & Err.Description
        'Launch a new instance of Word
        Err.Clear
        'On Error GoTo 0
        Set WordApp = CreateObject("Word.Application")

    End If
    WordApp.Visible = True
    
    On Error Resume Next 'if Outlook is already Running
    'Set OutApp = New Outlook.Application
    Set OutApp = GetObject("Outlook.Application")
    If Err.Number <> 0 Then
        'Launch a new instance of Out
        Err.Clear
        'On Error GoTo 0
        Set OutApp = CreateObject("Outlook.Application")
        
    End If
    'OutApp.Visible = True 'object does not support this method
    
jumptotests:
    
    Dim CustRow, TagPair As Variant
    Dim Email, EmailSubject, TagName, TagValue, replmnt As String
    Dim editor As Object
    Dim TagRow As Integer
    
With TS
    For Each CustRow In RowsClct
    
           Email = .Range(ContactCol & CustRow).Value
''            ' TODO OOP-allow for set of requirements set in worksheet
''            ' i wtedy If Requirements.AllMet() Then
''            If DaysLeft < 0 Then
''                SprawdzPlatnosc = SprawdzPlatnosc & CustRow & " ! po dacie ! sprawdz platnosc " & Name & " " & Surname & vbNewLine
''            ElseIf DaysLeft < Uprzedzenie Then
''                Call AddToSelection(ToSendRange, .Range("A" & CustRow).EntireRow)
''                ExitMsg = ExitMsg & CustRow & " przygotowuje email. " & Email & " " & Name & " " & Surname & vbNewLine
               Set WordDoc = WordApp.Documents.Open(Filename:=DocLoc, ReadOnly:=False)
               EmailSubject = TextBoxEmailSubject.Text
                    For Each TagPair In TagClct
                        
                        
                        
                        TagName = TagPair(0)
                        TagValue = TagPair(1)
                        replmnt = TS.Range(TagValue & CustRow).Value
                        EmailSubject = Replace(EmailSubject, TagName, replmnt)
                        With WordDoc.Content.Find
                            .Text = TagName
                            .Replacement.Text = replmnt
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll 'Forward:=True, Wrap:=wdFindContinue
                        End With

                    Next TagPair
                    Dim stringy As String
                'stringy = WordDoc.Content.Text 'without doing it first email doesnt copy at all ?!! WTF Microsoft!
                WordDoc.Content.Copy
                
                Set OutMail = OutApp.CreateItem(olMailItem)
                With OutMail
                    .To = Email
                    .Subject = EmailSubject
                    .BodyFormat = olFormatHTML
                   
                    
                    .Display
                    Set editor = .GetInspector.WordEditor
                    
                    editor.Content.Paste

                End With 'outmail
                
                
                WordDoc.Close False
    Next CustRow

.Activate
End With 'TS




End Sub


Function GetTagCollection() As Collection
    Dim AS_stg As Worksheet
    Set AS_stg = ThisWorkbook.Sheets("AutoSend_Settings")
    
    Dim TagRow, LastTagRow As Long
    Dim TagName, TagValue As String
    Dim RetCollection As Collection
    Set RetCollection = New Collection
    With AS_stg
         LastTagRow = .Range("D" & "99999").End(xlUp).row
    If LastTagRow < 7 Then
        AS_stg.Activate
                .Range("D7").Activate
                MsgBox "At least one tag required!!!", vbCritical, "AutoSend Configuration Error"
                Err.Raise Number:=vbObjectError + 513, _
              Description:="No tags in the config list !"
    End If
    
    
    For TagRow = 7 To LastTagRow
                        TagName = .Range("D" & TagRow).Value
                        TagValue = .Range("E" & TagRow).Value
                        Dim ar(0 To 1) As String
                        ar(1) = TagValue
                        ar(0) = TagName
                        Call RetCollection.add(ar)
        Next TagRow
    End With
    
    Set GetTagCollection = RetCollection
End Function


Private Sub btnConnTest_Click()
Call ConnTest_btnChanger(btnConnTest, TextBoxAndroidIP, TextBoxAndroidPORT)
End Sub

Private Sub EMAILbtn_Click()
ActiveSheet.Range("F3").Value = "EMAIL"
Call Test
End Sub

Private Sub SMSbtn_Click()
ActiveSheet.Range("F3").Value = "SMS"
Call Test
End Sub



Private Sub TextBoxSMS_Change()
Dim txtln As Integer
txtln = Len(TextBoxSMS.Text)
If txtln < 161 Then
    TextBoxSMS.BackColor = &HFFFFFF
Else
   TextBoxSMS.BackColor = &HAAAAFF
   
End If
ActiveSheet.Range("E1") = txtln


End Sub
