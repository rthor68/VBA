--Will open a blank compose e-mail window
Sub MailTo()

 ThisWorkbook.FollowHyperlink "mailto:" & IIf(sTo <> "", sTo, "") & "?" & IIf(sCC <> "", "&cc=" & sCC, 

"") & IIf(sBCC <> "", "&bcc=" & sBCC, "") & "&subject=" & sSubject

End Sub

--generate e-mail to multiple recipients froma spreadsheet that contains some of the addresses

Private Sub CommandButton72_Click()

    Dim date2  As String    ' Date field
    Dim email2 As String    ' Who Email is Going To
    Dim N1     As String    ' Name Field
    Dim subject1 As String  ' What the email is about
    Dim ccemail As String   ' Inspectors email
    Dim email3 As String    ' Prompted Email
    
    Sheets("Teams").Select    ' targets Office worksheet
    ccemail = WorksheetFunction.VLookup(cbauditor2, Range("B31:D34"), 3, 0)    'returns email address
    date2 = Worksheets("Office").Range("B3").Value    ' Date
    Worksheets("Office").Range("B2").Value = cbozone    'zone
    Sheets("Teams").Select    ' targets Office worksheet
    email2 = WorksheetFunction.VLookup(cbozone, Range("A5:D28"), 4, 0)    'returns email address
    
    email3 = Application.InputBox("Please enter Email To Automatically Email this To:")
    Worksheets("Office").Range("E4").Value = email3
        
    With OutMail
        .To = Array(email2, email3) ' who email is going to
        .CC = Array(ccemail, "joe.king@email.com")    ' facilatator
        .Subject = subject1    ' What the Email is about
        .Attachments.Add ActiveWorkbook.FullName
        .Body = ActiveSheet.Value
        .Send
    End With
    
End Sub

-- My attempt to combine a fixed range and a selected range (didn't work)

Sub EmailRange()
'Update 20131209
Dim WorkRng As Range

On Error Resume Next
xTitleId = "Select Cells"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
'Set r2 = Sheets("Sheet1").Range("C1:C6")
'Set myMultipleRange = Union(WorkRng, r2)
Application.ScreenUpdating = False
("C1:C6,WorkRng").Select
ActiveWorkbook.EnvelopeVisible = True
With ActiveSheet.MailEnvelope
    .Introduction = "CLDR Change completion"
    .Item.To = "alexander.thornhill@cls.ab.ca"
    .Item.Subject = "CLDR Change Completion NOtification"
    .Item.Send
End With
Application.ScreenUpdating = True
End Sub
