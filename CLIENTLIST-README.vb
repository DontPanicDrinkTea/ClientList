'CLIENT LIST
Option Explicit

Public dateval As String, myear As String, hlpath As String

Public ia_firststr As String, ia_emailstr As String, cc_emailstr As String

Public ctable As Range, cnum As Range, dfolder As Range, lname1 As Range, _
fname1 As Range, lname2 As Range, fname2 As Range, mdate As Range, _
mtime As Range, c_ia1 As Range, c_ia2 As Range, conf As Range, _
mtype As Range, c_email As Range, c_print As Range, c_draft As Range, _
c_edit As Range, c_approve As Range, c_sdrive As Range, c_meth As Range, _
c_sent As Range, c_location As Range, c_notes As Range, c_prov As Range, c_city As Range

Public Const cnum_col = 1
Public Const dfolder_col = 2
Public Const lname1_col = 3
Public Const fname1_col = 4
Public Const lname2_col = 5
Public Const fname2_col = 6
Public Const mdate_col = 7
Public Const mtime_col = 8
Public Const c_ia1_col = 9
Public Const c_ia2_col = 10
Public Const conf_col = 11
Public Const mtype_col = 12
Public Const c_email_col = 13
Public Const c_print_col = 14
Public Const c_draft_col = 15
Public Const c_edit_col = 16
Public Const c_approve_col = 17
Public Const c_sdrive_col = 18
Public Const c_meth_col = 19
Public Const c_sent_col = 20
Public Const c_location_col = 21
Public Const c_notes_col = 22
Public Const c_prov_col = 23
Public Const c_city_col = 24

Public iatable As Range, ia_ia As Range, ia_first As Range, ia_email As Range, cc_email As Range

Public Const ia_ia_col = 4
Public Const ia_first_col = 2
Public Const ia_email_col = 5
Public Const cc_email_col = 6

Public c_ws As Worksheet, ia_ws As Worksheet

Public c_rownum As Integer, ia_rownum As Integer

Public i As Integer, j As Integer, k As Integer

Public lastnames As String, fullnames As String

Public folder_path As String, conf_path As String, tocpdf_path As String, tocxl_path As String, _
reportpdf_path As String, reportdoc_path As String

Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) _
As Long



Function FolderCreate(ByVal path As String) As Boolean
    FolderCreate = True
    Dim fso As New FileSystemObject

    If FolderExists(path) Then
        Exit Function
    Else
        On Error GoTo DeadInTheWater
        fso.CreateFolder path ' could there be any error with this, like if the path is really screwed up?
        Exit Function
    End If

DeadInTheWater:
        MsgBox "A folder could not be created for the following path: " & path & ". Check the path name and try again."
        FolderCreate = False
        Exit Function
End Function

Function FolderExists(ByVal path As String) As Boolean
    FolderExists = False
    Dim fso As New FileSystemObject

    If fso.FolderExists(path) Then FolderExists = True
End Function

Public Sub PrintFile(ByVal strPathAndFilename As String)
    Call apiShellExecute(Application.hwnd, "print", strPathAndFilename, vbNullString, vbNullString, 0)
End Sub

Public Sub InitializeTables()
    Set c_ws = ThisWorkbook.Worksheets("Clients")
    Set ia_ws = ThisWorkbook.Worksheets("Advisors")
    Set ctable = c_ws.Range("Table6")
    Set iatable = ia_ws.Range("Table5")

    c_rownum = ctable.Rows.Count
    ia_rownum = iatable.Rows.Count
End Sub
Function c_VCIndex(colnum As Integer) As Integer()
    Dim r As Range
    Dim splitr As Range
    Dim c As Range
    Dim rnum As Integer
    Dim ccount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim arr() As Integer
    
    Set r = ctable.Columns(colnum)
    If r.SpecialCells(xlCellTypeVisible).Count > 0 Then
        Set splitr = r.SpecialCells(xlCellTypeVisible)
        ccount = splitr.Count
        ReDim Preserve arr(1 To ccount)
        i = 1
        For Each c In splitr
            arr(i) = c.Row - 1
            i = i + 1
        Next c
        c_VCIndex = arr()
    ElseIf r.SpecialCells(xlCellTypeVisible).Count = 0 Then
        MsgBox "No Records Available"
        End
    End If
End Function
Function ia_VCIndex(colnum As Integer) As Integer()
    Dim r As Range
    Dim splitr As Range
    Dim c As Range
    Dim rnum As Integer
    Dim ccount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim arr() As Integer
    
    Set r = iatable.Columns(colnum)
    If r.SpecialCells(xlCellTypeVisible).Count > 0 Then
        Set splitr = r.SpecialCells(xlCellTypeVisible)
        ccount = splitr.Count
        ReDim Preserve arr(1 To ccount)
        i = 1
        For Each c In splitr
            arr(i) = c.Row - 1
            i = i + 1
        Next c
        ia_VCIndex = arr()
    ElseIf r.SpecialCells(xlCellTypeVisible).Count = 0 Then
        MsgBox "No Advisor Information Available"
        End
    End If
End Function
Sub VCIndexTest()
    Dim r As Range
    Dim splitr As Range
    Dim c As Range
    Dim rnum As Integer
    Dim ccount As Integer
    Dim i As Integer
    Dim j As Integer

    Call InitializeTables
    If ctable.Rows("1:" & c_rownum).EntireRow.Hidden = True Then
        MsgBox "All rows are hidden"
    ElseIf ctable.Rows("1:" & c_rownum).EntireRow.Hidden = False Then
        MsgBox "All rows are visible"
    End If
    
    If ctable.Rows("1:" & c_rownum).EntireRow.Hidden = False Then
        Set r = ctable.Columns(11)
        r.Select
    End If
'        Set splitr = r.SpecialCells(xlCellTypeVisible)
'        ccount = splitr.Count
'        ReDim Preserve arr(1 To ccount)
'        i = 1
'        For Each c In splitr
'            arr(i) = c.Row
'            i = i + 1
'        Next c
'
'        VCIndex = arr()
'    ElseIf ctable.Rows("1:" & c_rownum).EntireRow.Hidden = True Then
'        MsgBox "No Records Available"
'        End
'    End If
End Sub

Public Sub InitializeClients(i)
    Call InitializeTables

    Set cnum = ctable(i, cnum_col)
    Set dfolder = ctable(i, dfolder_col)
    Set lname1 = ctable(i, lname1_col)
    Set fname1 = ctable(i, fname1_col)
    Set lname2 = ctable(i, lname2_col)
    Set fname2 = ctable(i, fname2_col)
    Set mdate = ctable(i, mdate_col)
    Set mtime = ctable(i, mtime_col)
    Set c_ia1 = ctable(i, c_ia1_col)
    Set c_ia2 = ctable(i, c_ia2_col)
    Set conf = ctable(i, conf_col)
    Set mtype = ctable(i, mtype_col)
    Set c_email = ctable(i, c_email_col)
    Set c_print = ctable(i, c_print_col)
    Set c_draft = ctable(i, c_draft_col)
    Set c_edit = ctable(i, c_edit_col)
    Set c_approve = ctable(i, c_approve_col)
    Set c_sdrive = ctable(i, c_sdrive_col)
    Set c_meth = ctable(i, c_meth_col)
    Set c_sent = ctable(i, c_sent_col)
    Set c_location = ctable(i, c_location_col)
    Set c_notes = ctable(i, c_notes_col)
    Set c_prov = ctable(i, c_prov_col)
    Set c_city = ctable(i, c_city_col)

    dateval = mdate.Value

    If dateval = "" Then
        myear = "Not Confirmed"
    Else
        myear = CStr(year(dateval))
    End If
    
    If lname2.Value <> "" And lname2.Value <> lname1.Value Then
        lastnames = lname1.Value & "," & lname2.Value
        fullnames = fname1.Value & " " & lname1.Value & " and " & fname2.Value & " " & lname2.Value
    ElseIf lname2.Value <> "" And lname2.Value = lname1.Value Then
        lastnames = lname1.Value
        fullnames = fname1.Value & " and " & fname2 & " " & lname1
    ElseIf lname2 = "" Then
        lastnames = lname1.Value
        fullnames = fname1.Value & " " & lname1.Value
    End If
    
    folder_path = "Z:\wrkgrp80\GR_WILL_ESTATE\SAvery\Client Files\" & myear & "\" & cnum & " - " & lastnames
    conf_path = folder_path & "\" & lastnames & "Conf.pdf"
    tocpdf_path = folder_path & "\" & lastnames & "TOCNEW.pdf"
    tocxl_path = folder_path & "\" & lastnames & "TOCNEW.xlsm"
    reportpdf_path = folder_path & "\" & lastnames & "Report - FinalNEW.pdf"
    reportdoc_path = folder_path & "\" & lastnames & "Report - FinalNEW.docx"

End Sub
Sub SetIA(ianame)
    Dim x As Integer
    Dim y() As Integer
    Call InitializeTables
    
    ia_ws.ListObjects("Table5").Range.AutoFilter Field:=ia_ia_col, Criteria1:=ianame
    y = ia_VCIndex(4)
    For x = 1 To UBound(y)
        Set ia_ia = iatable(y(x), ia_ia_col)
        Set ia_first = iatable(y(x), ia_first_col)
        Set ia_email = iatable(y(x), ia_email_col)
        Set cc_email = iatable(y(x), cc_email_col)
        
        ia_firststr = ia_first.Value
        ia_emailstr = ia_email.Value
        cc_emailstr = cc_email.Value
        If ia_emailstr = "" Then
            MsgBox "You are missing email for " & c_ia1
        End If
    Next x
End Sub


Sub SetFilters(colnum)
    'Check to see if filter is applied, and if so remove it
    If c_ws.Range("Table6").ListObject.ShowAutoFilter Then
        c_ws.ListObjects("Table6").AutoFilter.ShowAllData
    End If
    
    c_ws.ListObjects("Table6").Range.AutoFilter Field:=7, Criteria1:= _
        "<>"
    c_ws.ListObjects("Table6").Range.AutoFilter Field:=colnum, Criteria1:="bw"
End Sub

Sub ClearFilters()

    If c_ws.Range("Table6").ListObject.ShowAutoFilter Then
        c_ws.ListObjects("Table6").AutoFilter.ShowAllData
    End If
End Sub

Sub hyperlinks()
    Dim ai As Integer
    Dim bi As Integer
    Dim vcells() As Integer
    
    Application.ScreenUpdating = False
    
    Call InitializeTables
    Call SetFilters(dfolder_col)
    vcells = c_VCIndex(dfolder_col)
    
    For ai = 1 To UBound(vcells)
        Call InitializeClients(vcells(ai))
        dfolder.Value = "here"
''      Loop through date in table, if there is no hyperlink, add it
        If dfolder.hyperlinks.Count = 0 Then
            If Not FolderExists(hlpath) Then
               'client file doesn't exist, so create full path
                FolderCreate folder_path
            End If
            c_ws.hyperlinks.Add Anchor:=dfolder, _
            Address:=folder_path
 ''     If there is already a hyperlink, make sure it is correct
        ElseIf dfolder.hyperlinks.Count >= 1 Then
            dfolder.hyperlinks(1).Address = _
            folder_path
        End If
    Next ai

    Application.ScreenUpdating = True
    Call ClearFilters

End Sub

Sub SendConfirmation()
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim ai As Integer
    Dim bi As Integer
    Dim vcells() As Integer
  
    Application.ScreenUpdating = False
    
    Call InitializeTables
    Call SetFilters(conf_col)
    vcells = c_VCIndex(conf_col)
    
    'email to send
    For ai = 1 To UBound(vcells)
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
        Call InitializeClients(vcells(ai))
        Call SetIA(c_ia1.Value)
        With OutMail
            .To = ia_emailstr
            .CC = cc_emailstr & "; sharon.avery@rbc.com"
            .BCC = ""
            .Subject = lastnames & " Meeting Confirmation with Sharon Avery"
            .HTMLBody = "<p><font face=""Calibri"" size=""3"">Hello " & ia_firststr & ",<br><br>" _
                & "I can confirm that Sharon Avery is scheduled to meet with <b>" & fullnames & "</b> on <b>" & Format(mdate.Value, "long date") & "</b> at <b>" & Format(mtime.Value, "h:mm AM/PM") & "</b>.<br><br>" _
                & "In preparation for this meeting, I have attached: <br><br>" _
                & "• A list of suggested information and materials for your clients to obtain and bring for this meeting (if they have the items on hand);<br>" _
                & "• A list of suggested information for your clients to obtain and bring to the meeting if they have a private company and/or family trust;<br>" _
                & "• A copy of our WEC Confirmation Form to be signed during the meeting; and<br>" _
                & "• A Preliminary Information form (to be filled out by you). - <b>At a minimum, we need their HHID, names, date of birth, mailing address, and assets held with you.</b><br><br>" _
                & "Please gather as much information as you can and send it to me prior to the meeting date for Sharon to review. If your client has had a Compass Financial Plan prepared for them in the past, please send us a copy of this as well.<br><br>" _
                & "If you have any questions or concerns, please feel free to contact myself or Sharon.<br><br>" _
                & "Thank you very much for your help,<br><br><br></font>" _
                & "<font face=""Calibri"" size=""3"" font colour = ""blue""><b>Eva Smith</b> | Regional Administrative Coordinator | Atlantic Region | Wealth Management Services I Tel: (902) 494-5699 | <b>RBC Dominion Securities</b></font></p>"
            .Attachments.Add ("C:\Users\171856123\Documents\Admin\Confirmation Documents\Corporate Info for Meeting.pdf")
            .Attachments.Add ("C:\Users\171856123\Documents\Admin\Confirmation Documents\family inventoryenglish.pdf")
            .Attachments.Add ("C:\Users\171856123\Documents\Admin\Confirmation Documents\Household ID Explanation.pdf")
            .Attachments.Add ("C:\Users\171856123\Documents\Admin\Confirmation Documents\Sharon Avery Bio.pdf")
            .Attachments.Add ("C:\Users\171856123\Documents\Admin\Confirmation Documents\WEC Confirmation Form.pdf")
            .Attachments.Add ("C:\Users\171856123\Documents\Admin\Confirmation Documents\WEC Preparation List.pdf")
            .Attachments.Add ("C:\Users\171856123\Documents\Admin\Confirmation Documents\WEC_PreInfo_WM.pdf")
            .Display
        End With
        Set OutMail = Nothing
        Set OutApp = Nothing
        conf.Value = Date
    Next ai
            
        On Error GoTo 0
            
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With

    Call ClearFilters
    
    Application.ScreenUpdating = True

End Sub
Sub EmailReport()
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim ai As Integer
    Dim bi As Integer
    Dim vcells() As Integer
  
    Application.ScreenUpdating = False
    
    Call InitializeTables
    Call SetFilters(c_email_col)
    vcells = c_VCIndex(c_email_col)
    
    'email to send
    For ai = 1 To UBound(vcells)
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
        Call InitializeClients(vcells(ai))
        Call SetIA(c_ia1.Value)
        With OutMail
            .To = ia_emailstr
            .CC = cc_emailstr & "; sharon.avery@rbc.com"
            .BCC = ""
            .Subject = lastnames & " Will & Estate Consultation Report"
            .HTMLBody = "<p><font face=""Calibri"" size=""3"">Hello " & ia_firststr & ",<br><br>" _
                & "I have attached a copy of Sharon’s Will & Estate Planning Report generated from her meeting with <b>" & fullnames & "</b> for your review. If you notice anything in this report that needs to be changed, please let me know within a few days, and I will make any necessary edits for you.<br>" _
                & "If I do not hear from you within two business days, I will send you a package that includes a physical copy of this report for your records and a folder for your client, which contains a colour copy of the report along with any relevant support materials identified by Sharon.  You will receive these materials within seven business days via inter-office mail.  Please contact me directly if, for any reason, you do not receive the report package within this timeframe.<br>" _
                & "Thank you,<br><br><br>" _
                & "<b>Eva Smith</b> | Regional Administrative Coordinator | Atlantic Region | Wealth Management Services I Tel: (902) 494-5699 | <b>RBC Dominion Securities</b></font></p>"
            .Attachments.Add (reportpdf_path)
            .Attachments.Add (conf_path)
            .Attachments.Add (tocpdf_path)
            .Display
        End With
        Set OutMail = Nothing
        Set OutApp = Nothing
        conf.Value = Date
    Next ai
            
        On Error GoTo 0
            
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With

    Call ClearFilters
    
    Application.ScreenUpdating = True

End Sub

Sub FinalizeAndPrint()
    Dim objWord
    Dim objDoc
    Dim ai As Integer
    Dim bi As Integer
    Dim vcells() As Integer
    Dim cprint As String
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Call InitializeTables
    Call SetFilters(c_print_col)
    vcells = c_VCIndex(c_print_col)
    
    Set objWord = CreateObject("Word.Application")

    For ai = 1 To UBound(vcells)
        cprint = ""
        Call InitializeClients(vcells(ai))
        cprint = reportpdf_path & "," & conf_path & "," & tocpdf_path
        Set objDoc = objWord.Documents.Open(reportdoc_path)
        
        objWord.Visible = True
        objDoc.Application.Run "WEC1.GetHeadings", c_prov.Value, c_city.Value, conf_path, tocpdf_path, tocxl_path, reportpdf_path, cprint
    Next ai
    Call ClearFilters

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub CheckupEmail()
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim ai As Integer
    Dim bi As Integer
    Dim vcells() As Integer
  
    Application.ScreenUpdating = False
    
    Call InitializeTables
    Call SetFilters(c_print_col)
    vcells = c_VCIndex(c_print_col)
    
    'email to send
    For ai = 1 To UBound(vcells)
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
        Call InitializeClients(vcells(ai))
        Call SetIA(c_ia1.Value)
        With OutMail
            .To = ia_emailstr
            .CC = cc_emailstr & "; sharon.avery@rbc.com; melissa.picheca@rbc.com"
            .BCC = ""
            .Subject = lastnames & " Sent Will & Estate Consultation Report"
            .HTMLBody = "<p><font face=""Calibri"" size=""3"">Hello " & ia_firststr & ",<br><br>" _
                & "You may have already received this, but I wanted to let you know that I have sent you a re-dated colour copy of Sharon’s Will & Estate Planning Report generated from her meeting with <b>" & fullnames & "</b>. Please confirm whether this has been received (if you haven't already). If this report has not yet been received, please let me know as soon as possible, and I will rush another colour copy to you.<br>" _
                & "Thank you very much for your help,<br><br><br>" _
                & "<b>Eva Smith</b> | Regional Administrative Coordinator | Atlantic Region | Wealth Management Services I Tel: (902) 494-5699 | <b>RBC Dominion Securities</b></font></p>"
            .Attachments.Add (reportpdf_path)
            .Attachments.Add (conf_path)
            .Attachments.Add (tocpdf_path)
            .Display
        End With
        Set OutMail = Nothing
        Set OutApp = Nothing
        conf.Value = Date
    Next ai
            
        On Error GoTo 0
            
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With

    Call ClearFilters
    
    Application.ScreenUpdating = True

End Sub
Sub PrintJustThisLetter()
    Dim ai As Integer
    Dim bi As Integer
    Dim vcells() As Integer
    Dim cprint As String
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Call InitializeTables
    Call SetFilters(c_print_col)
    vcells = c_VCIndex(c_print_col)

    For ai = 1 To UBound(vcells)
        Call InitializeClients(vcells(ai))
        PrintFile (reportpdf_path)
        Application.Wait (Now + TimeValue("0:00:10"))
        PrintFile (tocpdf_path)
        Application.Wait (Now + TimeValue("0:00:10"))
        PrintFile (conf_path)
        Application.Wait (Now + TimeValue("0:00:10"))
    Next ai
    
End Sub


