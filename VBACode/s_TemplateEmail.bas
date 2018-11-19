Attribute VB_Name = "E_NameforInitialEmail"

Sub Email_NameforInitialEmail()

'Macro Setup
Application.ScreenUpdating = False
wsCurrent = ActiveSheet.Name

'Potential User Inputs
MasterWorksheet = "ExProjectTracker"
KeyHeaderTitle = "ServerName"
KeyHeaderColumnSetup = "A"
Column_EmailName = "First Email Recipients"
EmailIdentifier = "Enter note that email was sent here"
EmailSubjectText = "Email Subject Here"
EmailManager = "EmailManager"
FolderName = "Draft_FolderNameHere"
SentfromOther = "TeamNameHere"

'Clear Variables
n = 0

'Obtain Table Information
Worksheets(MasterWorksheet).Activate
Range("A2").Select
On Error Resume Next
ActiveSheet.ShowAllData
On Error GoTo 0
KeyHeaderColumn = Range(KeyHeaderColumnSetup & ":" & KeyHeaderColumnSetup).Address
HeaderRowSetup = WorksheetFunction.Match(KeyHeaderTitle, Range(KeyHeaderColumn), 0)
HeaderRowRange = Range(HeaderRowSetup & ":" & HeaderRowSetup).Address
TableStart = Cells(HeaderRowSetup, KeyHeaderColumnSetup).Address
LastRowofTable = Cells(Range(HeaderRowRange).Row, KeyHeaderColumnSetup).End(xlDown).Row 'This will only work if the key column is filled in every row. If not, use LastRowofTable = "manualcountoflastrow"
On Error GoTo NextRange
    With ActiveSheet
        LastFilterRow = .Range(Split(.AutoFilter.Range.Address, ":")(1)).Row
    End With
    LastRowofTable = WorksheetFunction.Max(LastRowofTable, LastFilterRow)
On Error GoTo 0
NextRange:
LastColumnofTable = Cells(HeaderRowSetup, Columns.Count).End(xlToLeft).Column
TableEnd = Cells(LastRowofTable, LastColumnofTable).Address
MasterWorksheetRange = Range(TableStart, TableEnd).Address
TableStartRowExtract = Range(TableStart).Row
TableEndRowExtract = Range(TableEnd).Row
TableStartColumnExtract = Range(TableStart).Column
TableEndColumnExtract = Range(TableEnd).Column
HeaderRowRangeExtract = TableStartColumnExtract & ":" & TableEndColumnExtract
HeaderRangewithRowandColumns = Range(Cells(HeaderRowSetup, KeyHeaderColumnSetup), Cells(HeaderRowSetup, TableEndColumnExtract)).Address
TotalMasterRows = TableEndRowExtract - TableStartRowExtract
TotalMasterColumns = TableEndColumnExtract - TableStartColumnExtract + 1

'Find These Columns
Column_Servers = Cells(, WorksheetFunction.Match("ServerName", Range(HeaderRowRange), 0)).Column
Column_FirstContact = Cells(, WorksheetFunction.Match("Contact1", Range(HeaderRowRange), 0)).Column
Column_SecondContact = Cells(, WorksheetFunction.Match("Contact2", Range(HeaderRowRange), 0)).Column
Column_ThirdContact = Cells(, WorksheetFunction.Match("Contact3", Range(HeaderRowRange), 0)).Column
Column_FourthContact = Cells(, WorksheetFunction.Match("Contact4", Range(HeaderRowRange), 0)).Column
Column_LastContact = Cells(, WorksheetFunction.Match("Contact5", Range(HeaderRowRange), 0)).Column
Column_Notes = Cells(, WorksheetFunction.Match("Notes", Range(HeaderRowRange), 0)).Column

'Find These Ranges
Range_Servers = Range(Cells(HeaderRowSetup + 1, Column_Servers), Cells(LastRowofTable - 2, Column_Servers)).Address                              '*May want to fix Lastrowoftable later

'Preparing AllServersTouched Array
Dim AllServersTouched() As Variant
ReDim AllServersTouched(0 To 5000)

'Build unique name array
i = 0
Dim ContactArray() As Variant
ReDim ContactArray(0 To 5000)
PossibleContacts = Range(Cells(HeaderRowSetup + 1, Column_FirstContact), Cells(LastRowofTable, Column_LastContact)).Address
For Each Cell In Range(PossibleContacts)
    If WorksheetFunction.Substitute(WorksheetFunction.Trim(Cell.Value), Chr(160), "") <> "" Then
        If IsInArray = Not IsError(Application.Match(WorksheetFunction.Substitute(WorksheetFunction.Trim(Cell.Value), Chr(160), ""), ContactArray, 0)) = True Then
            ContactArray(i) = WorksheetFunction.Substitute(WorksheetFunction.Trim(Cell.Value), Chr(160), "")
            i = i + 1
        End If
    End If
Next Cell

'If there are no contacts in the array, exit the sub
If i = 0 Then
    MsgBox ("There are no contacts in this array, exiting sub.")
    Exit Sub
Else
    ReDim Preserve ContactArray(0 To i - 1)
End If

'For each person in the array, one at a time find all servers related to them, append them to the table, and continue
i = 0
Dim ServersOwned() As Variant
ReDim ServersOwned(0 To 5000)
For Each Contact In ContactArray
    For Each Server In Range(Range_Servers)
        Row_Server = WorksheetFunction.Match(Server, Range(KeyHeaderColumn), 0)
        ConditionalAdd = "Not Needed" 'WorksheetFunction.VLookup(Server, Range(MasterWorksheetRange), WorksheetFunction.Match("Overall Status", Range(HeaderRowRange), 0), 0)
        If ConditionalAdd = "Not Needed" Then
            If Contact = Trim(Cells(Row_Server, Column_FirstContact).Value) Or Contact = Trim(Cells(Row_Server, Column_SecondContact).Value) Or Contact = Trim(Cells(Row_Server, Column_ThirdContact).Value) Or Contact = Trim(Cells(Row_Server, Column_FourthContact).Value) Or Contact = Trim(Cells(Row_Server, Column_LastContact).Value) Then
                ServersOwned(i) = Server
                i = i + 1
            End If
        End If
    Next Server
    If i = 0 Then
        GoTo NextContact:
    End If
    ReDim Preserve ServersOwned(0 To i - 1)
    
    'Loop through each server to find information and append table data per that server
    For Each Server In ServersOwned
        ServerName = WorksheetFunction.VLookup(Server, Range(MasterWorksheetRange), WorksheetFunction.Match("ServerName", Range(HeaderRowRange), 0), 0)
        ServerLocation = WorksheetFunction.VLookup(Server, Range(MasterWorksheetRange), WorksheetFunction.Match("Server Location", Range(HeaderRowRange), 0), 0)
        VirtualorPhysical = WorksheetFunction.VLookup(Server, Range(MasterWorksheetRange), WorksheetFunction.Match("Virtual_Physical", Range(HeaderRowRange), 0), 0)
        On Error Resume Next
        FirstName = Left(Contact, WorksheetFunction.Search(" ", Contact) - 1)
        On Error GoTo 0
        If FirstName = Empty Then
            FirstName = ","
            Else
            FirstName = " " & FirstName & ", "
        End If
        'Update Recipient List And First Run POCs
        Application.Calculation = xlCalculationManual
        RecipientListCell = Worksheets(EmailManager).Cells(WorksheetFunction.Match(Server, Worksheets(EmailManager).Range("A:A"), 0), WorksheetFunction.Match(Column_EmailName, Worksheets(EmailManager).Range("2:2"), 0)).Address
        FirstRunPOCCell = Worksheets(EmailManager).Cells(WorksheetFunction.Match(Server, Worksheets(EmailManager).Range("A:A"), 0), WorksheetFunction.Match("First Run POCs", Worksheets(EmailManager).Range("2:2"), 0)).Address
        If Worksheets(EmailManager).Range(RecipientListCell).Value = "" Then
            Worksheets(EmailManager).Range(RecipientListCell).Value = Contact & "; "
            Worksheets(EmailManager).Range(FirstRunPOCCell).Value = Worksheets(EmailManager).Range(RecipientListCell).Value
        Else
            Worksheets(EmailManager).Range(RecipientListCell).Value = Worksheets(EmailManager).Range(RecipientListCell).Value & Contact & "; "
            Worksheets("EmailManager").Range(FirstRunPOCCell).Value = Worksheets(EmailManager).Range(RecipientListCell).Value
        End If
        
        'Add to AllServersTouched Array if Unique
        If IsInArray = Not IsError(Application.Match(Server, AllServersTouched, 0)) = True Then
            AllServersTouched(n) = Server
            n = n + 1
        End If
        
        'Save Table String
        TableDataString = "<tr align=center> <td>" & ServerName & "</td> <td>" & ServerLocation & "</td> <td>" & VirtualorPhysical & "</td></tr>" & TableDataString
    Next Server
    Application.Calculation = xlCalculationAutomatic

    i = 0
    'Creating Email
    Dim aOutlook As Outlook.Application
    Dim aEmail As Outlook.MailItem
    Set aOutlook = New Outlook.Application
    Set aEmail = CreateItem(olMailItem)
    With aEmail
        .BodyFormat = olFormatHTML
        .To = Contact
        .CC = "ccname1@here.com; ccname2@here.com"
        .Subject = EmailSubjectText
        .HTMLBody = "<font size=3.5>" & "Hello" & FirstName & "<br><br>" & _
            "We are contacting you because of a,b,c. <br> <br>" & _
            "Table below with information from the tracker: <br> <br>" & _
            "<head><style>table, th, td {border: 1px solid black; border-collapse: collapse;}</style></head><body>" & _
            "<table width=50% border=1 border collapse=collapse> <font size =3>" & _
                "<tr bgcolor=#c00000 align=center & font color=White> <b>" & "<td>" & "Server Name" & "</td>" & "<td>" & "Server Location" & "</td>" & "<td>" & "Virtual or Physical" & "</td> </b> </tr>" & _
                "" & TableDataString & _
            "</table> <br>" & _
            "Example formatting if needed, <b> bold text, </b>, <u> underlined text </u> . <br> <br>" & _
            "Example link if needed: <a href='www.google.com'> Google </a> <br> <br>" & _
            "Thank you for your support, <br> <br>" & _
            "Signature here <br> <br>" & _
        .HTMLBody
        .Save
        .Close (olSave)
        On Error GoTo CreateFolder
ResumeMove:
        .Move GetNamespace("mapi").GetDefaultFolder(olFolderDrafts).Folders(FolderName)
        On Error GoTo 0
        
    End With

NextContact:
i = 0 'Used to be i = i + 1
TableDataString = ""
ReDim ServersOwned(0 To 5000)
Next Contact

'Save to Notes
    ReDim Preserve AllServersTouched(0 To n - 1)
    For Each ServerTouched In AllServersTouched
        Row_ServerTouched = WorksheetFunction.Match(ServerTouched, Range(KeyHeaderColumn), 0)
        NotesCell = Cells(Row_ServerTouched, Column_Notes).Address
        If Range(NotesCell).Value Like "*" & EmailIdentifier & "*" Then
            NotesSplit = Split(Range(NotesCell).Value, Chr(10))
            For Each Line In NotesSplit
                If Line Like "*" & EmailIdentifier & "*" Then
                    DateofEmailSent = Left(Line, WorksheetFunction.Search(": ", Line) - 1)
                    If CDate(DateofEmailSent) = Date Then
                        UpdateAgain = False
                    Else
                        UpdateAgain = True
                    End If
                End If
            Next Line
            Else
                FirstUpdate = True
        End If
        If FirstUpdate = True Then
            Range(NotesCell).Value = Date & ": " & EmailIdentifier & Chr(10) & Range(NotesCell).Value
        End If
        If UpdateAgain = True Then
            Range(NotesCell).Value = Date & ": " & EmailIdentifier & " Again" & Chr(10) & Range(NotesCell).Value
        End If
        UpdateAgain = Empty
        FirstUpdate = Empty
    Next ServerTouched
    n = 0
    ReDim AllServersTouched(0 To 5000)

Worksheets(wsCurrent).Activate
Application.ScreenUpdating = True

Exit Sub

'Troubleshooters
CreateFolder:
GetNamespace("mapi").GetDefaultFolder(olFolderDrafts).Folders.Add (FolderName)
MsgBox ("Drafts folder created, expand drafts on the left side of Outlook to see the drafted emails.")
On Error GoTo 0
GoTo ResumeMove

End Sub

