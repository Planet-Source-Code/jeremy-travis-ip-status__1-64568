VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Checker"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7890
      Top             =   840
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Height          =   495
      Left            =   510
      TabIndex        =   16
      Top             =   5460
      Width           =   1665
   End
   Begin VB.TextBox txtFields 
      Height          =   315
      Left            =   4110
      TabIndex        =   15
      Text            =   "209.68.48.118"
      Top             =   5310
      Width           =   1635
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   4080
      TabIndex        =   14
      Text            =   "Echo"
      Top             =   5790
      Width           =   1635
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Index           =   0
      Left            =   4050
      TabIndex        =   13
      Top             =   6330
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Index           =   1
      Left            =   4080
      TabIndex        =   12
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Index           =   2
      Left            =   4110
      TabIndex        =   11
      Top             =   7110
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Index           =   3
      Left            =   4110
      TabIndex        =   10
      Top             =   7530
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Index           =   4
      Left            =   4110
      TabIndex        =   9
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Index           =   5
      Left            =   4110
      TabIndex        =   8
      Top             =   8310
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3450
      TabIndex        =   2
      Top             =   4020
      Width           =   1695
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert new record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1740
      TabIndex        =   1
      Top             =   4020
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear ListView"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   4020
      Width           =   1695
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Load/Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      TabIndex        =   0
      Top             =   4020
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   30
      TabIndex        =   6
      Top             =   630
      Width           =   6825
      Begin MSComctlLib.ListView ListView3 
         Height          =   3015
         Left            =   5880
         TabIndex        =   26
         Top             =   180
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Time"
            Object.Width           =   1323
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   3450
         TabIndex        =   25
         Top             =   180
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Status"
            Object.Width           =   4233
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   90
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IP Address"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   6825
      Begin VB.Label lblDBRS 
         Alignment       =   2  'Center
         Caption         =   "sdefgsdfgsdfghsdfhsdghdfghdfghdfghdfghdfghdfghdfghdfghdfghdfghdfghdfghdfhgdfhdfghdfhdfh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6585
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "IP Address"
      Height          =   195
      Left            =   2700
      TabIndex        =   24
      Top             =   5340
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Send Packets"
      Height          =   195
      Left            =   2640
      TabIndex        =   23
      Top             =   5790
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Return Status"
      Height          =   195
      Left            =   2640
      TabIndex        =   22
      Top             =   6420
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Addredd (dec)"
      Height          =   195
      Left            =   2610
      TabIndex        =   21
      Top             =   6780
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Round Trip ms"
      Height          =   195
      Left            =   2640
      TabIndex        =   20
      Top             =   7200
      Width           =   1035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data Packet Size"
      Height          =   195
      Left            =   2610
      TabIndex        =   19
      Top             =   7590
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Data Returned"
      Height          =   195
      Left            =   2610
      TabIndex        =   18
      Top             =   7950
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Data Pointer"
      Height          =   195
      Left            =   2670
      TabIndex        =   17
      Top             =   8400
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'you have to declare the variables you use
Public ItemIndex As Integer 'index of a row in listview
Dim CurrentIP As String

Private Sub cmdPing_Click()
Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   
   If SocketsInitialize() Then
   
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
     On Error Resume Next
      success = Ping((CurrentIP), (Text2.Text), ECHO)
      
     'display the results
      Text4(0).Text = GetStatusCode(success)
      If Text4(0).Text = "ip req timed out" Then MsgBox "timeout"
      Text4(1).Text = ECHO.Address
      Text4(2).Text = ECHO.RoundTripTime & " ms"
      Text4(3).Text = ECHO.DataSize & " bytes"
      
      If Left$(ECHO.Data, 1) <> Chr$(0) Then
         pos = InStr(ECHO.Data, Chr$(0))
         Text4(4).Text = Left$(ECHO.Data, pos - 1)
      End If
   
      Text4(5).Text = ECHO.DataPointer
      
      SocketsCleanup
      
   Else
   
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
   
   End If
   End Sub

' Sample Access DAO Project - by Rob t.H. - ottooliebol@hotmail.com
'               DAO = Data Access Objects
' This example explains how to use DAO and a listview control
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
' This Application uses an Access 2000 Database. To work with this database, you
' need to use the Microsoft DAO 3.6 Object Library DLL!
' You can find it here: Menu Project, References
' --
' For Microsoft Office 2000 (Access, this example) you need version 3.6 of the DAO DLL,
' for Microsoft Office 97 (Access) you need version 3.51 of the DAO DLL.
' Version 3.6 will NOT work with Office 97 and below!!
' So with Project References, only use ONE DAO DLL!!!

Private Sub cmdRead_Click()
Timer1.Enabled = True

' if there is data in the listview, clear it first before populating it.
' we just simulate a click on the Clear ListView1 button
' (never type 2 times the same code if 1 time can also do the job :-))
If ListView1.ListItems.Count > 0 Then cmdClear_Click


' now we declare the variables we use with the database stuff.
' DB is the database itself, RS is the recordset (in this case: table People)
' --
' for this stuff to work, we need the Microsoft DAO 3.6 Object Library DLL!!
Dim DB As Database
Dim RS As Recordset

' Open database DAOtest.mdb (in the same path as the application files are)
Set DB = OpenDatabase(App.Path + "\IPlist.mdb")
' Open recordset (table: People) in database DB
Set RS = DB.OpenRecordset("IPLIST")


' If there are no records in the table: (EOF = End Of File)
If RS.EOF Then
' display text in the label (The Name Property is the name of the table (People)
    lblDBRS.Caption = "There were no records found in the database "
Else
' We found some records!!
' display in the label how many records we found
    lblDBRS.Caption = "There are " & RS.RecordCount & " sites to be monitored"

    ' Now we are going to read the records from the table.
    ' We use a loop for this, read until End Of File
    
    ' We use the integer 'a', so we know in which row we are writing (listview)
    Dim a As Integer
    a = 1
    
    Do Until RS.EOF
        
        ListView1.ListItems.Add , , RS!IPaddress
        ListView1.ListItems(a).ListSubItems.Add , , RS!Name
              
        ' Increase 'a' (our progress counter so we know in which row we are writing)
        ' with 1
        a = a + 1
        
        ' Move to the next row in the recordset
        RS.MoveNext
        
    ' Do the same stuff again, until we reached the End Of File
    Loop

End If

'we are done with the recordset and the database, so we close them now
RS.Close
DB.Close

End Sub

Private Sub cmdInsert_Click()

' Now we want to insert a new record to the table People.
' We are going to show another form to do this.
frmInsert.Show

'we have to reset ItemIndex, otherwise in some cases the ItemIndex will be remembered
ItemIndex = 0

End Sub

Private Sub cmdDelete_Click()


MsgBox "do you want to stop now?"
Timer1.Enabled = False
' the user wants to delete the selected record.
' if there are no records in the listview: exit sub
If ListView1.ListItems.Count = 0 Then Exit Sub

'if itemindex = 0 then there is nothing selected!
If ItemIndex <> 0 Then
    'ok, there is something selected!
    'now we will get the name from the items name, and ask for delete confirmation
    Dim Ask As String
    Ask = MsgBox("Are you sure that you want to delete '" & ListView1.ListItems.Item(ItemIndex).Text & "'?", vbYesNo + vbInformation, "Delete record")
        
    If Ask = vbYes Then
        'user had pressed yes, please delete ;-)
        'now we have to get the ID of the row in de database (item's tag!) and delete record from table
    
        Dim DB As Database
        Dim RS As Recordset
        
        ' Open database DAOtest.mdb (in the same path as the application files are)
        Set DB = OpenDatabase(App.Path + "\IPlist.mdb")
        ' Open recordset (table: People) in database DB
        Set RS = DB.OpenRecordset("IPlist")
        
        'we move the recordpointer to the first record, this way we can seek the whole table
        RS.MoveFirst

    
        If RS.NoMatch Then
            'if there was no match (the ID couldn't be found in the table)
            MsgBox ("The record can't be found in the table!"), vbOKOnly + vbCritical, "Error"
            
            
            'we reset itemindex to 0, this means that nothing is selected in listview1
            ItemIndex = 0
            
            'close recordset and database
            RS.Close
            DB.Close
            
            Exit Sub
        Else
            'there was a match! we will now delete the record from the database
            RS.Delete
            
            'and we will delete the row from the listview
            ListView1.ListItems.Remove (ItemIndex)
        End If
           
        'close recordset and database
        RS.Close
        DB.Close
    
    End If

'we reset itemindex to 0, this means that nothing is selected in listview1
ItemIndex = 0
    
End If

Timer1.Enabled = True
            
End Sub

Private Sub cmdClear_Click()
'clear the listview
ListView1.ListItems.Clear

'we have to reset ItemIndex, otherwise in some cases the ItemIndex will be remembered
ItemIndex = 0

'set caption to text no table opened
lblDBRS.Caption = "No table opened at the moment"
'set form caption
Form1.Caption = "DAO Example"


End Sub

Private Sub Form_Load()
'build the listviews column headers
Dim Header As ColumnHeader

    
    
' Set the caption of the label
lblDBRS.Caption = "No table opened at the moment"

'we have to reset ItemIndex, so there is nothing selected in listview1
ItemIndex = 0

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
ItemIndex = Item.Index

End Sub

Private Sub Timer1_Timer()
Dim x As Integer
Dim y As Integer

    'cmdRead_Click 'This reloads the DB every loop
    'ListView2.ListItems.Clear
    'ListView3.ListItems.Clear
    x = 1
    y = Me.ListView1.ListItems.Count
    
    Do Until x = y + 1
        ListView1.SelectedItem = ListView1.ListItems(x)
        CurrentIP = ListView1.SelectedItem
        cmdPing_Click
        'ListView2.ListItems.Add , , Text4(0).Text
        'ListView3.ListItems.Add , , Text4(2).Text
        
'DECLARE A VARIABLE TO ADD LISTITEM OBJECTS.
Dim itmX As ListItem
'I USE YOUR SELECTED ITEM TO POINT TO THE PROPER SUB-ITEMS
Set itmX = ListView1.SelectedItem
itmX.SubItems(2) = Text4(0).Text
itmX.SubItems(3) = Text4(2).Text

        x = x + 1
    Loop

End Sub
