VERSION 5.00
Begin VB.Form frmInsert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert new location"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsert.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3645
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtNewIp 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
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
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         Left            =   90
         TabIndex        =   3
         Top             =   630
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
' The user wants to cancel, so we simply unload this form from memory
Unload Me

End Sub

Private Sub cmdInsert_Click()
' OK, the user wants to insert the data to the table People
' But before we do that, we want to check if everything is filled in!

If txtNewIp.Text = "" Then MsgBox ("Please enter an IP address!"), vbOKOnly + vbCritical, "Error" _
: Exit Sub

If txtName.Text = "" Then MsgBox ("Please fill in the name!"), vbOKOnly + vbCritical, "Error" _
: Exit Sub


' OK, everything seems ok! Now we open a connection to the database

' Open database DAOtest.mdb (in the same path as the application files are)
Set DB = OpenDatabase(App.Path + "\IPLIST.mdb")
' Open recordset (table: People) in database DB
Set RS = DB.OpenRecordset("IPLIST")

'Now we are going to add a new record to the recordset! :-)
With RS 'use recordset
    
    .AddNew 'add new record
    
    !IPaddress = txtNewIp.Text 'field name must contain content of txtname
    !Name = txtName.Text 'field address must contain content of txtaddress
    
    .Update 'As soon as we use RS.Update, then the data is written to the table.

End With

' Reload the list to activate new entery
' if there is data in the listview, clear it first before populating it.
' we just simulate a click on the Clear ListView1 button
' (never type 2 times the same code if 1 time can also do the job :-))
If Form1.ListView1.ListItems.Count > 0 Then Form1.ListView1.ListItems.Clear


' now we declare the variables we use with the database stuff.
' DB is the database itself, RS is the recordset (in this case: table People)
' --
' for this stuff to work, we need the Microsoft DAO 3.6 Object Library DLL!!


' Open database DAOtest.mdb (in the same path as the application files are)
Set DB = OpenDatabase(App.Path + "\IPlist.mdb")
' Open recordset (table: People) in database DB
Set RS = DB.OpenRecordset("IPLIST")


' If there are no records in the table: (EOF = End Of File)
If RS.EOF Then
' display text in the label (The Name Property is the name of the table (People)
    Form1.lblDBRS.Caption = "There were no records found in the database "
Else
' We found some records!!
' display in the label how many records we found
    Form1.lblDBRS.Caption = "There are " & RS.RecordCount & " sites to be monitored"

    ' Now we are going to read the records from the table.
    ' We use a loop for this, read until End Of File
    
    ' We use the integer 'a', so we know in which row we are writing (listview)
    Dim a As Integer
    a = 1
    
    Do Until RS.EOF
        
        Form1.ListView1.ListItems.Add , , RS!IPaddress
        Form1.ListView1.ListItems(a).ListSubItems.Add , , RS!Name
        '
        
        ' While we are copying, we use a progressbar (pb) to show the progress
        Form1.pb.Value = Int(RS.PercentPosition)
               
        ' Increase 'a' (our progress counter so we know in which row we are writing)
        ' with 1
        a = a + 1
        
        ' Move to the next row in the recordset
        RS.MoveNext
        
    ' Do the same stuff again, until we reached the End Of File
    Loop

' Ok, we completed reading from the table, now we reset the progressbar
Form1.pb.Value = 0

End If

'we are done with the recordset and the database, so we close them now
RS.Close
DB.Close

' and now close the form
Unload Me

End Sub



Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
' we want the user to enter a phone number, so no other characters then numbers are
' allowed! The ascii range for numbers is 47..57.
' If a user presses a key different then a number, we just not write it in the textbox!
' Only exception is Ascii Char 8, this is backspace :-)

If KeyAscii <> 8 And KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0: Exit Sub

End Sub
