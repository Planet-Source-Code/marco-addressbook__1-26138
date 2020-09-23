VERSION 5.00
Begin VB.Form frmaddressbook 
   Caption         =   "Adress Book"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton shownames 
      Caption         =   "Show names"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear all"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtemails 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtnames 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.ListBox lstemails 
      Height          =   2010
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.ListBox lstnames 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Emal addresses"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Names"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblemails 
      Caption         =   "Email address:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblnames 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   735
   End
End
Attribute VB_Name = "frmaddressbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mydb As Database    'Remember to add the DAO 3.5 object in the references
Dim mytable As Recordset
Dim FieldPosition As Integer


Public Sub LoadDB(ByVal DBName)
 Set mydb = Workspaces(0).OpenDatabase(App.Path & DBName)
 Set mytable = mydb.OpenRecordset("table1")
End Sub

Private Sub cmdclear_Click()
 Call ClearBoxes
End Sub

Private Sub cmddelete_Click()

 If txtnames.Text <> "" Then
  If txtemails.Text <> "" Then
    mytable.MoveFirst
    mytable.Move FieldPosition
    If MsgBox("Do you really want to erase the entry " & lstnames.Text & "?", vbYesNo, "Confirm operation!") = vbNo Then Exit Sub
    mytable.Delete
    mytable.MovePrevious
    Call ClearBoxes
  Else
    MsgBox "Select from the list boxes the item  you want to erase!", vbExclamation, "Alert messge"
    Exit Sub
   End If
 Else
   MsgBox "Select from the list boxes the item you want to erase!", vbExclamation, "Alert messge"
 End If
 
End Sub

Private Sub cmdedit_Click()
 
 If txtnames.Text = "" Then
  MsgBox "Enter the bookmark name!", vbExclamation, "Attention!"
  Exit Sub
 End If
 If txtemails.Text = "" Then
  MsgBox "Enter the email address!", vbExclamation, "Attention!"
  Exit Sub
 End If
 
 mytable.MoveFirst
 mytable.Move FieldPosition

 mytable.Edit
  mytable.Fields(0) = Trim(txtnames.Text)
  mytable.Fields(1) = Trim(txtemails.Text)
 mytable.Update
 Call ClearBoxes


End Sub

Private Sub cmdexit_Click()
 Call ClearBoxes
 mydb.Close
 End
End Sub

Public Sub ShowDB()
 Do Until mytable.EOF = True
  lstnames.AddItem mytable(0)
  lstemails.AddItem mytable(1)
  mytable.MoveNext
 Loop
End Sub

Public Sub ClearBoxes()
 Clipboard.Clear
 lstnames.Clear
 lstemails.Clear
 txtnames.Text = ""
 txtemails.Text = ""
End Sub

Private Sub cmdupdate_Click()
 If txtnames.Text = "" Then
  MsgBox "Enter the bookmark name!", vbExclamation, "Attention!"
  Exit Sub
 End If
 If txtemails.Text = "" Then
  MsgBox "Enter the email address!", vbExclamation, "Attention!"
  Exit Sub
 End If
 
 mytable.AddNew
  mytable.Fields(0) = Trim(txtnames.Text)
  mytable.Fields(1) = Trim(txtemails.Text)
  mytable.Update
 Call ClearBoxes
 
End Sub


Private Sub Form_Load()
 
' ADDRESS BOOK  -  Createdy by Marco (marcoe_wap@vizzavi.it) - Freeware version
'
' Personal email address book that opens directly the default email software
' copies and pastes the email address selected so as to be ready to
' write and send an email message.
' Useful when you use different PCs with different email software
'
' Commands Description:
' EMAIL ADDRESSES => double click on an email address to open the default email software and send a mail to the address selected
' SHOW NAMES => show all the data sorted by name
' CLEAR ALL => clear all the fields without editing the database
' UPDATE => create a new record with the data inserted in the text boxes
' DELETE => delete the record selected
' EDIT => edit the current record with what entered in the text boxes
'
' For further information: marcoe_wap@vizzavi.it
 
 
 Dim Pwd, PwdInsered As String, Counter As Byte
 Call ClearBoxes
 Call LoadDB("\addressbook.mdb")
 mytable.Index = "name"
 Call ShowDB
End Sub

Private Sub lstemails_Click()
 lstnames.ListIndex = lstemails.ListIndex
 txtnames.Text = lstnames.Text
 txtemails.Text = lstemails.Text
 FieldPosition = lstnames.ListIndex

End Sub

Private Sub lstemails_DblClick()

 Dim EmailAddress As String
 EmailAddress = lstemails.Text
 If EmailAddress = "none" Then
  MsgBox "The item selected is not an email address!", vbExclamation, "Sending email"
  Exit Sub
 End If
 Clipboard.Clear
 ShellExecute hwnd, "open", "mailto:" & EmailAddress, vbNullString, vbNullString, SW_SHOW

End Sub

Private Sub lstnames_Click()
 lstemails.ListIndex = lstnames.ListIndex
 txtnames.Text = lstnames.Text
 txtemails.Text = lstemails.Text
 FieldPosition = lstnames.ListIndex

End Sub

Private Sub shownames_Click()
 Call ClearBoxes
 Call LoadDB("\addressbook.mdb")
 mytable.Index = "name"
 Call ShowDB

End Sub
