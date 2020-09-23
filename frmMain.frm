VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dynamic Buttons"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   8280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create a Button"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Keep in mind that the coordinates are pixels"
      Height          =   195
      Left            =   4080
      TabIndex        =   5
      Top             =   8400
      Width           =   3075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "x"
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   8400
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "y"
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   8400
      Width           =   75
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To create a button dynamicly you must first have an array of buttons
'to make an object part of an array you have to assign it an index
'in the properties window, then you just follow the steps below

Dim i As Integer 'Buttons Index
Dim checkXnumber As Boolean 'Contains the value that checks if x is numeric
Dim chechYnumber As Boolean 'Contains the value that checks if y is numeric
Private Sub Command1_Click(Index As Integer)
checkXnumber = IsNumeric(Text1.Text) 'Checks if x coordinate is numeric
checkYnumber = IsNumeric(Text2.Text) 'Checks if y coordinate is numeric

If checkXnumber = False Or checkYnumber = False Then
    MsgBox "Please enter the proper coordinates", vbSystemModal 'Tells user that he entered the data incorrectly, if the data was entered correctly it skips this step
    Text1.SetFocus
    Exit Sub
End If

If Text1.Text = "" Or Text2.Text = "" Then 'Tells user that he entered the data incorrectly, if the data was entered correctly it skips this step
    MsgBox "Please enter the proper coordinates", vbSystemModal
    Text1.SetFocus
    Exit Sub
End If

Load Command1(i) 'Creates the button by using the command(0) prototype
Command1(i).Visible = True 'Makes the button visible, Every time you create a button the default visible value is false so we need to change it to true
Command1(i).Left = Text1.Text 'Assigns the x coordinate
Command1(i).Top = Text2.Text  'Assigns the y coordinate
Command1(i).Caption = "Index: " & i 'Gives the button a unique caption
i = i + 1 'Increases the index of the button

End Sub


Private Sub Form_Load()
i = 1 'Assigns value of 1 because 0 already exists
End Sub
