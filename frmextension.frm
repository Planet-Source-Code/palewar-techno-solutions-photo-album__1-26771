VERSION 5.00
Begin VB.Form frmextension 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Extension"
   ClientHeight    =   2745
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhowto 
      Caption         =   "HELP"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   3975
   End
   Begin VB.CommandButton Cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2290
      TabIndex        =   3
      Top             =   1740
      Width           =   816
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1138
      TabIndex        =   2
      Top             =   1755
      Width           =   816
   End
   Begin VB.TextBox txtextension 
      Height          =   288
      Left            =   2145
      TabIndex        =   1
      Top             =   972
      Width           =   1584
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(Step 1 of 3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   252
      Left            =   1542
      TabIndex        =   5
      Top             =   324
      Width           =   1404
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      Caption         =   "Create Album"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   312
      Left            =   1050
      TabIndex        =   4
      Top             =   24
      Width           =   2364
   End
   Begin VB.Label lblextension 
      Caption         =   "Enter Extension of Picture Files (gif, jpg etc.)"
      Height          =   705
      Left            =   270
      TabIndex        =   0
      Top             =   870
      Width           =   1665
   End
End
Attribute VB_Name = "frmextension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
'Developed by Sachin Palewar, Nagpur (India)
'email:- palewar@hotmail.com
'web :- http://compuwhizkid.tripod.com
'************************************************
Private Sub Cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdhowto_Click()
ShellExecute hwnd, "open", App.Path + "\readme.txt", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub cmdok_Click()
Dim WithFiles As Long
If Trim(txtextension.Text) = "" Then
MsgBox "Please enter an extension for Picture Files.", vbCritical, "Sachin's Photo Album"
txtextension.SetFocus
Exit Sub
End If
extn = Trim(txtextension)
Me.Hide
'bring up the Browse For Folder dialog box
ReturnValue = BrowseForFolder(Me.hwnd, "Create Album:" + Chr(13) + "(Step 2 of 3)", WithFiles, RecycleBin)
Checkreturn
End Sub

Private Sub Form_Load()
Me.Top = 2200
Me.Left = 2700
End Sub
Sub Checkreturn()
Dim msg As Integer, msg1 As Integer
Dim i As Integer
If ReturnValue = "" Then
msg = MsgBox("Select Yes to browse for Folder" + Chr(13) + "Select No to Exit Software.", vbYesNo, "No Folder Selected")
If msg = vbYes Then
cmdok.Value = True
Exit Sub
Else
Unload Me
Exit Sub
End If
End If
i = 1
If Dir(ReturnValue & "\1." + extn) = "" Then
msg1 = MsgBox("Software can not find 1." + extn + " file in the Folder." + Chr(13) + "Select another Folder?", vbYesNo, "File Not Found")
If msg1 = vbYes Then
cmdok.Value = True
Exit Sub
Else
Unload Me
Exit Sub
End If
End If
While Dir(ReturnValue & "\" & i & "." & extn) <> ""
i = i + 1
Wend
frmmain.Show
frmmain.txtpage = i - 1 ' No. of  files in the folder
For j = 1 To i - 1
frmmain.cmbtpgno.AddItem j
Next
frmmain.cmbtpgno.ListIndex = -1
ReDim dwgarr(i - 1)
Unload Me
End Sub
