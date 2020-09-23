VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3105
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   2775
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   4455
      Begin VB.Timer Timer1 
         Interval        =   4000
         Left            =   360
         Top             =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "http://compuwhizkid.tripod.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2175
         MouseIcon       =   "frmSplash.frx":000C
         TabIndex        =   5
         Top             =   2415
         Width           =   2190
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Sachin Palewar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2175
         TabIndex        =   4
         Top             =   2220
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Developed By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2175
         TabIndex        =   3
         Top             =   1920
         Width           =   1320
      End
      Begin VB.Label lblproductname 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Make Your Own Album In 3 Steps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   210
         Width           =   3900
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   2265
         TabIndex        =   1
         Top             =   690
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
'Developed by Sachin Palewar, Nagpur (India)
'email:- palewar@hotmail.com
'web :- http://compuwhizkid.tripod.com
'************************************************
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 End Sub

Private Sub Form_Unload(Cancel As Integer)
frmextension.Show
End Sub

Private Sub Frame1_Click()
Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub imgLogo_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
ShellExecute hwnd, "open", "http://compuwhizkid.tripod.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label3_Click()
ShellExecute hwnd, "open", "http://compuwhizkid.tripod.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbRed
End Sub

Private Sub lblProductName_Click()
Unload Me
End Sub

Private Sub lblVersion_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
