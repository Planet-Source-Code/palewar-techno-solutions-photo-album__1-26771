VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Photo Album"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   5730
   ControlBox      =   0   'False
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Insert Titles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   768
      Left            =   156
      TabIndex        =   11
      Top             =   1965
      Width           =   5490
      Begin VB.TextBox txtsubtitle 
         Height          =   288
         Left            =   852
         TabIndex        =   1
         Top             =   300
         Width           =   2028
      End
      Begin VB.CommandButton cmdtinsert 
         Caption         =   "&Insert"
         Height          =   312
         Left            =   4590
         TabIndex        =   3
         Top             =   300
         Width           =   744
      End
      Begin VB.ComboBox cmbtpgno 
         Height          =   315
         Left            =   3930
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   588
      End
      Begin VB.Label lbltitleno 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   300
         Width           =   105
      End
      Begin VB.Label Label7 
         Caption         =   "Title"
         Height          =   228
         Left            =   492
         TabIndex        =   13
         Top             =   300
         Width           =   396
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Picture No."
         Height          =   195
         Left            =   2955
         TabIndex        =   12
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   396
      Left            =   3081
      TabIndex        =   8
      Top             =   3060
      Width           =   972
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Create"
      Height          =   396
      Left            =   1677
      TabIndex        =   7
      Top             =   3075
      Width           =   972
   End
   Begin VB.TextBox txtpage 
      Height          =   264
      Left            =   2424
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1470
      Width           =   540
   End
   Begin VB.TextBox txttitle 
      Height          =   264
      Left            =   2424
      TabIndex        =   0
      Top             =   930
      Width           =   2484
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CreateAlbum"
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
      Left            =   1464
      TabIndex        =   10
      Top             =   72
      Width           =   2364
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "(Step 3 of 3)"
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
      Left            =   1956
      TabIndex        =   9
      Top             =   372
      Width           =   1404
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Pictures"
      Height          =   195
      Left            =   945
      TabIndex        =   5
      Top             =   1485
      Width           =   975
   End
   Begin VB.Label lbltitle 
      Caption         =   "Album Title "
      Height          =   210
      Left            =   975
      TabIndex        =   4
      Top             =   990
      Width           =   975
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
'Developed by Sachin Palewar, Nagpur (India)
'email:- palewar@hotmail.com
'web :- http://compuwhizkid.tripod.com
'************************************************
Dim titles As String
Dim titlepgno As String
Private Sub Cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
If Trim(txttitle) = "" Then ' check if album title has been entered
MsgBox "Please Enter a Title.", vbCritical, "Title Required"
txttitle.SetFocus
Exit Sub
End If

If Trim(titles) = "" Then
MsgBox "You have not specified any Titles.", vbInformation, "Titles will not be included"
creatalbum
Exit Sub
End If
creatalbum
End Sub

Sub creatalbum()
Dim n As Integer
Dim albumno As Integer

If titles <> "" Then
ReDim titlearr(UBound(Split(titles, " %% ")))
ReDim titlepgnoarr(UBound(Split(titlepgno, " @@ ")))

For i = 1 To UBound(titlearr)
titlearr(i) = Split(titles, " %% ")(i)
titlepgnoarr(i) = Split(titlepgno, " @@ ")(i)
Next
End If

FileCopy App.Path + "\home.gif", ReturnValue + "\home.gif"
FileCopy App.Path + "\prev.gif", ReturnValue + "\prev.gif"
FileCopy App.Path + "\next.gif", ReturnValue + "\next.gif"

'Generating left navigation bar HTML
Open ReturnValue & "\left.htm" For Output As #1
Print #1, "<html><head><base target=_top></head>"
Print #1, "<body bgcolor=#ECECEC link=#000088><center><strong>"
Print #1, "<a href=index1.htm><img src=home.gif border=0 alt=Home></a></p><p>"

'include title hyperlinks
If Trim(titles) <> "" Then
For i = 1 To UBound(titlearr)
Print #1, "<font face=Arial size=-2><a href=index" & titlepgnoarr(i) & ".htm style=""text-decoration: none"">" & i & ". " & titlearr(i) & "</a><br><br>"
Next
End If

Print #1, "</body></html>"
Close #1

'create leftbottom navigation box
For j = 1 To txtpage
Open ReturnValue & "\prevnext" & j & ".htm" For Output As #2
Print #2, "<html><head></head><body bgcolor=#ECECEC><center>"
If j <> 1 Then
Print #2, "<a href=index" & j - 1; ".htm target=_top><img src=prev.gif border=0 alt=Previous></a>" 'if picture is not the first picture
End If
If j <> txtpage Then
Print #2, "<a href=index" & j + 1; ".htm target=_top><img src=next.gif border=0 alt=Next></a>" ' if picture is not the last picture
End If
Print #2, "</body></html>"
Close #2
Next

'create main index pages
For k = 1 To txtpage
Open ReturnValue & "\index" & k & ".htm" For Output As #3
Print #3, "<html><head><title>" & txttitle & " (Photo " & k & " of " & txtpage & ") </title>"
Print #3, "</head><frameset framespacing=0 border=false cols=148,*><frameset rows=*,21%>"
Print #3, "<frame name=left target=main src=left.htm scrolling=auto marginwidth=6 marginheight=10>"
Print #3, "<frame name=leftbottom src=prevnext" & k & ".htm scrolling=auto marginwidth=6 marginheight=10></frameset>"
Print #3, "<frame name=main src=" & k & "." & extn & " >< noframes >< body > "
Print #3, "<p>This page uses frames, but your browser doesn't support them.</p>"
Print #3, "</body></noframes></frameset></html>"
Close #3
Next

MsgBox "Photo Album Successfully Created.", vbInformation, "Success!!"
ShellExecute hwnd, "open", ReturnValue + "\index1.htm", vbNullString, vbNullString, conSwNormal
Unload Me
End Sub

Private Sub cmdtinsert_Click()
If Trim(txtsubtitle) = "" Then
MsgBox "Please enter title.", vbCritical, "Title Blank"
txtsubtitle.SetFocus
Exit Sub
End If

If cmbtpgno = "" Then
MsgBox "Please select a Picture no.", vbCritical, "Picture No. Not Selected"
cmbtpgno.SetFocus
Exit Sub
End If

titles = titles + " %% " + Trim(txtsubtitle)
titlepgno = titlepgno + " @@ " + Trim(cmbtpgno)

MsgBox "Title added.", vbInformation, "Done!!"

lbltitleno = Val(lbltitleno) + 1
txtsubtitle = ""
cmbtpgno.ListIndex = -1
txtsubtitle.SetFocus

End Sub

Private Sub Form_Load()
titlenosknown = False
End Sub

