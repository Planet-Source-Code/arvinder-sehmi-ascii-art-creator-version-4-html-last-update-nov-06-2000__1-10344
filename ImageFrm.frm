VERSION 5.00
Begin VB.Form ImageFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ImageFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4410
      Left            =   45
      Picture         =   "ImageFrm.frx":1472
      ScaleHeight     =   4410
      ScaleWidth      =   3690
      TabIndex        =   3
      Top             =   450
      Width           =   3690
   End
   Begin VB.PictureBox AppTitleBar 
      BackColor       =   &H80000002&
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   3690
      TabIndex        =   0
      Top             =   0
      Width           =   3750
      Begin VB.CommandButton AppEnd 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3375
         TabIndex        =   1
         Top             =   45
         Width           =   285
      End
      Begin VB.Label AppCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   240
         Left            =   405
         TabIndex        =   2
         Top             =   45
         Width           =   660
      End
      Begin VB.Image AppIcon 
         Height          =   240
         Left            =   45
         Stretch         =   -1  'True
         Top             =   45
         Width           =   240
      End
   End
End
Attribute VB_Name = "ImageFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API's (For TitleBar)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'Enable The Form to Be Moved
Public Sub MoveForm(Button As Integer)
 If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
 End If
End Sub
' Move The Form
Private Sub AppCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveForm Button
End Sub
' End Application
Private Sub AppEnd_Click()
 Unload Me
 End
End Sub
' Move The Form
Private Sub AppTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveForm Button
End Sub
Private Sub Form_Load()
 AppIcon.Picture = Me.Icon
End Sub
