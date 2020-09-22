VERSION 5.00
Begin VB.Form OptionsFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4485
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Options.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Manual Control:"
      Height          =   4020
      Left            =   1620
      TabIndex        =   7
      Top             =   450
      Width           =   9105
      Begin VB.Frame Frame7 
         Caption         =   "[Html Colour ASCII]"
         Height          =   1050
         Left            =   180
         TabIndex        =   33
         Top             =   2925
         Width           =   8745
         Begin VB.CheckBox OpenHtml 
            Caption         =   "Open Html File After Creation?"
            Height          =   285
            Left            =   135
            TabIndex        =   37
            Top             =   630
            Value           =   1  'Checked
            Width           =   2445
         End
         Begin VB.TextBox HtmlTextPattern 
            Height          =   330
            Left            =   1530
            TabIndex        =   36
            Text            =   "Text-Pattern-For-Html-Ascii-Art-"
            Top             =   225
            Width           =   2760
         End
         Begin VB.CommandButton HtmlMake 
            Caption         =   "Create And Write Html Ascii To File"
            Height          =   330
            Left            =   4455
            TabIndex        =   34
            Top             =   630
            Width           =   4110
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "It Is Difficult to get Html to show all 255 ASCII Characters, So Other Text Is Needed To Fix This Problem. "
            ForeColor       =   &H80000010&
            Height          =   420
            Left            =   4500
            TabIndex        =   38
            Top             =   180
            Width           =   4155
         End
         Begin VB.Label Label5 
            Caption         =   "Html Pattern Text:"
            Height          =   285
            Left            =   135
            TabIndex        =   35
            Top             =   270
            Width           =   1500
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "[Line Art Options]"
         Height          =   555
         Left            =   180
         TabIndex        =   29
         Top             =   2295
         Width           =   8745
         Begin VB.CheckBox LineArt 
            Caption         =   "Enable Line Art"
            Height          =   285
            Left            =   225
            TabIndex        =   31
            Top             =   225
            Width           =   1500
         End
         Begin VB.HScrollBar Tolorance 
            Height          =   240
            LargeChange     =   5
            Left            =   2970
            Max             =   255
            TabIndex        =   30
            Top             =   225
            Value           =   100
            Width           =   5640
         End
         Begin VB.Label Label3 
            Caption         =   "Tolorance."
            Height          =   240
            Left            =   2115
            TabIndex        =   32
            Top             =   225
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "[1) Resize Image]"
         Height          =   2040
         Left            =   180
         TabIndex        =   15
         Top             =   225
         Width           =   2265
         Begin VB.CommandButton ResizeClear 
            Caption         =   "Clear Resized Image"
            Height          =   420
            Left            =   90
            TabIndex        =   25
            Top             =   1530
            Width           =   2085
         End
         Begin VB.CommandButton ResizeIt 
            Caption         =   "Resize Image"
            Height          =   420
            Left            =   90
            TabIndex        =   18
            Top             =   1035
            Width           =   2085
         End
         Begin VB.OptionButton Sample 
            Caption         =   "Sample (Faster)"
            Height          =   195
            Left            =   45
            TabIndex        =   17
            Top             =   675
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton Interplate 
            Caption         =   "Interploate (Better Quality)"
            Height          =   195
            Left            =   45
            TabIndex        =   16
            Top             =   315
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "[3) Draw Ascii Image]"
         Height          =   2040
         Left            =   5175
         TabIndex        =   13
         Top             =   225
         Width           =   3705
         Begin VB.CommandButton ResetDraw 
            Caption         =   "Clear Ascii Image"
            Height          =   465
            Left            =   1935
            TabIndex        =   27
            Top             =   1485
            Width           =   1680
         End
         Begin VB.OptionButton Only4Chars 
            Caption         =   "Only Use 4 ASCII Characters."
            Height          =   240
            Left            =   180
            TabIndex        =   24
            Top             =   1170
            Width           =   2445
         End
         Begin VB.OptionButton AllChars 
            Caption         =   "Use All 255 ASCII Characters."
            Height          =   285
            Left            =   180
            TabIndex        =   23
            Top             =   855
            Value           =   -1  'True
            Width           =   2445
         End
         Begin VB.HScrollBar DetailH 
            Height          =   195
            Left            =   1395
            Max             =   4
            Min             =   1
            TabIndex        =   20
            Top             =   270
            Value           =   1
            Width           =   2220
         End
         Begin VB.HScrollBar DetailV 
            Height          =   195
            Left            =   1395
            Max             =   4
            Min             =   1
            TabIndex        =   19
            Top             =   585
            Value           =   1
            Width           =   2220
         End
         Begin VB.CommandButton DrawIt 
            Caption         =   "Draw Ascii Image"
            Height          =   465
            Left            =   90
            TabIndex        =   14
            Top             =   1485
            Width           =   1770
         End
         Begin VB.Label Label2 
            Caption         =   "Vertical Detail:"
            Height          =   195
            Left            =   90
            TabIndex        =   22
            Top             =   585
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Horizontal Detail:"
            Height          =   195
            Left            =   90
            TabIndex        =   21
            Top             =   270
            Width           =   1635
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "[2) Grey Scale Image]"
         Height          =   2040
         Left            =   2745
         TabIndex        =   8
         Top             =   225
         Width           =   2220
         Begin VB.CommandButton GreyedClear 
            Caption         =   "Clear Grey Scaled Image"
            Height          =   420
            Left            =   90
            TabIndex        =   26
            Top             =   1530
            Width           =   2040
         End
         Begin VB.CommandButton GreyIt 
            Caption         =   "Grey Scale Image"
            Height          =   420
            Left            =   90
            TabIndex        =   12
            Top             =   1035
            Width           =   2025
         End
         Begin VB.CheckBox BlueValues 
            Caption         =   "Use Blue Values"
            Height          =   240
            Left            =   180
            TabIndex        =   11
            Top             =   720
            Width           =   1635
         End
         Begin VB.CheckBox GreenValues 
            Caption         =   "Use Green Values"
            Height          =   285
            Left            =   180
            TabIndex        =   10
            Top             =   450
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox RedValues 
            Caption         =   "Use Red Values"
            Height          =   240
            Left            =   180
            TabIndex        =   9
            Top             =   225
            Width           =   1635
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Automatic"
      Height          =   2265
      Left            =   45
      TabIndex        =   3
      Top             =   450
      Width           =   1545
      Begin VB.CommandButton Command1 
         Caption         =   "Start (Html)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         TabIndex        =   39
         Top             =   720
         Width           =   1365
      End
      Begin VB.CommandButton StartIt 
         Caption         =   "Start (Auto)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton LoadImage 
         Caption         =   "Load New Image"
         Height          =   420
         Left            =   90
         TabIndex        =   5
         Top             =   1215
         Width           =   1365
      End
      Begin VB.CommandButton Save 
         Caption         =   "Save Ascii Image"
         Height          =   420
         Left            =   90
         TabIndex        =   4
         Top             =   1710
         Width           =   1365
      End
   End
   Begin VB.PictureBox AppTitleBar 
      BackColor       =   &H80000002&
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   10665
      TabIndex        =   0
      Top             =   0
      Width           =   10725
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
         Left            =   10350
         TabIndex        =   1
         Top             =   45
         Width           =   285
      End
      Begin VB.Label AppCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
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
         Width           =   810
      End
      Begin VB.Image AppIcon 
         Height          =   240
         Left            =   45
         Stretch         =   -1  'True
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.Label HelpMe 
      Alignment       =   2  'Center
      ForeColor       =   &H80000010&
      Height          =   1230
      Left            =   0
      TabIndex        =   28
      Top             =   2970
      Width           =   1590
   End
End
Attribute VB_Name = "OptionsFrm"
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
Private Sub AppCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 MoveForm Button
End Sub
' End Application
Private Sub AppEnd_Click()
 Unload Me
 End
End Sub
' Move The Form
Private Sub AppTitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 MoveForm Button
End Sub

Private Sub Command1_Click()
 HtmlTextPattern.Text = "0"
 MainFrm.ResizeImage
 MainFrm.CreateHtmlART
 If OpenHtml.Value = 1 Then Shell "start " & HTMLFileName, vbHide
End Sub

Private Sub HtmlMake_Click()
 MainFrm.ResizeImage
 MainFrm.CreateHtmlART
 If OpenHtml.Value = 1 Then Shell "start " & HTMLFileName, vbHide
End Sub

Private Sub ResetDraw_Click()
 MainFrm.CancelDraw = True ' Cancel Draw
 MainFrm.ASCII_Image.Text = "" 'Reset Image
 MainFrm.ASCIIProg.Caption = "0%" ' Reset Percent
End Sub

Private Sub DrawIt_Click()
 ' Draw The Ascii Art
 MainFrm.CreateART
End Sub

Private Sub GreyedClear_Click()
 MainFrm.Greyed.Cls
 MainFrm.GreyProg.Caption = "0%"
End Sub

Private Sub LoadImage_Click()
 Const MaxHW = 1725 ' The Largest Width And Height Of The Resized Image
 'Load A New Image
 On Error Resume Next
 File = Open_File(Me.hWnd)
 If Trim(File) = "" Then Exit Sub
 With MainFrm               '\
  .Resized.Width = MaxHW    ' \
  .Resized.Height = MaxHW   '  >------ Reset The Size Of The Small Images
  .Greyed.Width = MaxHW     ' /
  .Greyed.Height = MaxHW    '/
  .Resized.Cls
  .Greyed.Cls    ' Clear Old Images
 End With
 With ImageFrm
  .Pic.Picture = LoadPicture(File) ' Load Picture
  If .Pic.Width > .Pic.Height Then
   MainFrm.Resized.Height = (.Pic.Height / (.Pic.Width / MaxHW)) '--Resize The height
   MainFrm.Greyed.Height = (.Pic.Height / (.Pic.Width / MaxHW))  '/
  Else
   MainFrm.Resized.Width = (.Pic.Width / (.Pic.Height / MaxHW)) '--Resize The Width
   MainFrm.Greyed.Width = (.Pic.Width / (.Pic.Height / MaxHW))  '/
  End If
 End With
End Sub

Private Sub ResizeClear_Click()
 MainFrm.Resized.Cls
 MainFrm.ResizeProg.Caption = "0%"
End Sub

Private Sub StartIt_Click()
 'Start the Auto Process
 MainFrm.Start
 GreyIt.Enabled = True
 DrawIt.Enabled = True
End Sub

Private Sub GreenValues_Click() ' Make Sure At Least One Colour Is Selected
 If BlueValues.Value = 0 _
 And RedValues.Value = 0 _
 And GreenValues.Value = 0 _
 Then GreenValues.Value = 1
End Sub
Private Sub RedValues_Click() ' Make Sure At Least One Colour Is Selected
 If BlueValues.Value = 0 _
 And RedValues.Value = 0 _
 And GreenValues.Value = 0 _
 Then RedValues.Value = 1
End Sub
Private Sub BlueValues_Click() ' Make Sure At Least One Colour Is Selected
 If BlueValues.Value = 0 _
 And RedValues.Value = 0 _
 And GreenValues.Value = 0 _
 Then BlueValues.Value = 1
End Sub

Private Sub GreyIt_Click()
 ' Greyscale resized image
 MainFrm.GreyScaleImage
End Sub

Private Sub Form_Load()
 AppIcon.Picture = Me.Icon
 HelpMe.Caption = "If You Need Help," & vbCr & "Just E-Mail Me:" & vbCr & "Arvi@Sehmi.org.uk" & vbCr & vbCr & "www.Arvinder.co.uk"
End Sub
Private Sub ResizeIt_Click()
 'Resize Loaded Image
 If Sample.Value = True Then MainFrm.ResizeImage Else MainFrm.InterpolateResizeImage
End Sub

Private Sub Save_Click()
 'Save ASCII Art
 Dim Filename As String
 On Error Resume Next
 InitDlgs ' Initialize Dialogs
 Filename = Save_File(MainFrm.hWnd) ' Show SaveFile Dlg
 Filename = Left(Filename, Len(Filename) - 1) ' Trim Of Last Char (It Is A Null Char, So Get Rid Of It)
 If Trim(Filename) = "" Then Exit Sub ' Check If Filename  Is Valid
 If LCase(Right(Filename, 4)) <> ".txt" Then Filename = Filename & ".txt" ' Check For Extension
 Open Filename For Output As #2 ' Open File
  Print #2, "Best Viewed In The Font: Terminal" '--Write output
  Print #2, MainFrm.ASCII_Image.Text            '/
 Close #2 ' close file
End Sub

