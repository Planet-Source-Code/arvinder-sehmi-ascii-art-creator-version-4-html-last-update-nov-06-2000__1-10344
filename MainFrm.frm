VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "MainFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleWidth      =   2670
      TabIndex        =   10
      Top             =   4590
      Width           =   2670
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3: Create The ASCII Art."
         Height          =   195
         Left            =   45
         TabIndex        =   12
         Top             =   0
         Width           =   2130
      End
      Begin VB.Label ASCIIProg 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2160
         TabIndex        =   11
         Top             =   0
         Width           =   510
      End
   End
   Begin VB.PictureBox Greyed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   5085
      ScaleHeight     =   1725
      ScaleWidth      =   1725
      TabIndex        =   4
      Top             =   3285
      Width           =   1725
   End
   Begin VB.PictureBox Resized 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   5085
      ScaleHeight     =   1725
      ScaleWidth      =   1725
      TabIndex        =   3
      Top             =   990
      Width           =   1725
   End
   Begin VB.PictureBox AppTitleBar 
      BackColor       =   &H80000002&
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6855
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
         Left            =   6480
         TabIndex        =   1
         Top             =   45
         Width           =   285
      End
      Begin VB.Label AppCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image To ASCII Art, By Arvinder Sehmi"
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
         Width           =   4005
      End
      Begin VB.Image AppIcon 
         Height          =   240
         Left            =   45
         Stretch         =   -1  'True
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.TextBox ASCII_Image 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   450
      Width           =   4965
   End
   Begin VB.Label GreyProg 
      Caption         =   "0%"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6390
      TabIndex        =   9
      Top             =   2970
      Width           =   420
   End
   Begin VB.Label ResizeProg 
      Caption         =   "0%"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6390
      TabIndex        =   8
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label2 
      Caption         =   "Step 2: Grey Scaling The Image..."
      Height          =   510
      Left            =   5085
      TabIndex        =   6
      Top             =   2745
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Step 1: Resizing The Loaded Image...."
      Height          =   465
      Left            =   5085
      TabIndex        =   5
      Top             =   495
      Width           =   1770
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For Fake Title Bar
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
' Api's For Needed To Resize And Grey Scale The Image
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Hold Colour Values (0->255) For A Certain Point
Dim rRed As Integer, rGreen As Integer, rBlue As Integer
Public CancelDraw As Boolean ' If The Draw Process Needs To Be Canceled Halfway Through

' Move The Form On Titlebar Drag
Public Sub MoveForm(Button As Integer)
 If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    AlignForms ' Position Forms
 End If
End Sub
' Move The Form On Titlebar Drag
Private Sub AppCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 MoveForm Button
End Sub
'Unload The Application
Private Sub AppEnd_Click()
 Unload Me
 End
End Sub
Private Sub AppTitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 MoveForm Button ' Move The Form On Titlebar Drag
End Sub
Sub AlignForms() ' Position Forms
 ImageFrm.Left = Me.Left + Me.Width '----Set Form To The Right
 ImageFrm.Top = Me.Top              '/
 OptionsFrm.Top = Me.Top + Me.Height '------Set Form To Bottom
 OptionsFrm.Left = Me.Left           '/
End Sub

' Get Information From The Greyscale Image
' And Turn It Into ASCII Art
Public Sub CreateART()
 Const c0 = " ": Const c1 = "°": Const c2 = "±": Const c3 = "²": Const c4 = "Û" ' Set Values For 4 Shade Ascii Art
 CancelDraw = False
 Dim Col As Integer, Txt As String, Z As Double, Total As Long  'Declare Variables
 Dim AsciiPic As String
 ASCII_Image.Text = "" ' Reset Image
 Total = (Resized.Height * Resized.Width)
 For y = 0 To Resized.Height Step 15 * (3 - (OptionsFrm.DetailV.Value / 2)) 'Go Through Each Pixel
  For x = 0 To Resized.Width Step (15 * (3 - (OptionsFrm.DetailH.Value / 2))) / (8 / 6) 'In The Greyscale Image
    RGBfromLONG (GetPixel(Greyed.hdc, x / 15, y / 15)) ' Get The Red, Green, And Blue Values Of That Pixel
    Col = rRed ' We Only need Red, As In A Greyscale Image: Red = Blue = Green
    If OptionsFrm.Only4Chars.Value = True Then ' If Only 4 Char Mode Is Selected
     If Col >= 0 And Col <= 51 Then Txt = c4   '\
     If Col > 51 And Col <= 102 Then Txt = c3  ' \
     If Col > 102 And Col <= 153 Then Txt = c2 '  > Set Character Depending On Darkness
     If Col > 153 And Col <= 204 Then Txt = c1 ' /
     If Col > 204 And Col <= 255 Then Txt = c0 '/
    Else ' All Char Mode Is Selected
     For Z = 0 To 255 Step 5.3125 ' There Are 48 Different Char Colours ( 255/5.3212 = 48)
      If Col >= Z And Col <= Z + 5.3125 Then Txt = Chr(Ascii(Z / 5.3125)) ' If Colour Matches Then Get The Correct ASCII
     Next Z
    End If
   ''ASCII_Image.Text = ASCII_Image.Text & Txt ' Add The Ascii To The Image
   AsciiPic = AsciiPic & Txt
  Next x
  ''ASCII_Image.Text = ASCII_Image.Text & CStr(Chr(13) + Chr(10)) ' End Of Line, Place A Line Feed
  AsciiPic = AsciiPic & CStr(Chr(13) + Chr(10))
  ASCIIProg.Caption = Int(((x * y) / Total) * 100) & "%"
  ASCIIProg.Refresh
  ' Cancel Draw If Asked
  If CancelDraw = True Then ASCII_Image.Text = "": CancelDraw = False: ASCIIProg.Caption = "0%": Exit Sub
  DoEvents 'Refresh Stuff
 Next y
 ASCIIProg.Caption = "100%"
 ASCII_Image.Text = AsciiPic
End Sub
' Get Information From The Colour Image
' And Turn It Into HTML ASCII Art
Public Sub CreateHtmlART()
 
 Dim Filename As String, Ext As String
 On Error Resume Next
 InitDlgs ' Initialize Dialogs
 Filename = Save_File(MainFrm.hWnd) ' Show SaveFile Dlg
 Filename = Left(Filename, Len(Filename) - 1) ' Trim Of Last Char (It Is A Null Char, So Get Rid Of It)
 If Trim(Filename) = "" Then Exit Sub ' Check If Filename  Is Valid
 
 Ext = LCase(Right(Filename, 4))
 If (Ext = ".htm") Or (Ext = "html") Then Filename = LCase(Filename) Else Filename = Filename & ".htm"
 
 HTMLFileName = Filename
 Open Filename For Output As #3 ' Open File
 CancelDraw = False
 Dim Col As Integer, Txt As String, Z As Double, Total As Long  'Declare Variables
 Dim AsciiPic As String
 Dim Letter As Integer
 ASCII_Image.Text = "" ' Reset Image
 AsciiPic = MakeHtmlHead(OptionsFrm.HtmlTextPattern)
 Total = (Resized.Height * Resized.Width)
 For y = 0 To Resized.Height Step 15 * (3 - (OptionsFrm.DetailV.Value / 2)) 'Go Through Each Pixel
  For x = 0 To Resized.Width Step (15 * (3 - (OptionsFrm.DetailH.Value / 2))) / (8 / 6) 'In The Greyscale Image
    RGBfromLONG (GetPixel(Greyed.hdc, x / 15, y / 15)) ' Get The Red, Green, And Blue Values Of That Pixel
    Col = rRed ' We Only need Red, As In A Greyscale Image: Red = Blue = Green
    If OptionsFrm.Only4Chars.Value = True Then ' If Only 4 Char Mode Is Selected
     If Col >= 0 And Col <= 51 Then Txt = c4   '\
     If Col > 51 And Col <= 102 Then Txt = c3  ' \
     If Col > 102 And Col <= 153 Then Txt = c2 '  > Set Character Depending On Darkness
     If Col > 153 And Col <= 204 Then Txt = c1 ' /
     If Col > 204 And Col <= 255 Then Txt = c0 '/
    Else ' All Char Mode Is Selected
     'For Z = 0 To 255 Step 5.3125 ' There Are 48 Different Char Colours ( 255/5.3212 = 48)
      'If Col >= Z And Col <= Z + 5.3125 Then
        'Txt = Chr(Ascii(Z / 5.3125)) ' If Colour Matches Then Get The Correct ASCII
        Letter = Letter + 1
        Txt = Mid(OptionsFrm.HtmlTextPattern, Letter, 1)
        If Letter = Len(OptionsFrm.HtmlTextPattern) Then Letter = 0
        If GetPixel(Resized.hdc, x / 15, y / 15) = GetPixel(Resized.hdc, (x / 15) - 1, y / 15) Then
         If GetPixel(Resized.hdc, x / 15, y / 15) = GetPixel(Resized.hdc, (x / 15) + 1, y / 15) Then
          Txt = Txt
         Else
          Txt = Txt & "</font>"
         End If
        Else
         If GetPixel(Resized.hdc, x / 15, y / 15) = GetPixel(Resized.hdc, (x / 15) + 1, y / 15) Then
          Txt = "<font color=" & Sp & LongToHex(GetPixel(Resized.hdc, x / 15, y / 15)) & Sp & ">" & Txt
         Else
          Txt = "<font color=" & Sp & LongToHex(GetPixel(Resized.hdc, x / 15, y / 15)) & Sp & ">" & Txt & "</font>"
         End If
        End If
      'End If
     'Next Z
    End If
   ''ASCII_Image.Text = ASCII_Image.Text & Txt ' Add The Ascii To The Image
   AsciiPic = AsciiPic & Txt
  Next x
  ''ASCII_Image.Text = ASCII_Image.Text & CStr(Chr(13) + Chr(10)) ' End Of Line, Place A Line Feed
  AsciiPic = AsciiPic & "<br>" & CStr(Chr(13) + Chr(10))
  Print #3, AsciiPic
  AsciiPic = ""
  ASCIIProg.Caption = Int(((x * y) / Total) * 100) & "%"
  ASCIIProg.Refresh
  ' Cancel Draw If Asked
  If CancelDraw = True Then ASCII_Image.Text = "": CancelDraw = False: Close #3: ASCIIProg.Caption = "0%": Exit Sub
  DoEvents 'Refresh Stuff
 Next y
 ASCIIProg.Caption = "100%"
 Print #3, MakeHtmlFoot
 Close #3 ' close file
 ASCII_Image.Text = AsciiPic
End Sub
Public Function LongToHex(Colour As Long) As String
    
    
    Dim rHex As String, gHex As String, bHex As String
    Dim rCol As Integer, gCol As Integer, bCol As Integer
    
    rCol = Colour Mod &H100:    Colour = Colour \ &H100
    gCol = Colour Mod &H100:    Colour = Colour \ &H100
    bCol = Colour Mod &H100
        
    rHex = Hex(rCol):    If Len(rHex) < 2 Then rHex = "0" & rHex
    gHex = Hex(gCol):    If Len(gHex) < 2 Then gHex = "0" & gCol
    bHex = Hex(bCol):    If Len(bHex) < 2 Then bHex = "0" & bHex
    
    LongToHex = "#" & rHex & gHex & bHex
End Function

'Make The Header For The Html File
Public Function MakeHtmlHead(Title As String) As String
 Dim Ln As String ' Holds Line Chars
 Dim Sp As String ' Holds Speach Mark Chars
 Ln = CStr(Chr(13) + Chr(10)) 'LineFeed
 Sp = Chr(34) ' Speach Mark
 MakeHtmlHead = _
 "<html>" & Ln & _
 "<head>" & Ln & _
 "<meta http-equiv=" & Sp & "Content-Type" & Sp & " content=" & Sp & "text/html; charset=iso-8859-1" & Sp & ">" & Ln & _
 "<title>" & Title & "</title>" & Ln & _
 "<meta name=" & Sp & "Microsoft Border" & Sp & " content=" & Sp & "none" & Sp & ">" & Ln & _
 "</head>" & Ln & _
 "<body>" & Ln & _
 "<font face=" & Sp & "Terminal" & Sp & "><p><small><small><br>" & Ln
End Function
'make The Footer For The Html File
Public Function MakeHtmlFoot() As String
 Dim Ln As String ' Holds Line Chars
 Ln = CStr(Chr(13) + Chr(10)) 'LineFeed
 MakeHtmlFoot = Ln & _
 "</font></p>" & Ln & _
 "</body>" & Ln & _
 "</html>"
End Function
Private Sub Form_Activate()
 AlignForms ' Align Forms
End Sub
Private Sub Form_Load()
 AppIcon.Picture = Me.Icon ' Set Icon
 InitDlgs ' Initailize Open & Save Dialogues
 OptionsFrm.DetailH.Value = 4 '-Set Defaults
 OptionsFrm.DetailV.Value = 4 '/
 With ImageFrm
  If .Pic.Width > .Pic.Height Then
   MainFrm.Resized.Height = (.Pic.Height / (.Pic.Width / 1725)) '--Resize The height
   MainFrm.Greyed.Height = (.Pic.Height / (.Pic.Width / 1725))  '/
  Else
   MainFrm.Resized.Width = (.Pic.Width / (.Pic.Height / 1725)) '--Resize The Width
   MainFrm.Greyed.Width = (.Pic.Width / (.Pic.Height / 1725))  '/
  End If
  .Show   '-Show other Forms
  OptionsFrm.Show
 End With
End Sub
Public Sub Start()
 If OptionsFrm.Sample.Value = True Then ResizeImage 'Resize the image
 If OptionsFrm.Interplate.Value = True Then InterpolateResizeImage 'Accuratly Resize The Image
 GreyScaleImage ' Turn It Into Greyscale
 CreateART 'make ASCII Image
End Sub
Public Sub ResizeImage()
 ' OLD RESIZE SUB.
 'Dim rX, rY, Total As Long
 'Total = (ImageFrm.Pic.Height * ImageFrm.Pic.Width)
 'For y = 0 To ImageFrm.Pic.Height Step ((ImageFrm.Pic.Height / Resized.Height) * 15) '-Cycle ThroughLarge Image
 '   For x = 0 To ImageFrm.Pic.Width Step ((ImageFrm.Pic.Width / Resized.Width) * 15) '/
 '       SetPixel Resized.hdc, rX / 15, rY / 15, GetPixel(ImageFrm.Pic.hdc, x / 15, y / 15) ' Get Every Other Pixel, and Place It Down In Another Picturebox
 '       rX = rX + 15 ' Set Pixel Position On Small Image
 '   Next x
 '   Resized.Refresh
 '   rY = rY + 15 ' Goto Next Line
 '   rX = 0 ' Start On Left Side
 '   ResizeProg.Caption = CInt(((x * y) / Total) * 100) & "%"
 '   ResizeProg.Refresh
 'Next y
 Call StretchBlt(Resized.hdc, 0, 0, Resized.Width / Screen.TwipsPerPixelX, Resized.Height / Screen.TwipsPerPixelY, ImageFrm.Pic.hdc, 0, 0, ImageFrm.Pic.Width / Screen.TwipsPerPixelX, ImageFrm.Pic.Height / Screen.TwipsPerPixelY, vbSrcCopy)
 Resized.Refresh
End Sub
Public Sub InterpolateResizeImage()
 Dim rX, rY, Total As Long
 On Error Resume Next
 Total = (ImageFrm.Pic.Height * ImageFrm.Pic.Width)
 For y = 0 To ImageFrm.Pic.Height Step ((ImageFrm.Pic.Height / Resized.Height) * 15) - 1 '-Cycle Through large image
    For x = 0 To ImageFrm.Pic.Width Step ((ImageFrm.Pic.Width / Resized.Width) * 15) - 1 '/
        Dim cRed, cBlue, cGreen ' Reset Variables
        For sX = x To (x + ((ImageFrm.Pic.Width / Resized.Width) * 15)) Step 15       '\
            For sY = y To (y + ((ImageFrm.Pic.Height / Resized.Height) * 15)) Step 15 ' \
               RGBfromLONG (GetPixel(ImageFrm.Pic.hdc, sX / 15, sY / 15))             '  \
               cRed = (cRed + rRed) / 2                                               '   \-- Get An Average Colour For
               cBlue = (cBlue + rBlue) / 2                                            '   /-- Surronding Pixels
               cGreen = (cGreen + rGreen) / 2                                         '  /
            Next sY                                                                   ' /
        Next sX                                                                       '/
        SetPixel Resized.hdc, rX / 15, rY / 15, RGB(cRed, cGreen, cBlue) ' place Down The Pixel
        rX = rX + 15 ' Set Pixel Position On Small Image
    Next x
    Resized.Refresh
    rY = rY + 15 ' Goto Next Line
    rX = 0 ' Start On Left Side
    ResizeProg.Caption = CInt(((x * y) / Total) * 100) & "%"
    ResizeProg.Refresh
 Next y
End Sub
Public Sub GreyScaleImage()
 Dim AveCol As Integer, a As Integer, Total As Long
 Total = (Resized.Height * Resized.Width)
 For y = 0 To Resized.Height Step 15 ' cycle throght Image
  For x = 0 To Resized.Width Step 15
    AveCol = 0 'reset average colour
    a = 0 ' reset number of values added
    RGBfromLONG (GetPixel(Resized.hdc, x / 15, y / 15))
    If OptionsFrm.RedValues.Value = 1 Then AveCol = AveCol + rRed: a = a + 1 ' get the red value
    If OptionsFrm.BlueValues.Value = 1 Then AveCol = AveCol + rBlue: a = a + 1 ' get the blue value and add it to the red
    If OptionsFrm.GreenValues.Value = 1 Then AveCol = AveCol + rGreen: a = a + 1 ' get the green value and add it to the green
    If AveCol <= 0 Then AveCol = 0 ' error corrector
    AveCol = (AveCol / a) ' divide total by number of additions
    If OptionsFrm.LineArt.Value = 1 Then If AveCol > Val(OptionsFrm.Tolorance.Value) Then AveCol = 255 Else AveCol = 0
    SetPixel Greyed.hdc, x / 15, y / 15, RGB(AveCol, AveCol, AveCol) ' set pixel
  Next x
  Greyed.Refresh
  GreyProg.Caption = Int(((x * y) / Total) * 100) & "%"
  GreyProg.Refresh
 Next y
End Sub
Private Function RGBfromLONG(LongCol As Long)
 ' Get The Red, Blue And Green Values Of A Colour From The Long Value
 Dim Blue As Double, Green As Double, Red As Double, GreenS As Double, BlueS As Double
 Blue = Fix((LongCol / 256) / 256)
 Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
 Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
 rRed = Red: rBlue = Blue: rGreen = Green
End Function

'------------------------------------------------------------------------------------
' The Next Sub Contains a Database Of Character's Darkness
' For Example, a White Colour (Value Of 0) would Have No Black, Ie A Space " "
' A Darker Colour (Value Of Say 20) would Have Alot Of Black, Ie A An Eight "8"
'
' I Used Another Program That I Had Made To Create This Sub. The Basics Of It
' Is To Draw Each Character, and see How Many Black Pixels It Contains, It Then
' Creates The Sub. For Example If Col=0 Then There Are No Black Pixles, If Col =48
' All Are Black Pixles. Using This A Greyscale Colour Chart Can Be Made Using Letters.
'--------------------------------------------------------------------------------------
Public Function Ascii(Col As Integer) As String
Dim Rand As Integer
Col = 48 - Col
''''''''''''''''''''''''''''''''''''''
If Col = 0 Then Ascii = 32: Exit Function
''''''''''''''''''''''''''''''''''''''
If Col = 1 Or Col = 2 Then Ascii = 250: Exit Function
''''''''''''''''''''''''''''''''''''''
If Col = 4 Or Col = 3 Then
 Randomize: Rand = CInt(Rnd * 3)
 If Rand = 0 Then Ascii = 7: Exit Function
 If Rand = 1 Then Ascii = 46: Exit Function
 If Rand = 2 Then Ascii = 126: Exit Function
 If Rand = 3 Then Ascii = 249: Exit Function
 Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 5 Then
 Randomize: Rand = CInt(Rnd * 6)
 If Rand = 0 Then Ascii = 39: Exit Function
 If Rand = 1 Then Ascii = 44: Exit Function
 If Rand = 2 Then Ascii = 45: Exit Function
 If Rand = 3 Then Ascii = 47: Exit Function
 If Rand = 4 Then Ascii = 92: Exit Function
 If Rand = 5 Then Ascii = 94: Exit Function
 If Rand = 6 Then Ascii = 96: Exit Function
Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 6 Then
 Randomize: Rand = CInt(Rnd * 5)
 If Rand = 0 Then Ascii = 95: Exit Function
 If Rand = 1 Then Ascii = 124: Exit Function
 If Rand = 2 Then Ascii = 174: Exit Function
 If Rand = 3 Then Ascii = 175: Exit Function
 If Rand = 4 Then Ascii = 192: Exit Function
 If Rand = 5 Then Ascii = 196: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 7 Then
 Randomize: Rand = CInt(Rnd * 10)
 If Rand = 0 Then Ascii = 40: Exit Function
 If Rand = 1 Then Ascii = 41: Exit Function
 If Rand = 2 Then Ascii = 60: Exit Function
 If Rand = 3 Then Ascii = 62: Exit Function
 If Rand = 4 Then Ascii = 105: Exit Function
 If Rand = 5 Then Ascii = 141: Exit Function
 If Rand = 6 Then Ascii = 217: Exit Function
 If Rand = 7 Then Ascii = 218: Exit Function
 If Rand = 8 Then Ascii = 231: Exit Function
 If Rand = 9 Then Ascii = 246: Exit Function
 If Rand = 10 Then Ascii = 253: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 8 Then
 Randomize: Rand = CInt(Rnd * 17)
 If Rand = 0 Then Ascii = 22: Exit Function
 If Rand = 1 Then Ascii = 28: Exit Function
 If Rand = 2 Then Ascii = 58: Exit Function
 If Rand = 3 Then Ascii = 108: Exit Function
 If Rand = 4 Then Ascii = 139: Exit Function
 If Rand = 5 Then Ascii = 140: Exit Function
 If Rand = 6 Then Ascii = 161: Exit Function
 If Rand = 7 Then Ascii = 169: Exit Function
 If Rand = 8 Then Ascii = 170: Exit Function
 If Rand = 9 Then Ascii = 179: Exit Function
 If Rand = 10 Then Ascii = 191: Exit Function
 If Rand = 11 Then Ascii = 212: Exit Function
 If Rand = 12 Then Ascii = 241: Exit Function
 If Rand = 13 Then Ascii = 244: Exit Function
 If Rand = 14 Then Ascii = 245: Exit Function
 If Rand = 15 Then Ascii = 247: Exit Function
 If Rand = 16 Then Ascii = 248: Exit Function
 If Rand = 17 Then Ascii = 252: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 9 Then
 Randomize: Rand = CInt(Rnd * 3)
 If Rand = 0 Then Ascii = 43: Exit Function
 If Rand = 1 Then Ascii = 59: Exit Function
 If Rand = 2 Then Ascii = 118: Exit Function
 If Rand = 3 Then Ascii = 193: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 10 Then
 Randomize: Rand = CInt(Rnd * 23)
 If Rand = 0 Then Ascii = 33: Exit Function
 If Rand = 1 Then Ascii = 34: Exit Function
 If Rand = 2 Then Ascii = 49: Exit Function
 If Rand = 3 Then Ascii = 61: Exit Function
 If Rand = 4 Then Ascii = 63: Exit Function
 If Rand = 5 Then Ascii = 106: Exit Function
 If Rand = 6 Then Ascii = 114: Exit Function
 If Rand = 7 Then Ascii = 116: Exit Function
 If Rand = 8 Then Ascii = 120: Exit Function
 If Rand = 9 Then Ascii = 123: Exit Function
 If Rand = 10 Then Ascii = 125: Exit Function
 If Rand = 11 Then Ascii = 155: Exit Function
 If Rand = 12 Then Ascii = 168: Exit Function
 If Rand = 13 Then Ascii = 173: Exit Function
 If Rand = 14 Then Ascii = 188: Exit Function
 If Rand = 15 Then Ascii = 189: Exit Function
 If Rand = 16 Then Ascii = 190: Exit Function
 If Rand = 17 Then Ascii = 194: Exit Function
 If Rand = 18 Then Ascii = 195: Exit Function
 If Rand = 19 Then Ascii = 224: Exit Function
 If Rand = 20 Then Ascii = 229: Exit Function
 If Rand = 21 Then Ascii = 236: Exit Function
 If Rand = 22 Then Ascii = 239: Exit Function
 If Rand = 23 Then Ascii = 251: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 11 Then
 Randomize: Rand = CInt(Rnd * 21)
 If Rand = 0 Then Ascii = 26: Exit Function
 If Rand = 1 Then Ascii = 27: Exit Function
 If Rand = 2 Then Ascii = 55: Exit Function
 If Rand = 3 Then Ascii = 73: Exit Function
 If Rand = 4 Then Ascii = 74: Exit Function
 If Rand = 5 Then Ascii = 76: Exit Function
 If Rand = 6 Then Ascii = 84: Exit Function
 If Rand = 7 Then Ascii = 89: Exit Function
 If Rand = 8 Then Ascii = 91: Exit Function
 If Rand = 9 Then Ascii = 93: Exit Function
 If Rand = 10 Then Ascii = 99: Exit Function
 If Rand = 11 Then Ascii = 102: Exit Function
 If Rand = 12 Then Ascii = 110: Exit Function
 If Rand = 13 Then Ascii = 115: Exit Function
 If Rand = 14 Then Ascii = 117: Exit Function
 If Rand = 15 Then Ascii = 180: Exit Function
 If Rand = 16 Then Ascii = 211: Exit Function
 If Rand = 17 Then Ascii = 213: Exit Function
 If Rand = 18 Then Ascii = 226: Exit Function
 If Rand = 19 Then Ascii = 230: Exit Function
 If Rand = 20 Then Ascii = 242: Exit Function
 If Rand = 21 Then Ascii = 243: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 12 Then
 Randomize: Rand = CInt(Rnd * 19)
 If Rand = 0 Then Ascii = 19: Exit Function
 If Rand = 1 Then Ascii = 36: Exit Function
 If Rand = 2 Then Ascii = 107: Exit Function
 If Rand = 3 Then Ascii = 111: Exit Function
 If Rand = 4 Then Ascii = 121: Exit Function
 If Rand = 5 Then Ascii = 122: Exit Function
 If Rand = 6 Then Ascii = 148: Exit Function
 If Rand = 7 Then Ascii = 149: Exit Function
 If Rand = 8 Then Ascii = 154: Exit Function
 If Rand = 9 Then Ascii = 159: Exit Function
 If Rand = 10 Then Ascii = 162: Exit Function
 If Rand = 11 Then Ascii = 176: Exit Function
 If Rand = 12 Then Ascii = 183: Exit Function
 If Rand = 13 Then Ascii = 198: Exit Function
 If Rand = 14 Then Ascii = 200: Exit Function
 If Rand = 15 Then Ascii = 205: Exit Function
 If Rand = 16 Then Ascii = 208: Exit Function
 If Rand = 17 Then Ascii = 235: Exit Function
 If Rand = 18 Then Ascii = 238: Exit Function
 If Rand = 19 Then Ascii = 240: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 13 Then
 Randomize: Rand = CInt(Rnd * 20)
 If Rand = 0 Then Ascii = 13: Exit Function
 If Rand = 1 Then Ascii = 24: Exit Function
 If Rand = 2 Then Ascii = 25: Exit Function
 If Rand = 3 Then Ascii = 29: Exit Function
 If Rand = 4 Then Ascii = 67: Exit Function
 If Rand = 5 Then Ascii = 86: Exit Function
 If Rand = 6 Then Ascii = 88: Exit Function
 If Rand = 7 Then Ascii = 90: Exit Function
 If Rand = 8 Then Ascii = 101: Exit Function
 If Rand = 9 Then Ascii = 104: Exit Function
 If Rand = 10 Then Ascii = 109: Exit Function
 If Rand = 11 Then Ascii = 129: Exit Function
 If Rand = 12 Then Ascii = 147: Exit Function
 If Rand = 13 Then Ascii = 151: Exit Function
 If Rand = 14 Then Ascii = 163: Exit Function
 If Rand = 15 Then Ascii = 164: Exit Function
 If Rand = 16 Then Ascii = 184: Exit Function
 If Rand = 17 Then Ascii = 197: Exit Function
 If Rand = 18 Then Ascii = 202: Exit Function
 If Rand = 19 Then Ascii = 207: Exit Function
 If Rand = 20 Then Ascii = 214: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 14 Then
 Randomize: Rand = CInt(Rnd * 17)
 If Rand = 0 Then Ascii = 11: Exit Function
 If Rand = 1 Then Ascii = 38: Exit Function
 If Rand = 2 Then Ascii = 52: Exit Function
 If Rand = 3 Then Ascii = 70: Exit Function
 If Rand = 4 Then Ascii = 75: Exit Function
 If Rand = 5 Then Ascii = 97: Exit Function
 If Rand = 6 Then Ascii = 119: Exit Function
 If Rand = 7 Then Ascii = 135: Exit Function
 If Rand = 8 Then Ascii = 142: Exit Function
 If Rand = 9 Then Ascii = 150: Exit Function
 If Rand = 10 Then Ascii = 152: Exit Function
 If Rand = 11 Then Ascii = 153: Exit Function
 If Rand = 12 Then Ascii = 165: Exit Function
 If Rand = 13 Then Ascii = 167: Exit Function
 If Rand = 14 Then Ascii = 171: Exit Function
 If Rand = 15 Then Ascii = 181: Exit Function
 If Rand = 16 Then Ascii = 210: Exit Function
 If Rand = 17 Then Ascii = 237: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 15 Then
 Randomize: Rand = CInt(Rnd * 19)
 If Rand = 0 Then Ascii = 12: Exit Function
 If Rand = 1 Then Ascii = 37: Exit Function
 If Rand = 2 Then Ascii = 42: Exit Function
 If Rand = 3 Then Ascii = 50: Exit Function
 If Rand = 4 Then Ascii = 51: Exit Function
 If Rand = 5 Then Ascii = 54: Exit Function
 If Rand = 6 Then Ascii = 57: Exit Function
 If Rand = 7 Then Ascii = 80: Exit Function
 If Rand = 8 Then Ascii = 83: Exit Function
 If Rand = 9 Then Ascii = 85: Exit Function
 If Rand = 10 Then Ascii = 112: Exit Function
 If Rand = 11 Then Ascii = 113: Exit Function
 If Rand = 12 Then Ascii = 128: Exit Function
 If Rand = 13 Then Ascii = 130: Exit Function
 If Rand = 14 Then Ascii = 137: Exit Function
 If Rand = 15 Then Ascii = 138: Exit Function
 If Rand = 16 Then Ascii = 172: Exit Function
 If Rand = 17 Then Ascii = 227: Exit Function
 If Rand = 18 Then Ascii = 228: Exit Function
 If Rand = 19 Then Ascii = 234: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 16 Then
 Randomize: Rand = CInt(Rnd * 19)
 If Rand = 0 Then Ascii = 15: Exit Function
 If Rand = 1 Then Ascii = 16: Exit Function
 If Rand = 2 Then Ascii = 17: Exit Function
 If Rand = 3 Then Ascii = 21: Exit Function
 If Rand = 4 Then Ascii = 79: Exit Function
 If Rand = 5 Then Ascii = 98: Exit Function
 If Rand = 6 Then Ascii = 100: Exit Function
 If Rand = 7 Then Ascii = 103: Exit Function
 If Rand = 8 Then Ascii = 132: Exit Function
 If Rand = 9 Then Ascii = 133: Exit Function
 If Rand = 10 Then Ascii = 136: Exit Function
 If Rand = 11 Then Ascii = 156: Exit Function
 If Rand = 12 Then Ascii = 160: Exit Function
 If Rand = 13 Then Ascii = 186: Exit Function
 If Rand = 14 Then Ascii = 187: Exit Function
 If Rand = 15 Then Ascii = 209: Exit Function
 If Rand = 16 Then Ascii = 225: Exit Function
 If Rand = 17 Then Ascii = 232: Exit Function
 If Rand = 18 Then Ascii = 233: Exit Function
 If Rand = 19 Then Ascii = 254: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 17 Then
 Randomize: Rand = CInt(Rnd * 13)
 If Rand = 0 Then Ascii = 53: Exit Function
 If Rand = 1 Then Ascii = 56: Exit Function
 If Rand = 2 Then Ascii = 72: Exit Function
 If Rand = 3 Then Ascii = 77: Exit Function
 If Rand = 4 Then Ascii = 78: Exit Function
 If Rand = 5 Then Ascii = 81: Exit Function
 If Rand = 6 Then Ascii = 127: Exit Function
 If Rand = 7 Then Ascii = 131: Exit Function
 If Rand = 8 Then Ascii = 145: Exit Function
 If Rand = 9 Then Ascii = 157: Exit Function
 If Rand = 10 Then Ascii = 158: Exit Function
 If Rand = 11 Then Ascii = 182: Exit Function
 If Rand = 12 Then Ascii = 185: Exit Function
 If Rand = 13 Then Ascii = 216: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 18 Then
 Randomize: Rand = CInt(Rnd * 14)
 If Rand = 0 Then Ascii = 4: Exit Function
 If Rand = 1 Then Ascii = 6: Exit Function
 If Rand = 2 Then Ascii = 30: Exit Function
 If Rand = 3 Then Ascii = 31: Exit Function
 If Rand = 4 Then Ascii = 35: Exit Function
 If Rand = 5 Then Ascii = 65: Exit Function
 If Rand = 6 Then Ascii = 68: Exit Function
 If Rand = 7 Then Ascii = 69: Exit Function
 If Rand = 8 Then Ascii = 71: Exit Function
 If Rand = 9 Then Ascii = 82: Exit Function
 If Rand = 10 Then Ascii = 87: Exit Function
 If Rand = 11 Then Ascii = 144: Exit Function
 If Rand = 12 Then Ascii = 166: Exit Function
 If Rand = 13 Then Ascii = 199: Exit Function
 If Rand = 14 Then Ascii = 201: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 19 Then
 Randomize: Rand = CInt(Rnd * 8)
 If Rand = 0 Then Ascii = 1: Exit Function
 If Rand = 1 Then Ascii = 5: Exit Function
 If Rand = 2 Then Ascii = 18: Exit Function
 If Rand = 3 Then Ascii = 20: Exit Function
 If Rand = 4 Then Ascii = 48: Exit Function
 If Rand = 5 Then Ascii = 134: Exit Function
 If Rand = 6 Then Ascii = 203: Exit Function
 If Rand = 7 Then Ascii = 204: Exit Function
 If Rand = 8 Then Ascii = 215: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 20 Then
 Randomize: Rand = CInt(Rnd * 3)
 If Rand = 0 Then Ascii = 14: Exit Function
 If Rand = 1 Then Ascii = 64: Exit Function
 If Rand = 2 Then Ascii = 66: Exit Function
 If Rand = 3 Then Ascii = 206: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 21 Then
 Randomize: Rand = CInt(Rnd * 2)
 If Rand = 0 Then Ascii = 3: Exit Function
 If Rand = 1 Then Ascii = 143: Exit Function
 If Rand = 2 Then Ascii = 146: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col = 22 Or Col = 23 Then Ascii = 23: Exit Function
''''''''''''''''''''''''''''''''''''''
If Col = 24 Or Col = 25 Then
 Randomize: Rand = CInt(Rnd * 4)
 If Rand = 0 Then Ascii = 177: Exit Function
 If Rand = 1 Then Ascii = 220: Exit Function
 If Rand = 2 Then Ascii = 221: Exit Function
 If Rand = 3 Then Ascii = 222: Exit Function
 If Rand = 4 Then Ascii = 223: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col >= 26 And Col < 32 Then Ascii = 2: Exit Function
''''''''''''''''''''''''''''''''''''''
If Col >= 32 And Col < 38 Then
 Randomize: Rand = CInt(Rnd * 1)
 If Rand = 0 Then Ascii = 10: Exit Function
 If Rand = 1 Then Ascii = 178: Exit Function
End If
''''''''''''''''''''''''''''''''''''''
If Col >= 38 And Col < 42 Then Ascii = 8: Exit Function
''''''''''''''''''''''''''''''''''''''
If Col >= 42 And Col <= 48 Then Ascii = 219: Exit Function
End Function
