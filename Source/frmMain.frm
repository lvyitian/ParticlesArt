VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ParticlesArt"
   ClientHeight    =   9525
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13995
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   635
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   933
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picImageView 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9525
      Left            =   2895
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   633
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   738
      TabIndex        =   10
      Top             =   0
      Width           =   11100
      Begin VB.PictureBox picSrc 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4515
         Left            =   0
         Picture         =   "frmMain.frx":424A
         ScaleHeight     =   301
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   501
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   7515
      End
      Begin VB.PictureBox picImage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   1320
         MousePointer    =   15  'Size All
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   72
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   11
         Top             =   1320
         Width           =   960
      End
   End
   Begin VB.PictureBox picLM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":754B
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picVAdjust 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   0
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   0
      ScaleMode       =   0  'User
      ScaleWidth      =   1033.223
      TabIndex        =   8
      Top             =   9525
      Width           =   13995
   End
   Begin VB.Timer tmrLoadDragdrop 
      Enabled         =   0   'False
      Left            =   7080
      Top             =   3960
   End
   Begin VB.PictureBox picOutput 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   13995
      TabIndex        =   7
      Top             =   9525
      Width           =   13995
   End
   Begin VB.PictureBox picCtrlPan 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9525
      Left            =   0
      ScaleHeight     =   635
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.FileListBox FilesCount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Frame Frames 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "位置"
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   2415
         Begin VB.TextBox txtXOff 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   360
            TabIndex        =   24
            Text            =   "-64"
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtYOff 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   360
            TabIndex        =   23
            Text            =   "127"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtZOff 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   360
            TabIndex        =   22
            Text            =   "-64"
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton DirX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "X"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   21
            Top             =   1320
            Width           =   495
         End
         Begin VB.OptionButton DirY 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Y"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   20
            Top             =   1320
            Width           =   495
         End
         Begin VB.OptionButton DirZ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Z"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   19
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Labels 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "X:"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   180
         End
         Begin VB.Label Labels 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Y:"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   180
         End
         Begin VB.Label Labels 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Z:"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   180
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "朝向"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   30
            TabIndex        =   25
            Top             =   1300
            Width           =   735
         End
      End
      Begin VB.Frame Frames 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "大小"
         ForeColor       =   &H80000008&
         Height          =   2175
         Index           =   2
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   2415
         Begin VB.Frame Frames 
            BackColor       =   &H80000005&
            Height          =   450
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   1600
            Width           =   2200
            Begin VB.Label cmdApplySizing 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "应用"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   30
               TabIndex        =   14
               Top             =   130
               Width           =   2145
            End
         End
         Begin VB.TextBox txtWidth 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   4
            Text            =   "300"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtHeight 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   6
            Text            =   "500"
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox ChkRatio 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "保持宽高比"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   1350
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.Label HeightBlock 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   17
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label WidthBlock 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   16
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "宽:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "高:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   585
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser GifView 
      Height          =   16455
      Left            =   2880
      TabIndex        =   15
      Top             =   0
      Width           =   11175
      ExtentX         =   19711
      ExtentY         =   29025
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu File 
      Caption         =   "文件"
      Begin VB.Menu OpenF 
         Caption         =   "打开"
         Begin VB.Menu OpenImage 
            Caption         =   "图片"
         End
         Begin VB.Menu OpenAnimation 
            Caption         =   "视频/GIF"
         End
      End
      Begin VB.Menu Save 
         Caption         =   "导出为mcfunction文件"
      End
      Begin VB.Menu DivingLine 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMe 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu PlayGif 
      Caption         =   "播放"
      Enabled         =   0   'False
   End
   Begin VB.Menu StopGif 
      Caption         =   "暂停"
      Visible         =   0   'False
   End
   Begin VB.Menu Setting 
      Caption         =   "设置"
      Begin VB.Menu FrameSetting 
         Caption         =   "帧数"
         Enabled         =   0   'False
         Begin VB.Menu Frame_24 
            Caption         =   "24Fps"
         End
         Begin VB.Menu Frame_30 
            Caption         =   "30Fps"
         End
         Begin VB.Menu Frame_60 
            Caption         =   "60Fps"
         End
      End
      Begin VB.Menu GameVersion 
         Caption         =   "游戏版本"
         Begin VB.Menu Version1_12 
            Caption         =   "1.12"
         End
         Begin VB.Menu Version1_13 
            Caption         =   "1.13以上"
         End
      End
      Begin VB.Menu Language 
         Caption         =   "语言"
         Begin VB.Menu English 
            Caption         =   "English"
         End
         Begin VB.Menu Chinese 
            Caption         =   "简体中文"
         End
      End
   End
   Begin VB.Menu About 
      Caption         =   "关于"
      Begin VB.Menu ParticlesArt 
         Caption         =   "ParticlesArt"
      End
      Begin VB.Menu AA55 
         Caption         =   "0xAA55论坛"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LangMsg(16) As String
Dim frmh As Long, frmw As Long
Dim FPS As Long
Dim ImageFile As String, OpenFileType As String, SaveFileTitle As String
Dim GenGif As Boolean
Dim Folder As String
Dim PathFunction As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As Any) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As Any) As Long
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_NOVALIDATE = &H100
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type RGB32
    R As Byte
    G As Byte
    B As Byte
    X As Byte
End Type

Private Type RGB24
    B As Byte
    G As Byte
    R As Byte
End Type

Private Type Color_t
    R As Long
    G As Long
    B As Long
End Type

Private Type Color_w
    R As Single
    G As Single
    B As Single
End Type

Private IsGif As Boolean
Private IsAnimation As Boolean

Private m_DragdroppedFileName As String
Private m_Width As Long, m_Height As Long
Private m_WidthEdit As Boolean, m_HeightEdit As Boolean

'Private Const m_Rhythm As Single = 0.075

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Const DIB_RGB_COLORS = 0
Private Const HALFTONE = 4

'加载图像
Private Sub Load_Picture(Path As String)

On Error GoTo ErrHandler

Dim Pic As StdPicture
Set Pic = LoadPicture(Path)

m_Width = ScaleX(Pic.Width, vbHimetric, vbPixels)
m_Height = ScaleY(Pic.Height, vbHimetric, vbPixels)
txtWidth.Text = m_Width
txtHeight.Text = m_Height

Set picSrc.Picture = Pic
picImage.Visible = False
picImage.Cls
picImage.Move 0, 0, m_Width, m_Height
picImage.Visible = True
picImage_Resize
picImage.PaintPicture picSrc.Picture, 0, 0, picImage.ScaleWidth, picImage.ScaleHeight

Exit Sub
ErrHandler:
MsgBox LangMsg(0) & vbCrLf & "(" & Err.Number & ")" & Err.Description, vbExclamation, LangMsg(1)
End Sub

Private Sub AA55_Click()
Call ShellExecute(hWnd, "open", "https://www.0xaa55.com", vbNullString, vbNullString, &H0)
End Sub

Private Sub SetChinese()
File.Caption = "文件"
OpenF.Caption = "打开"
Save.Caption = "导出为mcfunction文件"
ExitMe.Caption = "退出"
Setting.Caption = "设置"
Language.Caption = "语言"
About.Caption = "关于"
AA55.Caption = "0xAA55论坛"
GameVersion.Caption = "游戏版本"
PlayGif.Caption = "播放"
StopGif.Caption = "暂停"
FrameSetting.Caption = "帧数"
OpenImage.Caption = "图片"
OpenAnimation.Caption = "视频/GIF"
Version1_13.Caption = "1.13及以上"
Frames(1).Caption = "位置"
Frames(2).Caption = "大小"
Labels(5).Caption = "宽:"
Labels(6).Caption = "高:"
Labels(3).Caption = "朝向"
ChkRatio.Caption = "保持宽高比"
cmdApplySizing.Caption = "应用"
LangMsg(0) = "错误：加载图片失败。"
LangMsg(1) = "加载图片失败"
LangMsg(2) = "错误：输入的图片大小无效。"
LangMsg(3) = "不支持此大小"
LangMsg(4) = "仍可以加载图片"
LangMsg(5) = "按下 [确定]下载，[取消] 忽略"
LangMsg(6) = "错误：ffmpeg.exe丢失。"
LangMsg(7) = "加载错误"
LangMsg(8) = "图像文件 (.bmp .jpg .jpeg)|*.bmp;*.jpg;*.jpeg|所有文件|*.*|"
LangMsg(9) = "选择文件"
LangMsg(12) = "错误：操作错误。"
LangMsg(13) = "方块"
LangMsg(14) = "视频文件 (.wmv .avi .asf .mpeg .mpg .rm .rmvb .ram .flv .mp4 .3gp .mov .divx .dv .vob .mkv .qt .cpk .fli .flc .f4v .m4v .mod .m2t .webm .mts .m2ts .3g2 .mpe .ts .div .lavf .dirac)|*.wmv;*.avi;*.asf;*.mpeg;*.mpg;*.rm;*.rmvb;*.ram;*.flv;*.mp4;*.3gp;*.mov;*.divx;*.dv;*.vob;*.mkv;*.qt;*.cpk;*.fli;*.flc;*.f4v;*.m4v;*.mod;*.m2t;*.webm;*.mts;*.m2ts;*.3g2;*.mpe;*.ts;*.div;*.lavf;*.dirac|GIF文件 (.gif)|*.gif|所有文件|*.*|"
LangMsg(15) = "选择functions文件夹(存档\data\functions)"
LangMsg(16) = "选择datapacks文件夹(存档\datapacks)"
WidthBlock.Caption = txtWidth.Text / 5 & LangMsg(13)
HeightBlock.Caption = txtHeight.Text / 5 & LangMsg(13)
End Sub

Private Sub Chinese_Click()
English.Checked = False
Chinese.Checked = True
SetChinese
End Sub

Private Sub cmdApplySizing_Click()
Dim NewWidth As Long, NewHeight As Long
NewWidth = Val(txtWidth.Text)
NewHeight = Val(txtHeight.Text)

If NewWidth = 0 Or NewHeight = 0 Then
    MsgBox LangMsg(2), vbExclamation, LangMsg(3)
    Exit Sub
End If

GifView.Visible = False
picImage.Move picImage.Left + picImage.Width \ 2 - NewWidth \ 2, picImage.Top + picImage.Height \ 2 - NewHeight \ 2, NewWidth, NewHeight
picImage_Resize
picImage.Cls
picImage.PaintPicture picSrc.Picture, 0, 0, picImage.ScaleWidth, picImage.ScaleHeight
If GenGif = True Or IsGif Then
GenGif = False
PlayGif_Click
End If
End Sub

Private Sub Freeze()
On Error Resume Next

Dim Ctrl As Control
For Each Ctrl In Controls
    Ctrl.Enabled = False
Next
End Sub

Private Sub UnFreeze()
On Error Resume Next

Dim Ctrl As Control
For Each Ctrl In Controls
    Ctrl.Enabled = True
Next
End Sub

Private Sub GenEnv(Path As String)
Dim PathMain As String, I As Long

If Version1_13.Checked Then

If Dir(Path & "\ParticlesArt\", vbDirectory) = "" Then MkDir (Path & "\ParticlesArt\")
If Dir(Path & "\ParticlesArt\data\", vbDirectory) = "" Then MkDir (Path & "\ParticlesArt\data\")

Open Path & "\ParticlesArt\pack.mcmeta" For Output Access Write As #1
Print #1, "{"
Print #1, "   ""pack"": {"
Print #1, "      ""pack_format"": 1,"
Print #1, "      ""description"": ""ParticlesArt"""
Print #1, "      }"
Print #1, "}"
Close #1

If Not Dir(Path & "\ParticlesArt\data\particlesart\", vbDirectory) = "" Then
    Do
    I = I + 1
    PathMain = "\ParticlesArt\data\particlesart" & I
    PathFunction = "particlesart" & I
    Loop Until Dir(Path & PathMain, vbDirectory) = ""
    Folder = Path & PathMain & "\functions\"
    MkDir (Path & PathMain)
    MkDir (Folder)
Else
    Folder = Path & "\ParticlesArt\data\particlesart\functions\"
    PathFunction = "particlesart"
    MkDir (Path & "\ParticlesArt\data\particlesart\")
    MkDir (Folder)
End If

Else

If Not Dir(Path & "\particlesart\", vbDirectory) = "" Then
    Do
    I = I + 1
    PathMain = "\particlesart" & I
    PathFunction = "particlesart" & I
    Loop Until Dir(Path & PathMain, vbDirectory) = ""
    Folder = Path & PathMain & "\"
    MkDir (Folder)
Else
    Folder = Path & "\particlesart\"
    PathFunction = "particlesart"
    MkDir (Folder)
End If

End If
End Sub

Private Sub AnimationGen()
Dim I As Long

    FilesCount.Path = App.Path & "\image"
For I = 1 To FilesCount.ListCount
    Load_Picture App.Path & "\image\frame" & I & ".jpg"
    ImageGen Folder & "frame" & I & ".mcfunction"
Next

    Open Folder & "main.mcfunction" For Output Access Write As #1
    Print #1, "scoreboard objectives add Frame dummy"
For I = 1 To FilesCount.ListCount
    If Version1_13.Checked Then
        Print #1, "execute if score @p Frame matches " & I & " run function " & PathFunction & ":frame" & I
    Else
        Print #1, "execute @a[score_Frame=" & I & ",score_Frame_min=" & I & "] ~ ~ ~ function " & PathFunction & ":frame" & I
    End If
Next
    If Version1_13.Checked Then
        Print #1, "execute if score @p Frame matches " & I + 1 & " run scoreboard players set @a Frame 0"
    Else
        Print #1, "execute @a[score_Frame=" & I + 1 & ",score_Frame_min=" & I + 1 & "] ~ ~ ~ scoreboard players set @a Frame 0"
    End If
    Print #1, "scoreboard players add @a Frame 1"
    Close #1
End Sub
Private Sub ImageGen(Path As String)
    Dim BmpWidth As Long, BmpHeight As Long
    BmpWidth = picImage.ScaleWidth
    BmpHeight = picImage.ScaleHeight
    Dim Z As Long, X As Long
    Dim I&
    Dim WritePix As Color_w
    Dim ThisPix As Color_t '当前像素颜色值
    Dim XOff As Long, YOff As Long, ZOff As Long
    Dim CIPtr As Long, Temp As Byte
    
    XOff = Val(txtXOff.Text)
    YOff = Val(txtYOff.Text)
    ZOff = Val(txtZOff.Text)
    
    Freeze
    
    Dim hTempDC24 As Long, Ptr24 As Long, Pitch24 As Long, Line24() As RGB24
    hTempDC24 = CreateDIB24(Ptr24, BmpWidth, BmpHeight)
    Pitch24 = CalcPitch(24, BmpWidth)
    ReDim Line24(BmpWidth - 1)
    Open Path For Output Access Write As #1
    
    '搬运原图到自己的DIB
    BitBlt hTempDC24, 0, 0, BmpWidth, BmpHeight, picImage.hDC, 0, 0, vbSrcCopy
    
    For Z = 0 To picImage.ScaleHeight - 1
        CIPtr = (picImage.ScaleHeight - 1 - Z) * BmpWidth
        '复制一行源RGB
        CopyMemory Line24(0), ByVal Ptr24, BmpWidth * 3
            For X = 0 To picImage.ScaleWidth - 1
            If Line24(X).R = 0 Then WritePix.R = 0.001 Else WritePix.R = Round(Line24(X).R / 255, 3)
            If Line24(X).G = 0 Then WritePix.G = 0.001 Else WritePix.G = Round(Line24(X).G / 255, 3)
            If Line24(X).B = 0 Then WritePix.B = 0.001 Else WritePix.B = Round(Line24(X).B / 255, 3)
            If DirX.Value Then
            '循环中浮点计算通病
                If Version1_13.Checked Then
                Print #1, "particle minecraft:dust " & WritePix.R & " " & WritePix.G & " " & WritePix.B & " 1 " & XOff & " " & Format(Round(0.1 + YOff + 0.2 * Z, 2), ".00") & " " & Format(Round(0.2 * X + 0.1 + ZOff, 2), ".00") & " ~ ~ ~ 1 0 force"
                Else
                Print #1, "particle reddust " & XOff & " " & Round(0.1 + YOff + 0.2 * Z, 1) & " " & Round(0.2 * X + 0.1 + ZOff, 1) & " " & WritePix.R & " " & WritePix.G & " " & WritePix.B & " 1 0 force"
                End If
            ElseIf DirY.Value Then
                If Version1_13.Checked Then
                    Print #1, "particle minecraft:dust " & WritePix.R & " " & WritePix.G & " " & WritePix.B & " 1 " & Format(Round(0.1 + XOff - 0.2 * X, 2), ".00") & " " & YOff & " " & Format(Round(0.2 * Z + 0.1 + ZOff, 2), ".00") & " ~ ~ ~ 1 0 force"
                Else
                    Print #1, "particle reddust " & Round(0.1 + XOff - 0.2 * X, 1) & " " & YOff & " " & Round(0.2 * Z + 0.1 + ZOff, 1) & " " & WritePix.R & " " & WritePix.G & " " & WritePix.B & " 1 0 force"
                End If
            ElseIf DirZ.Value Then
                If Version1_13.Checked Then
                    Print #1, "particle minecraft:dust " & WritePix.R & " " & WritePix.G & " " & WritePix.B & " 1 " & Format(Round(0.2 * X + 0.1 + XOff, 2), ".00") & " " & Format(Round(0.1 + YOff + 0.2 * Z, 2), ".00") & " " & ZOff & " ~ ~ ~ 1 0 force"
                Else
                    Print #1, "particle reddust " & Round(0.2 * X + 0.1 + XOff, 1) & " " & Round(0.1 + YOff + 0.2 * Z, 1) & " " & ZOff & " " & WritePix.R & " " & WritePix.G & " " & WritePix.B & " 1 0 force"
                End If
            End If
            Next
            '复制索引颜色
        CopyMemory ByVal Ptr24, Line24(0), BmpWidth * 3
        
        Ptr24 = Ptr24 + Pitch24
    Next
    
    '擦屁股
    DeleteDC hTempDC24
    '套娃
    
    Close #1
    
    UnFreeze

End Sub

Private Sub SetEnglish()
File.Caption = "File"
OpenF.Caption = "Open"
Save.Caption = "Export .mcfunction"
ExitMe.Caption = "Exit"
Setting.Caption = "Setting"
Language.Caption = "Language"
About.Caption = "About"
AA55.Caption = "0xAA55 Forum"
GameVersion.Caption = "GameVersion"
PlayGif.Caption = "Play"
StopGif.Caption = "Stop"
FrameSetting.Caption = "FPS"
OpenImage.Caption = "Image"
OpenAnimation.Caption = "Video/GIF"
Version1_13.Caption = "1.13 or later"
Frames(1).Caption = "Position"
Frames(2).Caption = "Size"
Labels(5).Caption = "Width:"
Labels(6).Caption = "Height:"
Labels(3).Caption = "Toward"
ChkRatio.Caption = "Keep aspect ratio"
cmdApplySizing.Caption = "Apply"
LangMsg(0) = "Error: Failed to load image."
LangMsg(1) = "Failed to load image"
LangMsg(2) = "Error: Invalid size entered."
LangMsg(3) = "Unsupported size"
LangMsg(4) = "Pictures can still be loaded."
LangMsg(5) = "Press [OK] to download,[Cancel] to ignore"
LangMsg(6) = "Error: ffmpeg.exe missing"
LangMsg(7) = "Loading error"
LangMsg(8) = "Image file (.bmp .jpg .jpeg)|*.bmp;*.jpg;*.jpeg|All|*.*|"
LangMsg(9) = "Select file"
LangMsg(12) = "Error: Operation error."
LangMsg(13) = "Blocks"
LangMsg(14) = "Video File (.wmv .avi .asf .mpeg .mpg .rm .rmvb .ram .flv .mp4 .3gp .mov .divx .dv .vob .mkv .qt .cpk .fli .flc .f4v .m4v .mod .m2t .webm .mts .m2ts .3g2 .mpe .ts .div .lavf .dirac)|*.wmv;*.avi;*.asf;*.mpeg;*.mpg;*.rm;*.rmvb;*.ram;*.flv;*.mp4;*.3gp;*.mov;*.divx;*.dv;*.vob;*.mkv;*.qt;*.cpk;*.fli;*.flc;*.f4v;*.m4v;*.mod;*.m2t;*.webm;*.mts;*.m2ts;*.3g2;*.mpe;*.ts;*.div;*.lavf;*.dirac|GIF File (.gif)|*.gif|All|*.*|"
LangMsg(15) = "Select functions folder (world\data\functions)"
LangMsg(16) = "Select datapacks folder (world\datapacks)"
WidthBlock.Caption = txtWidth.Text / 5 & LangMsg(13)
HeightBlock.Caption = txtHeight.Text / 5 & LangMsg(13)
End Sub

Private Sub English_Click()
English.Checked = True
Chinese.Checked = False
SetEnglish
End Sub

Private Sub ExitMe_Click()
End
End Sub
Public Function GetFolder() As String
Dim OutFile As String
OutFile = SaveFile()
If OutFile = "" Then Exit Function
GetFolder = Mid(OutFile, 1, InStrRev(OutFile, "\") - 1)
End Function

Private Sub Frame_24_Click()
Frame_24.Checked = True
Frame_30.Checked = False
Frame_60.Checked = False
FPS = 24
If GenGif = True Or IsGif Then
GenGif = False
PlayGif_Click
End If
End Sub

Private Sub Frame_30_Click()
Frame_24.Checked = False
Frame_30.Checked = True
Frame_60.Checked = False
FPS = 30
If GenGif = True Or IsGif Then
GenGif = False
PlayGif_Click
End If
End Sub

Private Sub Frame_60_Click()
Frame_24.Checked = False
Frame_30.Checked = False
Frame_60.Checked = True
FPS = 60
If GenGif = True Or IsGif Then
GenGif = False
PlayGif_Click
End If
End Sub

Private Sub ParticlesArt_Click()
Call ShellExecute(hWnd, "open", "https://github.com/Tao0Lu/ParticlesArt", vbNullString, vbNullString, &H0)
End Sub

Private Sub Form_Load()

Dim I As Long
Dim Path As String
Dim Image As String
Path = App.Path

If GetSystemDefaultLCID = &H804 Then
Chinese.Checked = True
SetChinese
Else
English.Checked = True
SetEnglish
End If

If Dir(App.Path & "\ffmpeg.exe") = "" Then
    If MsgBox(LangMsg(6) & vbCrLf & LangMsg(4) & vbCrLf & LangMsg(5) & Err.Description, vbInformation Or vbOKCancel, LangMsg(7)) = vbCancel Then
    OpenAnimation.Enabled = False
    Else
        Call ShellExecute(hWnd, "open", "https://ffmpeg.zeranoe.com/builds/", vbNullString, vbNullString, &H0)
        End
    End If
End If
Show

'初始化
Frame_30.Checked = True
FPS = 30
Version1_13.Checked = True
SaveFileTitle = LangMsg(16)
DirY.Value = True

m_Width = picSrc.Width
m_Height = picSrc.Height
picImage.Move 0, 0, m_Width, m_Height
picImage_Resize
txtWidth.Text = m_Width
txtHeight.Text = m_Height
picImage.Cls
picImage.PaintPicture picSrc.Picture, 0, 0, picImage.ScaleWidth, picImage.ScaleHeight
End Sub
Private Sub Form_Resize()
On Error Resume Next
picImageView.Width = ScaleWidth - picImageView.Left
GifView.Width = ScaleWidth - GifView.Left
End Sub
Function OpenFile()
Dim OFN As OPENFILENAME
With OFN
    .lStructSize = Len(OFN)
    .hwndOwner = hWnd
    .lpstrFilter = Replace(OpenFileType, "|", vbNullChar)
    .nFilterIndex = 1
    .lpstrFile = String(256, 0)
    .nMaxFile = 256
    .lpstrTitle = LangMsg(9)
    .flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXTENSIONDIFFERENT
End With
If GetOpenFileName(ByVal VarPtr(OFN)) Then
'这里不弄ByVal VarPtr的话……VB会把这个结构体里的所有字符串转换为ANSI
    OpenFile = Trim$(Replace(OFN.lpstrFile, vbNullChar, ""))
End If
End Function

Private Sub LimitImageBoxPosition(X As Long, Y As Long)
If picImage.Width <= picImageView.ScaleWidth Then
    X = (picImageView.ScaleWidth - picImage.Width) \ 2
Else
    If X > 0 Then X = 0
    If X < picImageView.ScaleWidth - picImage.Width Then X = picImageView.ScaleWidth - picImage.Width
End If
If picImage.Height <= picImageView.ScaleHeight Then
    Y = (picImageView.ScaleHeight - picImage.Height) \ 2
Else
    If Y > 0 Then Y = 0
    If Y < picImageView.ScaleHeight - picImage.Height Then Y = picImageView.ScaleHeight - picImage.Height
End If
End Sub

Private Function SaveFile() As String
Dim rtn As Long, pos As Integer
Dim Save As OPENFILENAME
    Save.lStructSize = Len(Save)
    Save.hwndOwner = hWnd
    Save.hInstance = App.hInstance
    Save.lpstrFilter = Replace("文件夹|*.~#~", "|", vbNullChar)
    Save.lpstrFile = "Directories" & String$(255 - Len("Image"), 0)
    Save.nMaxFile = 255
    Save.lpstrTitle = SaveFileTitle
    Save.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_OVERWRITEPROMPT
    rtn = GetSaveFileName(Save)

    If rtn > 0 Then
        pos = InStr(Save.lpstrFile, Chr$(0))
        If pos > 0 Then
            SaveFile = Left$(Save.lpstrFile, pos - 1)
        End If
    End If
    
    Exit Function
    
End Function

Private Sub OpenAnimation_Click()
Dim ImageType As String
Dim I As Long, R As Long, P As Long
OpenFileType = LangMsg(14)
ImageFile = OpenFile
If ImageFile = "" Then Exit Sub
Clean
IsAnimation = True
PlayGif.Enabled = True
FrameSetting.Enabled = True
GenGif = False
ImageType = Right(ImageFile, Len(ImageFile) - InStrRev(ImageFile, "."))
If LCase(ImageType) = "gif" Then
    IsGif = True
    Load_Picture ImageFile
    GifView.ZOrder 0
    GifView.Visible = True
    GifView.Navigate (ImageFile)
    PlayGif.Visible = False
    StopGif.Visible = True
Else
    I = Shell(App.Path & "\ffmpeg.exe -i " & ImageFile & " -ss 1 " & App.Path & "\show.jpg -y", vbHide)
    P = OpenProcess(SYNCHRONIZE, False, I)
    R = WaitForSingleObject(P, INFINITE)
    R = CloseHandle(P)
    Load_Picture App.Path & "\show.jpg"
End If
End Sub

Private Sub OpenImage_Click()
OpenFileType = LangMsg(8)
ImageFile = OpenFile
If ImageFile = "" Then Exit Sub
Clean
PlayGif.Enabled = False
FrameSetting.Enabled = False
GifView.Visible = False
Load_Picture ImageFile
End Sub

Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static DragX As Single, DragY As Single
If Button And 1 Then
    Dim NewX As Long, NewY As Long
    NewX = picImage.Left + X - DragX
    NewY = picImage.Top + Y - DragY
    LimitImageBoxPosition NewX, NewY
    picImage.Move NewX, NewY
Else
    DragX = X
    DragY = Y
End If
End Sub

Private Sub picImage_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count Then
    picImage.Cls
    m_DragdroppedFileName = Data.Files(1)
    tmrLoadDragdrop.Enabled = True
    tmrLoadDragdrop.Interval = 1
End If
End Sub

Private Sub picImage_Resize()
Dim X&, Y&

X = picImage.Left
Y = picImage.Top

LimitImageBoxPosition X, Y

picImage.Move X, Y
End Sub

Private Sub picImageView_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
picImage_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub picImageView_Resize()
picImage_Resize
End Sub

Private Sub picVAdjust_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Static DragY As Single
If Button And 1 Then
    Dim NewHeight As Single
    NewHeight = picOutput.Height + DragY - Y
    If NewHeight > ScaleHeight - 24 Then NewHeight = ScaleHeight - 24
    If NewHeight < 8 Then NewHeight = 8
    picOutput.Height = NewHeight
Else
    DragY = Y
End If
End Sub

Private Function CreateDIB24(PtrOut As Long, ByVal Width, ByVal Height) As Long
Dim BMIF As BITMAPINFOHEADER
With BMIF
    .biSize = 40
    .biWidth = Width
    .biHeight = Height
    .biPlanes = 1
    .biBitCount = 24
End With

Dim hTempDC As Long
hTempDC = CreateCompatibleDC(hDC)

Dim hDIB As Long
hDIB = CreateDIBSection(hDC, BMIF, DIB_RGB_COLORS, PtrOut, 0, 0)
DeleteObject SelectObject(hTempDC, hDIB)
DeleteObject hDIB

CreateDIB24 = hTempDC
End Function

Private Function CalcPitch(ByVal Bits As Long, ByVal Width As Long) As Long
CalcPitch = (((Width * Bits - 1) \ 32) + 1) * 4
End Function

Private Sub StopGif_Click()
GifView.Navigate ""
GifView.Visible = False
PlayGif.Visible = True
StopGif.Visible = False
End Sub
Private Sub PlayGif_Click()
Dim T_Width As Long, T_Height As Long
Dim I As Long, R As Long, P As Long
If Not GenGif Then
T_Width = Val(txtWidth.Text)
T_Height = Val(txtHeight.Text)
I = Shell(App.Path & "\ffmpeg.exe -i " & ImageFile & " -s " & T_Width & "*" & T_Height & " -r " & FPS & " " & App.Path & "\show.gif -y", vbHide)
P = OpenProcess(SYNCHRONIZE, False, I)
R = WaitForSingleObject(P, INFINITE)
R = CloseHandle(P)
GenGif = True
End If
GifView.Visible = True
GifView.ZOrder 0
If IsGif And Not GenGif Then GifView.Navigate (ImageFile) Else GifView.Navigate (App.Path & "\show.gif")
PlayGif.Visible = False
StopGif.Visible = True
End Sub

Private Sub Save_Click()
Dim Path As String
Dim T_Width As Long, T_Height As Long
Dim I As Long, R As Long, P As Long
If Version1_13.Checked Then
SaveFileTitle = LangMsg(16)
Else
SaveFileTitle = LangMsg(15)
End If
Path = GetFolder()
If Path = "" Then Exit Sub
GenEnv Path

If IsAnimation Then
    T_Width = Val(txtWidth.Text)
    T_Height = Val(txtHeight.Text)
    If Dir(App.Path & "\image\", vbDirectory) = "" Then MkDir (App.Path & "\image\")
        I = Shell(App.Path & "\ffmpeg.exe -i " & ImageFile & " -s " & T_Width & "*" & T_Height & " -r " & FPS & " " & App.Path & "\image\frame%d.jpg -y", vbHide)
        P = OpenProcess(SYNCHRONIZE, False, I)
        R = WaitForSingleObject(P, INFINITE)
        R = CloseHandle(P)
        AnimationGen
    Else
        ImageGen Folder & "image.mcfunction"
 End If
End Sub



Private Sub txtWidth_Change()
Dim NewWidth As Long

If m_HeightEdit Or txtWidth.Text = "" Then
    m_HeightEdit = False
    Exit Sub
End If

m_WidthEdit = True

NewWidth = Val(txtWidth.Text)
If NewWidth Then
    If ChkRatio.Value Then
        txtHeight.Text = m_Height * NewWidth \ m_Width
    End If
End If

WidthBlock.Caption = txtWidth.Text / 5 & LangMsg(13)
HeightBlock.Caption = txtHeight.Text / 5 & LangMsg(13)

End Sub

Private Sub txtHeight_Change()
Dim NewHeight As Long

If m_WidthEdit Or txtHeight.Text = "" Then
    m_WidthEdit = False
    Exit Sub
End If

m_HeightEdit = True

NewHeight = Val(txtHeight.Text)
If NewHeight Then
    If ChkRatio.Value Then
        txtWidth.Text = m_Width * NewHeight \ m_Height
    End If
End If

WidthBlock.Caption = txtWidth.Text / 5 & LangMsg(13)
HeightBlock.Caption = txtHeight.Text / 5 & LangMsg(13)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Clean
End Sub
Private Sub Clean()
If Not Dir(App.Path & "\show.jpg") = "" Then Kill App.Path & "\show.jpg"
If Not Dir(App.Path & "\show.gif") = "" Then Kill App.Path & "\show.gif"

If Not Dir(App.Path & "\image\", vbDirectory) = "" Then
FilesCount.Path = App.Path & "\image"
If Not FilesCount.ListCount = 0 Then
Kill App.Path & "\image\*.*"
RmDir App.Path & "\image\"
Else
RmDir App.Path & "\image\"
End If
End If

GifView.Navigate ""
End Sub

Private Sub Version1_12_Click()
Version1_12.Checked = True
Version1_13.Checked = False
End Sub

Private Sub Version1_13_Click()
Version1_12.Checked = False
Version1_13.Checked = True
End Sub
