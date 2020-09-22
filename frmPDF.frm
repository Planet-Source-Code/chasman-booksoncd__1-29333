VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{CA8A9783-280D-11CF-A24D-444553540000}#1.3#0"; "pdf.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPDF 
   BackColor       =   &H80000004&
   Caption         =   "PDF / HTML / TEXT / RTF eBook File Viewer - Click on a Book to View"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPDF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleMode       =   0  'User
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser htmlBrowser 
      Height          =   7065
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1110
      Width           =   11265
      ExtentX         =   19870
      ExtentY         =   12462
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
   Begin RichTextLib.RichTextBox txtViewer 
      Height          =   7065
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1110
      Visible         =   0   'False
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   12462
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmPDF.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.DirListBox dirCurrent 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   5640
      TabIndex        =   1
      Top             =   30
      Width           =   2655
   End
   Begin VB.FileListBox lstFiles 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   0
      Pattern         =   "*.pdf"
      TabIndex        =   0
      ToolTipText     =   "Click a Book to View"
      Top             =   30
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   1140
      Left            =   8310
      TabIndex        =   4
      Top             =   -60
      Width           =   2925
      Begin VB.OptionButton optTEXT 
         BackColor       =   &H80000004&
         Caption         =   "TEXT/RTF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1710
         TabIndex        =   10
         Top             =   210
         Width           =   1155
      End
      Begin VB.OptionButton optHTML 
         BackColor       =   &H80000004&
         Caption         =   "HTML"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   810
         TabIndex        =   8
         Top             =   210
         Width           =   915
      End
      Begin VB.OptionButton optPDF 
         BackColor       =   &H80000004&
         Caption         =   "PDF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   645
      End
      Begin VB.DriveListBox drvTarget 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   450
         Width           =   1065
      End
      Begin VB.Label lblBookCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "PDF eBooks in list : ???"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   780
         Width           =   2745
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Current Drive :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   480
         Width           =   1485
      End
   End
   Begin PdfLib.Pdf pdfViewer 
      Height          =   7065
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1110
      Width           =   11235
      _Version        =   131072
      _ExtentX        =   19817
      _ExtentY        =   12462
      _StockProps     =   0
      SRC             =   ""
   End
   Begin VB.Label lblOwner 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "created 4u by Chas. Meyer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   11250
      TabIndex        =   9
      Top             =   60
      Width           =   795
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   14250
      Picture         =   "frmPDF.frx":04BD
      Top             =   210
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   13350
      Picture         =   "frmPDF.frx":08FF
      Top             =   210
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   12390
      Picture         =   "frmPDF.frx":0D41
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'i have alot of ebooks/manuals in pdf, html form and it was
'a huge hassle to go thru the steps of opening the next file
'so, i put this together, hope some of you find this helpful
'please let me know whacha think, changes, ideas, etc...
'i put the pdf.ocx files in the zip, but psc might not let it thru
'chasman7@excite.com

'add to project these components:
'Acrobat Control for ActiveX - pdf.ocx - adobe
'comes with adobe reader in the reader\activex dir.
'the pdf.ocx is not supported by adobe, works with inet explorer and no doc avail.
'Micorsoft Internet Controls - shdocvw.dll - vb6

'use this declare statement to play sounds
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'string variables used in program
Dim msg, DefaultDrive, DefaultPath As String

Private Sub PlaySound(strFileName As String)
'play wave file
    sndPlaySound strFileName, 1
End Sub

Private Sub lstFiles_Click()
'load file user has chosen
    If optPDF.Value = True Then

'show selected pdf file in the pdfViewer
        pdfViewer.LoadFile (DefaultPath & "\" & lstFiles.FileName)
        DoEvents        'give file a moment to load b4 sound
        PlaySound "PDF.wav"
''        pdfViewer.setZoom (100) 'set view area size
    
    ElseIf optHTML.Value = True Then
'show selected html file in the htmlBrowser
        htmlBrowser.Navigate DefaultPath & "\" & lstFiles.FileName
        DoEvents        'give file a moment to load b4 sound
        PlaySound "HTML.wav"
    
    ElseIf optTEXT.Value = True Then
        txtViewer.LoadFile DefaultPath & "\" & lstFiles.FileName
        DoEvents        'give file a moment to load b4 sound
        PlaySound "TEXT.wav"
    End If
End Sub

Private Sub Form_Load()
'proceed upon initial load of frmPDF
    PlaySound "Load.wav"

'default file list to pdf's
    optPDF.Value = True

'set var. to current drive
    DefaultDrive = drvTarget.Drive

'set var. to current directory path
    DefaultPath = dirCurrent.Path

'call viewer resize function
    Call Form_Resize

'get current file type file count for current directory
    Call GetNewCount

'set htmlBrowser not visible
    htmlBrowser.Visible = False

'set txtViewer not visible
    txtViewer.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'exit program module
    Cancel = True
    
    msg = MsgBox("Are You Sure About Exiting The Program!", vbYesNo + vbQuestion, "PDF / HTML eBook Viewer")
    
    If msg = vbYes Then
        PlaySound "ExitProg.wav"
        End 'see ya
    
    Else
        frmPDF.Show 'decided to read some more, eh
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
'resize pdf viewing area upon form resize
    pdfViewer.Width = ScaleWidth
'subtract about an inch from height so bottom of control is seen
'this is from not having the control @ top 0
    pdfViewer.Height = ScaleHeight - 1115

'resize html viewing area upon form resize
    txtViewer.Width = ScaleWidth
'subtract about an inch from height so bottom of control is seen
'this is from not having the control @ top 0
    txtViewer.Height = ScaleHeight - 1115
    
'resize html viewing area upon form resize
    htmlBrowser.Width = ScaleWidth
'subtract about an inch from height so bottom of scroll control is seen
'this is from not having the control @ top 0
    htmlBrowser.Height = ScaleHeight - 1115
End Sub

Private Sub drvTarget_Change()
'error on - drive is empty
On Error GoTo ErrHandler

'set directory list to current drive
    dirCurrent.Path = drvTarget.Drive

'set var. to current drive
    DefaultDrive = drvTarget.Drive

'set var. to current path
    DefaultPath = dirCurrent.Path

    Exit Sub

'if drive is empty, mostly for floppies, zips, etc., then show message and reset drvTarget to default
ErrHandler:
    MsgBox "Doh, Put a Disk in Drive " & UCase(drvTarget.Drive), vbOKOnly, "Error Reading This Drive"
    drvTarget.Drive = DefaultDrive
    dirCurrent.Path = DefaultPath
End Sub

Private Sub dirCurrent_Change()
'set dir. file list to current dir under drive selected
    lstFiles.Path = dirCurrent.Path

'set var. to current path
    DefaultPath = dirCurrent.Path

'get current file type file count for current directory
    Call GetNewCount
End Sub

Private Sub optHTML_Click()
'play html open sound
    PlaySound "HTML.wav"
    optTEXT.ForeColor = vbBlack
    optPDF.ForeColor = vbBlack
    optHTML.ForeColor = vbRed
    
'set file pattern
    lstFiles.Pattern = "*.htm*"

'make html wiewing area visible
    htmlBrowser.Visible = True

'make pdf viewing area not visible
    pdfViewer.Visible = False

'make text viewing area not visible
    txtViewer.Visible = False

'get current file type file count for current directory
    Call GetNewCount
End Sub

Private Sub optPDF_Click()
'play pdf open sound
    PlaySound "PDF.wav"
    optTEXT.ForeColor = vbBlack
    optPDF.ForeColor = vbRed
    optHTML.ForeColor = vbBlack

'set file pattern
    lstFiles.Pattern = "*.pdf"

'make html wiewing area not visible
    htmlBrowser.Visible = False

'make pdf viewing area visible
    pdfViewer.Visible = True

'make text viewing area not visible
    txtViewer.Visible = False

'get current file type file count for current directory
    Call GetNewCount
End Sub

Private Sub optTEXT_Click()
'play text open sound
    PlaySound "TEXT.wav"
    optTEXT.ForeColor = vbRed
    optPDF.ForeColor = vbBlack
    optHTML.ForeColor = vbBlack
    
'set file pattern
    lstFiles.Pattern = "*.rtf;*.txt"

'make html wiewing area not visible
    htmlBrowser.Visible = False

'make pdf viewing area not visible
    pdfViewer.Visible = False

'make text viewing area visible
    txtViewer.Visible = True
    
'get current file type file count for current directory
    Call GetNewCount
End Sub

Private Sub GetNewCount()
'reset book count for users current directory and file type
    If optPDF.Value = True Then
        lblBookCount.Caption = "PDF eBooks in list : " & lstFiles.ListCount
    
    ElseIf optHTML.Value = True Then
        lblBookCount.Caption = "HTML eBooks in list : " & lstFiles.ListCount
    
    ElseIf optTEXT.Value = True Then
        lblBookCount.Caption = "TEXT/RTF eBooks in list : " & lstFiles.ListCount
    End If
End Sub
