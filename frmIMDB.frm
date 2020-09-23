VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIMDB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   ClientHeight    =   4230
   ClientLeft      =   3105
   ClientTop       =   4065
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIMDB.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   Begin VB.PictureBox tmpPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   2520
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSWinsockLib.Winsock sckPicture 
      Left            =   840
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView setTree 
      Height          =   1935
      Left            =   1980
      TabIndex        =   2
      Top             =   300
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Field"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   38100
      EndProperty
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1800
      ScaleHeight     =   315
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   60
      Width           =   3315
   End
   Begin VB.PictureBox picCover 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin MSComctlLib.ProgressBar picProgress 
         Height          =   315
         Left            =   300
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblPic 
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Label lblPlot 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   435
   End
End
Attribute VB_Name = "frmIMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private thisIMDBData As MOVIE_DATA_IMDB
Private PIC_DATA As String
Public FP As Boolean

Public PICSIZE As Long


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'find the linked object and
'clear info
For i = 0 To UBound(WNDWS)
    With WNDWS(i)
        If .pHwnd = Me.hwnd Then
            'this the form
            For ii = 1 To frmMain.wnds.Buttons.Count
                If frmMain.wnds.Buttons(ii).Tag = i Then
                    'remove the associated tab
                    frmMain.wnds.Buttons.Remove (ii)
                    Exit Sub
                End If
            Next
        End If
    End With
Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
picCover.Move 2, 2, (Me.ScaleWidth / 2) - 4, Me.ScaleHeight - 4
picTitle.Move picCover.Left + picCover.Width + 2, 2, Me.ScaleWidth - (picCover.Width + 6)
setTree.Move picTitle.Left, picTitle.Top + picTitle.Height + 2, picTitle.Width
setTree.ColumnHeaders(1).Width = ((setTree.Width / 2) - 3)
setTree.ColumnHeaders(2).Width = ((setTree.Width / 2) - 3)
lblPlot.Move picTitle.Left, setTree.Top + setTree.Height + 2, picTitle.Width, Me.ScaleHeight - (picTitle.Height + setTree.Height + 8)
picProgress.Move 24, (picCover.Height / 2) - (picProgress.Height / 2), picCover.Width - 48
lblPic.Move picProgress.Left, picProgress.Top + picProgress.Height, picProgress.Width
DRAW_IMAGE
Set_Title
End Sub

Public Sub Set_Title()
picTitle.Cls
picTitle.CurrentX = (picTitle.ScaleWidth / 2) - (picTitle.TextWidth(thisIMDBData.mTitle) / 2)
picTitle.CurrentY = (picTitle.ScaleHeight / 2) - (picTitle.TextHeight(thisIMDBData.mTitle) / 2)

picTitle.Print thisIMDBData.mTitle

End Sub

Public Sub LoadIMDBData(sF1 As String, sF2 As String, sF3 As String, sF4 As String, sF5 As String, sF6 As String, sF7 As String, sF8 As String, sF9 As String)
With thisIMDBData
    .Country = sF1
    .CoverURL = sF2
    .Language = sF3
    .mDate = sF4
    .mGenre = sF5
    .mSypnosys = sF6
    .mTitle = sF7
    .Runtime = sF8
    .userRating = sF9
    
End With

setTree.ListItems.Clear

Dim aa As ListItem
With thisIMDBData
    Set aa = setTree.ListItems.Add(, , "Date :")
        aa.SubItems(1) = .mDate
    Set aa = setTree.ListItems.Add(, , "Genre:")
        aa.SubItems(1) = .mGenre
    Set aa = setTree.ListItems.Add(, , "Runtime:")
        aa.SubItems(1) = .Runtime
    Set aa = setTree.ListItems.Add(, , "Language:")
        aa.SubItems(1) = .Language
    Set aa = setTree.ListItems.Add(, , "Country:")
        aa.SubItems(1) = .Country
    Set aa = setTree.ListItems.Add(, , "User Rating:")
        aa.SubItems(1) = .userRating
lblPlot.Caption = FixPlot(.mSypnosys)
Me.Caption = .mTitle

If Trim(.CoverURL) <> "" Then
    Me.Show
    Form_Resize
    DownloadImage .CoverURL, sckPicture
    'DownloadImage "http://www.get-right.com/getright52beta4.exe", sckPicture
End If

End With

Form_Resize
End Sub



Private Sub sckPicture_DataArrival(ByVal bytesTotal As Long)

Dim SS() As String
Dim NEWDATA As String
Dim HSPL() As String
Dim HYSPL() As String

sckPicture.GetData NEWDATA

SS = Split(NEWDATA, vbCrLf & vbCrLf)
If Not (FP) Then
If UBound(SS) > 0 Then


    'get image size
    HSPL = Split(SS(0), "content-length: ", , vbTextCompare)
    If UBound(HSPL) > 0 Then
        HYSPL = Split(HSPL(1), vbCrLf, , vbTextCompare)
        If UBound(HYSPL) > 0 Then
            PICSIZE = Val(HYSPL(0))
            picProgress.Max = PICSIZE
            picProgress.Visible = True
            
            
        Else
            GoTo nopicsize:
        End If
    Else
        GoTo nopicsize:
    End If
    
    
    
nopicsize:

    NEWDATA = ""
    For i = 1 To UBound(SS)
        If i <> UBound(SS) Then
            NEWDATA = NEWDATA & SS(i) & vbCrLf & vbCrLf
        Else
            NEWDATA = NEWDATA & SS(i)
        End If
    Next
    FP = True
End If
End If
DoEvents
PIC_DATA = PIC_DATA & NEWDATA
If picProgress.Visible Then
    On Error GoTo progerr:
    picProgress.Value = Len(PIC_DATA)
    lblPic.Caption = "Getting Picture: " & Round(Len(PIC_DATA) / 1024, 2) & " of " & Round(PICSIZE / 1024, 2) & " KB"
Else
    
    lblPic.Caption = "Getting Picture: " & Round(Len(PIC_DATA) / 1024, 2) & " KB"
    
End If

progerr:




End Sub

Public Sub GotPicture(Optional sOutFile As String)
'MsgBox Len(PIC_DATA)
Dim OutFile1 As String
OutFile1 = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & sOutFile

'remove trailing VBCRLF

'PIC_DATA = Left(PIC_DATA, Len(PIC_DATA) - 2)


Open OutFile1 For Output As #1
    Print #1, PIC_DATA
Close #1

picProgress.Visible = False
gotheader = False
lblPic.Visible = False

tmpPic.Picture = LoadPicture(OutFile1)
DRAW_IMAGE

PICSIZE = 0
End Sub

Public Sub StartImage()
PIC_DATA = ""
FP = False
picProgress.Visible = False

PICSIZE = 0
DoEvents
lblPic.Visible = True
picProgress.Value = 0
End Sub

Public Sub DRAW_IMAGE()

Dim ORIGW As Long
Dim ORIGH As Long

Dim SCLW As Long
Dim SCLH As Long

Dim IMGW As Long
Dim IMGH As Long
Dim XX As Long
Dim YY As Long

Dim ASPW As Long
Dim ASPH As Long


ORIGW = tmpPic.ScaleWidth
ORIGH = tmpPic.ScaleHeight
SCLW = picCover.ScaleWidth
SCLH = picCover.ScaleHeight

XX = 0
YY = 0
    
IMGW = SCLW
IMGH = SCLH

    
If ORIGW < SCLW And ORIGH < SCLH Then
    'picture is small enough to fit into the
    'picture box!
    If frmMain.mnuVSC.Checked = False Then
        XX = (SCLW / 2) - (ORIGW / 2)
        YY = (SCLH / 2) - (ORIGH / 2)
        IMGW = ORIGW
        IMGH = ORIGH
    Else
        XX = 0
        YY = 0
        IMGW = SCLW
        IMGH = SCLH
    End If
    GoTo DRAWIT:
End If

If ORIGW > SCLW And ORIGH > SCLH Then
    'image is too large to fit into the picture
    'box, shrink it with aspect ratio
    XX = 0
    YY = 0
    
    IMGW = SCLW
    IMGH = SCLH
    
    
    GoTo DRAWIT:
End If

If ORIGW > SCLW Then
    'only the width is too large to fit
    'resize for aspect ratio
End If



DRAWIT:
picCover.Cls
SetStretchBltMode picCover.hdc, STRETCH_HALFTONE
StretchBlt picCover.hdc, XX, YY, IMGW, IMGH, tmpPic.hdc, 0, 0, ORIGW, ORIGH, vbSrcCopy

End Sub
