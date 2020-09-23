VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "IMDB Tool"
   ClientHeight    =   7680
   ClientLeft      =   5145
   ClientTop       =   2910
   ClientWidth     =   8235
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   Begin MSComctlLib.Toolbar wnds 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgs"
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4440
      Top             =   2820
   End
   Begin MSComctlLib.ImageList imgs 
      Left            =   1800
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1394
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock imdbHTML 
      Left            =   360
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7425
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgs"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Get movie info with IMDB code"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search for movies"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMovie 
      Caption         =   "&Movie"
      Begin VB.Menu mnuBID 
         Caption         =   "&Get Movie info By ID..."
      End
      Begin VB.Menu mnuSrch 
         Caption         =   "&Search Database..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&View"
      Begin VB.Menu mnuVSC 
         Caption         =   "Stretch Cover"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HTML_SOURCE As String
Public Function GetPageSource(sPage As String) As String
'connect to IMDB Server

Dim HEADERS As String

AddLogEvent "Attempting to connect to IMDB"
If imdbHTML.State <> sckClosed Then imdbHTML.Close

imdbHTML.Connect IMDB_HostName, 80

'WAIT FOR A CONNECTION XXXXXXXXXXXXXXXXX
Do
    Select Case imdbHTML.State
        Case sckConnecting, sckConnected, sckResolvingHost, sckHostResolved, sckConnectionPending
        Case Else
            If imdbHTML.State <> sckClosed Then imdbHTML.Close
            AddLogEvent "Could not connect to IMDB!"
            Exit Function
            'ERROR
    End Select
DoEvents
Loop While imdbHTML.State <> sckConnected
'WAIT FOR A CONNECTION XXXXXXXXXXXXXXXXX

'WE ARE CONNECTED XXXXXXXXXXXXXXXXXXXXXX
HTML_SOURCE = ""
AddLogEvent "Connected to IMDB"

HEADERS = "GET " & sPage & " HTTP/1.1" & vbCrLf & _
    "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & _
    "Accept -Language: en -us" & vbCrLf & _
    "Accept -Encoding: gzip , deflate" & vbCrLf & _
    "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.1.4322)" & vbCrLf & _
    "Host: " & imdbHTML.RemoteHost & vbCrLf & _
    "Content-Length: 1" & vbCrLf & _
    "Connection: Close" & vbCrLf & vbCrLf


'SEND HEADERS XXXXXXXXXXXXXXXXXXXXXXXXXX
If imdbHTML.State = sckConnected Then
    imdbHTML.SendData HEADERS
    DoEvents
Else
    AddLogEvent "Unknown Error, Abort"
    Exit Function
End If

Do
'Getting data from server
DoEvents
Loop While imdbHTML.State = sckConnected
GetPageSource = HTML_SOURCE

AddLogEvent "Operation Completed"

End Function

Private Sub imdbHTML_DataArrival(ByVal bytesTotal As Long)
Dim TEMP_DATA As String
imdbHTML.GetData TEMP_DATA

DoEvents
HTML_SOURCE = HTML_SOURCE & TEMP_DATA
DoEvents
sb1.Panels(1).Text = "Getting HTML Source: " & Len(HTML_SOURCE) & " Bytes "
End Sub


Private Sub MDIForm_Load()
Init

End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo dontcont:

'When a file is dragged onto the form, find IMDB codes in it and
'automatically download movie information

Dim FILENAME As String
Dim IMDBCODE1 As String
Dim HTMLBUF As String

FILENAME = Data.Files(1)

If Trim(FILENAME) = "" Then
    Exit Sub
End If

'deal with files less than 2 megs
If FileLen(FILENAME) > (2000000) Then
    AddLogEvent "Dragged File to Big!"
    Exit Sub
End If

Dim FILEBUF As String
FILEBUF = Space$(FileLen(FILENAME))
Open FILENAME For Binary Access Read As #1
    Get #1, , FILEBUF
Close #1

Dim FBSPL() As String
FBSPL = Split(FILEBUF, "tt", , vbTextCompare)
If UBound(FBSPL) > 0 Then
    For i = 1 To UBound(FBSPL)
        Select Case Left(FBSPL(i), 1)
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
                'valid imdb code
                IMDBCODE1 = "tt" & Left(FBSPL(i), 7)
                AddLogEvent "Found IMDB Code..."
                    AddQueueEntry IMDBCODE1
                    'HTMLBUF = frmMain.GetPageSource("/title/" & IMDBCODE1 & "/")
                    'ParseTitlePage HTMLBUF

                Exit Sub
        End Select
    Next
Else
    AddLogEvent "No IMDB code in file!"
    Exit Sub
End If

Exit Sub
dontcont:
AddLogEvent "File Error!"
End Sub

Private Sub mnuAbout_Click()
MsgBox "IMDB Data tool" & vbCrLf & vbCrLf & "Programmed by TecCRC" & vbCrLf & "Icons by Stock Sources" & vbCrLf & vbCrLf & "A tool to download movie information" & vbCrLf & _
                                                                                                                           "from the Internet Movie Database and" & vbCrLf & _
                                                                                                                           "display it in this nice little interface", vbInformation, "About IMDB Tool"
End Sub

Private Sub mnuBID_Click()
frmInputCode.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnumdd_Click()
For i = 133093 To 134093
    AddQueueEntry "tt" & "0" & i
Next
End Sub

Private Sub mnuVSC_Click()
mnuVSC.Checked = Not (mnuVSC.Checked)

End Sub

Private Sub Timer1_Timer()
sb1.Panels(3).Text = GetQueues() & " Queued"
ExecuteQueue
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Movie info by Code
        frmInputCode.Show
    Case 2 'Search
        frmSearch.Show
End Select
End Sub

Private Sub wnds_ButtonClick(ByVal Button As MSComctlLib.Button)

        WNDWS(Button.Tag).pFrm.SetFocus
        

End Sub
