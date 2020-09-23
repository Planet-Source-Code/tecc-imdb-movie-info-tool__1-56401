VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search Movies"
   ClientHeight    =   6240
   ClientLeft      =   3090
   ClientTop       =   2250
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
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   Begin MSComctlLib.ImageList imgs 
      Left            =   2460
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":035E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   1005
      ButtonWidth     =   1799
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgs"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Get Checked"
            Object.ToolTipText     =   "Download Selected Movie Information"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   4260
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "20"
      ToolTipText     =   "Maximum Movies to Display"
      Top             =   300
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   315
      Left            =   5340
      TabIndex        =   4
      Top             =   300
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Enter movie name or terms to search for"
      Top             =   300
      Width           =   4095
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   5970
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstResults 
      Height          =   4575
      Left            =   60
      TabIndex        =   0
      Top             =   1320
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgs"
      SmallIcons      =   "imgs"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Movie Title"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IMDB ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckSearch 
      Left            =   300
      Top             =   540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   8
      Top             =   570
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Max Results:"
      Height          =   255
      Left            =   4260
      TabIndex        =   5
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Search Text:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   5175
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HTML_RES As String


Private Sub Command1_Click()
Searchit
End Sub

Public Sub Searchit()

Dim HEADER As String
Dim SearchQuery As String

If sckSearch.State <> sckClosed Then sckSearch.Close


sckSearch.Connect "www.imdb.com", 80
disablecommands False
Do
DoEvents
Select Case sckSearch.State
    Case sckConnected, sckConnecting, sckResolvingHost, sckHostResolved, sckConnectionPending
    Case Else
        If sckSearch.State <> sckClosed Then sckSearch.Close
        MsgBox "Error while connecting to Internet movie database!", vbCritical, "Connection Error"
        disablecommands True
        
        Exit Sub
End Select
Loop Until sckSearch.State = sckConnected

'connected, request search
'find?tt=on;mx=20;q=last

'form the search query, and replace any
'spaces in the query with %20's, which
'are code for HTML spaces

SearchQuery = "/find?tt=on;mx=" & Text2.Text & ";q=" & formatHTML(Text1.Text)


'q      =    Query
'tt     =    IMDB CODE!! MUST BE ON
'mx     =    Max Results
HEADER = "GET " & SearchQuery & " HTTP/1.1" & vbCrLf & _
    "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & _
    "Accept -Language: en -us" & vbCrLf & _
    "Accept -Encoding: gzip , deflate" & vbCrLf & _
    "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.1.4322)" & vbCrLf & _
    "Host: " & sckSearch.RemoteHost & vbCrLf & _
    "Content-Length: 0" & vbCrLf & _
    "Connection: Close" & vbCrLf & vbCrLf

If sckSearch.State = sckConnected Then
    DoEvents
    HTML_RES = ""
    lstResults.ListItems.Clear
    DoEvents
    sckSearch.SendData HEADER
    DoEvents
Else
    disablecommands True
    Exit Sub
End If

Do
DoEvents
Loop While sckSearch.State = sckConnected
disablecommands True

'completed
parseSearchHTML HTML_RES
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text1.Top = tb.Height + Label1.Height
Label1.Move 4, Text1.Top - Label1.Height
lstResults.Move 2, Text1.Height + Text1.Top + 4, Me.ScaleWidth - 4, Me.ScaleHeight - (Text1.Height + Text1.Top + sb.Height + 6)
Command1.Move Me.ScaleWidth - (Command1.Width + 6), Text1.Top
Text2.Move Command1.Left - (Text2.Width + 6), Text1.Top
Label2.Move Text2.Left, Label1.Top

Text1.Move 4, Text1.Top, Me.ScaleWidth - (Text2.Width + Command1.Width + 22)
lstResults.ColumnHeaders(1).Width = lstResults.Width / 2

End Sub

Private Sub disablecommands(ef As Boolean)
Command1.Enabled = ef
Text1.Enabled = ef
Text2.Enabled = ef
lstResults.Enabled = ef
Label1.Enabled = False
Label2.Enabled = False
lstResults.MousePointer = IIf(ef, ccArrow, ccArrowHourglass)
End Sub

Private Sub lstResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If lstResults.SortOrder = lvwDescending Then
    lstResults.SortOrder = lvwAscending
Else
    lstResults.SortOrder = lvwDescending
End If
lstResults.SortKey = ColumnHeader.Index - 1
lstResults.Sorted = True

End Sub

Private Sub lstResults_DblClick()
On Error GoTo noe:
Dim CDE As String
CDE = lstResults.SelectedItem.SubItems(2)
If Left(Trim(CDE), 2) = "tt" Then
    'valid code
    AddQueueEntry CDE
    DoEvents
    lstResults.SelectedItem.ForeColor = RGB(150, 150, 150)
End If
noe:
End Sub

Private Sub sckSearch_DataArrival(ByVal bytesTotal As Long)
Dim INHTML As String
sckSearch.GetData INHTML

HTML_RES = HTML_RES & INHTML
End Sub

Public Sub parseSearchHTML(inStrng As String)
'like
'<a href="/title/tt0156729/">Last Night (1998/I)</a>
Dim S() As String
Dim S1() As String
Dim S2() As String
Dim S3() As String
Dim S4() As String

Dim LI As ListItem
Dim numRes As Integer


Dim IMDB_CODE As String
Dim IMDB_TITLE As String
Dim IMDB_DATE As String

S = Split(inStrng, "<a href=" & """" & "/title/tt", , vbTextCompare)
If UBound(S) > 0 Then
    For i = 1 To UBound(S)
        'cycle through title links
        S1 = Split(S(i), "</a>", , vbTextCompare)
        If UBound(S1) > 0 Then
            'valid links
            S2 = Split(S1(0), "/" & """" & ">", , vbTextCompare)
            If UBound(S2) > 0 Then
                'valid title link
                IMDB_CODE = "tt" & S2(0)
                IMDB_TITLE = FixHTMLChars(S2(1))
                S3 = Split(IMDB_TITLE, "(", , vbTextCompare)
                If UBound(S3) > 0 Then
                    'date exists, parse it
                    S4 = Split(S3(1), ")")
                    If UBound(S4) > 0 Then
                        'full date
                        IMDB_TITLE = Trim(S3(0))
                        IMDB_DATE = S4(0)
                        
                        Set LI = lstResults.ListItems.Add(, , IMDB_TITLE, , 1)
                            LI.SubItems(1) = IMDB_DATE
                            LI.SubItems(2) = IMDB_CODE
                        
                        
                    Else
                        Set LI = lstResults.ListItems.Add(, , IMDB_TITLE, , 1)
                            'LI.SubItems(1) = IMDB_DATE
                            LI.SubItems(2) = IMDB_CODE
                    End If
                Else
                    Set LI = lstResults.ListItems.Add(, , IMDB_TITLE, , 1)
                        'LI.SubItems(1) = IMDB_DATE
                        LI.SubItems(2) = IMDB_CODE
                End If
            End If
            
            
        End If
        DoEvents
    Next
    sb.Panels(1).Text = lstResults.ListItems.Count & " Results. "
    
Else
    lstResults.ListItems.Add , , "No Results!"
End If


End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
    For i = 1 To lstResults.ListItems.Count
        If lstResults.ListItems(i).Checked = True Then
            AddQueueEntry lstResults.ListItems(i).SubItems(2)
            DoEvents
        End If
    Next
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Searchit
    KeyCode = -1
    Exit Sub
End If
End Sub
