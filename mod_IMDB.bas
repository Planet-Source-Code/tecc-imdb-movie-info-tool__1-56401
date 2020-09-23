Attribute VB_Name = "mod_IMDB"
Public queueexecute As Boolean


Public Type MOVIE_DATA_IMDB
    mTitle As String
    mSypnosys As String
    mGenre As String
    userRating As String
    mDate As String
    Runtime As String
    Country As String
    Language As String
    CoverURL As String
End Type

Public CurrentIMDBdata As MOVIE_DATA_IMDB

Public Const IMDB_SearchTitlePage As String = "IMDb name and title search"

Public Const IMDB_SEP_Title_LEFT1 As String = "<title>"
Public Const IMDB_SEP_Title_RIGHT1 As String = "</title>"

'IMDB Code Variables

Public Const SPL_GENRE_LEFT1 As String = "Genre:</b>" & vbLf & "<"
Public Const SPL_GENRE_RIGHT1 As String = "</a>"
Public Const SPL_GENRE_LEFT2 As String = ">"

Public Const SPL_PLOTOUTLINE_LEFT1 As String = "Plot Outline:</b>"
Public Const SPL_PLOTOUTLINE_LEFT2 As String = "Plot Summary:</b>"
Public Const SPL_PLOTOUTLINE_RIGHT1 As String = "<a href"

Public Const SPL_USERRATING_LEFT1 As String = "alt=" & """" & "_" & """" & "></a>" & vbLf & "<b>"
Public Const SPL_USERRATING_RIGHT1 As String = "</b>"

Public Const SPL_RUNTIME_LEFT1 As String = "Runtime:</b>" & vbLf
Public Const SPL_RUNTIME_RIGHT1 As String = vbLf

Public Const SPL_COUNTRY_LEFT1 As String = "Country:</b>" & vbLf
Public Const SPL_COUNTRY_RIGHT1 As String = "</a>" & vbLf
Public Const SPL_COUNTRY_LEFT2 As String = ">"

Public Const SPL_LANGUAGE_LEFT1 As String = "Language:</b>" & vbLf
Public Const SPL_LANGUAGE_RIGHT1 As String = "</a>" & vbLf
Public Const SPL_LANGUAGE_LEFT2 As String = ">"

'box image
Public Const SPL_COVER_LEFT1 As String = "alt=" & """" & "cover" & """"
Public Const SPL_COVER_RIGHT1 As String = "</a>" & vbLf
Public Const SPL_COVER_LEFT2 As String = "src=" & """"
Public Const SPL_COVER_RIGHT2 As String = """" & " "


Public Const IMDB_HostName As String = "www.imdb.com"

Public Type TransFerQueue
    sType As Byte
    lStatus As Byte
    IMDBCODE As String
    sData As String
    lTag As String
    mPicURL As String
    sSock As Winsock
    
End Type

Public Queue() As TransFerQueue


Public Sub AddLogEvent(sLog As String)
frmMain.sb1.Panels(2).Text = sLog

End Sub


Public Sub ParseTitlePage(indatas As String)
Dim SPL() As String
Dim SPL1() As String
Dim SPL2() As String
Dim SPL3() As String
Dim DSPL() As String
Dim DSPL1() As String

Dim MOVIE_TITLE As String



SPL = Split(indatas, "<title>", , vbTextCompare)
SPL1 = Split(indatas, "</title>", , vbTextCompare)

With CurrentIMDBdata

.Country = ""
.CoverURL = ""
.Language = ""
.mDate = ""
.mGenre = ""
.mSypnosys = ""
.mTitle = ""
.Runtime = ""
.userRating = ""

If UBound(SPL) > 0 Then
    If UBound(SPL1) > 0 Then
        'got page title
        SPL2 = Split(SPL(1), "</title>", , vbTextCompare)
        .mTitle = SPL2(0)
        'fix quoted movie titles
        SPL3 = Split(.mTitle, "&#34;", , vbTextCompare)
        If UBound(SPL3) > 0 Then
            'quotes exist
            .mTitle = ""
            For i = 0 To UBound(SPL3)
                If i <> UBound(SPL3) Then
                    .mTitle = .mTitle & SPL3(i) & """"
                Else
                    .mTitle = .mTitle & SPL3(i)
                End If
            Next

        Else

        End If
    End If
End If



'Seperate Date and Title
If .mTitle <> "" Then
'parse date
    DSPL = Split(.mTitle, "(", , vbTextCompare)
    If UBound(DSPL) > 0 Then
        'date in parenthases
        
        
        dspl2 = Split(DSPL(1), ")", , vbTextCompare)
        If UBound(dspl2) > 0 Then
            'date here
            .mDate = Trim(dspl2(0))
            .mTitle = Trim(DSPL(0))
            
        End If
    End If
End If


Erase SPL
Erase SPL2
Erase SPL1
Erase SPL3
SPL = Split(indatas, SPL_GENRE_LEFT1, , vbTextCompare)

If UBound(SPL) > 0 Then
    'found genre
    SPL1 = Split(SPL(1), SPL_GENRE_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'genre here
    SPL2 = Split(SPL1(0), SPL_GENRE_LEFT2, , vbTextCompare)
        If UBound(SPL2) > 0 Then
        .mGenre = SPL2(1)
        End If
    End If
End If
Erase SPL
Erase SPL2
Erase SPL1
Erase SPL3
SPL = Split(indatas, SPL_GENRE_LEFT1, , vbTextCompare)

If UBound(SPL) > 0 Then
    'found genre
    SPL1 = Split(SPL(1), SPL_GENRE_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'genre here
    SPL2 = Split(SPL1(0), SPL_GENRE_LEFT2, , vbTextCompare)
        If UBound(SPL2) > 0 Then
        .mGenre = SPL2(1)
        End If
    End If
End If


Erase SPL
Erase SPL2
Erase SPL1
Erase SPL3
SPL = Split(indatas, SPL_PLOTOUTLINE_LEFT1, , vbTextCompare)

If UBound(SPL) > 0 Then
    'method 2 "PLOT OUTLINE"
    'found plot outline
    SPL1 = Split(SPL(1), SPL_PLOTOUTLINE_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'Plot here
   
        .mSypnosys = SPL1(0)
        
    End If
Else
    'method 2 "PLOT SUMMARY"
    SPL = Split(LCase(indatas), SPL_PLOTOUTLINE_LEFT2, , vbTextCompare)
    If UBound(SPL) > 0 Then
    'found plot outline
    SPL1 = Split(SPL(1), SPL_PLOTOUTLINE_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'Plot here
   
        .mSypnosys = SPL1(0)
    End If
    End If
End If


'User Rating
Erase SPL
Erase SPL2
Erase SPL1
Erase SPL3
SPL = Split(indatas, SPL_USERRATING_LEFT1, , vbTextCompare)

If UBound(SPL) > 0 Then
    'found User Rating
    SPL1 = Split(SPL(1), SPL_USERRATING_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'User Rating Here
    
        .userRating = SPL1(0)
    End If
End If


'Runtime
Erase SPL
Erase SPL2
Erase SPL1
Erase SPL3
SPL = Split(indatas, SPL_RUNTIME_LEFT1, , vbTextCompare)

If UBound(SPL) > 0 Then
    'found runtime
    SPL1 = Split(SPL(1), SPL_RUNTIME_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'runtime here
    
        .Runtime = Trim(SPL1(0))
    End If
End If


'Country
Erase SPL
Erase SPL2
Erase SPL1
Erase SPL3
SPL = Split(indatas, SPL_COUNTRY_LEFT1, , vbTextCompare)
If UBound(SPL) > 0 Then
    'found Country
    SPL1 = Split(SPL(1), SPL_COUNTRY_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'Country here
        SPL2 = Split(SPL1(0), SPL_COUNTRY_LEFT2, , vbTextCompare)
        If UBound(SPL2) > 0 Then
            'found country
            .Country = Trim(SPL2(1))
            'fix weird error with country
            'formatting
            SPL3 = Split(.Country, "</a", , vbTextCompare)
            If UBound(SPL3) > 0 Then
                'error is there, replace it
                .Country = Trim(SPL3(0))
            End If
        End If
        
    End If
End If

'Language
Erase SPL
Erase SPL2
Erase SPL1
Erase SPL3
SPL = Split(indatas, SPL_LANGUAGE_LEFT1, , vbTextCompare)
If UBound(SPL) > 0 Then
    'found language
    SPL1 = Split(SPL(1), SPL_LANGUAGE_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'language here
        SPL2 = Split(SPL1(0), SPL_LANGUAGE_LEFT2, , vbTextCompare)
        If UBound(SPL2) > 0 Then
            'found language
            .Language = Trim(SPL2(1))
            'fix weird error with language
            'formatting
            SPL3 = Split(.Language, "</a", , vbTextCompare)
            If UBound(SPL3) > 0 Then
                'error is there, replace it
                .Language = Trim(SPL3(0))
            End If
        End If
        
    End If
End If

    'You see, IMDB puts an ALT code into the cover
    'picture, so we can easilly find it on the page!
    'BOX COVER ART/IMAGE ************************
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Erase SPL
Erase SPL2
Erase SPL1
Erase SPL3

'search for cover image
SPL = Split(indatas, SPL_COVER_LEFT1, , vbTextCompare)

If UBound(SPL) > 0 Then
    'found partial box image code, parse
    'further
    SPL1 = Split(SPL(1), SPL_COVER_RIGHT1, , vbTextCompare)
    If UBound(SPL1) > 0 Then
        'full box image code found, make sure
        'its valid... </a> tag is not too unique
        SPL2 = Split(SPL1(0), SPL_COVER_LEFT2, , vbTextCompare)
        If UBound(SPL2) > 0 Then
            'we are 90% sure this is the cover
            'box art, we will now just format
            'it to return just the image url...
            SPL3 = Split(SPL2(1), SPL_COVER_RIGHT2, , vbTextCompare)
            If UBound(SPL3) > 0 Then
                '100% url for our image!
                .CoverURL = Trim(SPL3(0))

                'MsgBox .CoverURL
            End If
        End If
    End If
End If

    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'BOX COVER ART/IMAGE ************************

End With

ShowMovieData CurrentIMDBdata
End Sub


Public Sub ParseSearchPage(sInTX As String)

End Sub

Public Sub ShowMovieData(sIMDBDAT As MOVIE_DATA_IMDB)
'Dim A As New frmIMDB
Dim BB As Button
Dim NWID As Long

'Load A
NWID = NewWindow()

With sIMDBDAT
    WNDWS(NWID).pFrm.LoadIMDBData .Country, .CoverURL, .Language, .mDate, .mGenre, .mSypnosys, .mTitle, .Runtime, .userRating
    Set BB = frmMain.wnds.Buttons.Add(, , Trim(Left(.mTitle, 16)), , 3)
        BB.Tag = NWID
        BB.ToolTipText = .mTitle

End With
WNDWS(NWID).pHwnd = WNDWS(NWID).pFrm.hwnd
WNDWS(NWID).pFrm.Show
End Sub

Public Sub AddQueueEntry(sIMDBCode As String, Optional ByVal lType As Byte = 0, Optional ByVal lTag As Long, Optional URL As String)
Dim nn As Integer
For i = 0 To UBound(Queue)
    If Queue(i).lStatus = 0 Then
        'free queue entry
        nn = i
        GoTo fnd:
    End If
Next
nn = UBound(Queue) + 1
ReDim Preserve Queue(nn)

fnd:
With Queue(nn)
    .lStatus = 1 'waiting
    .sType = lType
    .IMDBCODE = sIMDBCode
    .lTag = lTag
    .mPicURL = URL
    .sData = ""

End With
End Sub

Public Sub ExecuteQueue()
Dim HTMLBUF As String
If Not queueexecute Then
For i = 0 To UBound(Queue)
    If Queue(i).lStatus = 1 Then
        'queue waiting
        Queue(i).lStatus = 2 'transfering
        queueexecute = True
        DoEvents
        Select Case Queue(i).sType
            Case 0 'title page
                HTMLBUF = frmMain.GetPageSource("/title/" & Queue(i).IMDBCODE & "/")
                ParseTitlePage HTMLBUF
                DoEvents
                Queue(i).lStatus = 0 'completed
                queueexecute = False
                Exit Sub
            Case 1 'image download
                queueexecute = True
                frmMain.Timer1.Enabled = False
                
                DoEvents
                With Queue(i)
                    DownloadImage .mPicURL, .sSock
                    DoEvents
                    queueexecute = False
                    frmMain.Timer1.Enabled = True
                    .lStatus = 0
                    Exit Sub
                End With
        End Select
    End If
Next
End If
End Sub

Public Function GetQueues() As Integer
Dim GQ As Integer

For i = 0 To UBound(Queue)
    If Queue(i).lStatus <> 0 Then
        GQ = GQ + 1
        
    End If
Next
GetQueues = GQ
End Function


Public Sub Init()
ReDim Queue(0)
ReDim WNDWS(0)
End Sub

Public Function ParseURL(sURL As String, ByRef sHost As String, ByRef sPath As String, ByRef sFile As String) As Integer
'http://www.google.com/a/b/c/d/ing.jpg
Dim s_A As String
Dim s_B As String
Dim NEW_URL As String
Dim TMPHOST As String
Dim TMPPATH As String
Dim TMPFILE As String

Dim SEP() As String
Dim SEP1() As String
Dim SEP2() As String

s_A = "://"     'must be http://
s_B = "/"       'path sep

SEP = Split(sURL, s_A, , vbTextCompare)
If UBound(SEP) > 0 Then
    'must have http:// before host
    NEW_URL = SEP(1)
Else
    'www.a.ccom/a/b
    NEW_URL = sURL
End If

SEP1 = Split(Trim(NEW_URL), s_B)

If UBound(SEP1) > 0 Then
    'valid url
    TMPHOST = SEP1(0)
    TMPFILE = SEP1(UBound(SEP1))
    TMPPATH = "/"
    For i = 1 To UBound(SEP1)
        If i <> UBound(SEP1) Then
            TMPPATH = TMPPATH & SEP1(i) & "/"
        Else
            TMPPATH = TMPPATH & SEP1(i)
        End If
    Next
End If

sHost = TMPHOST
sPath = TMPPATH
sFile = TMPFILE
End Function

Public Sub DownloadImage(sURL As String, lSock As Winsock)
Dim MHOST As String
Dim Mpath As String
Dim MFile As String
Dim HEADER As String
If lSock.State <> sckClosed Then lSock.Close
lSock.Parent.StartImage
 
 ParseURL sURL, MHOST, Mpath, MFile
 
If MHOST = "" Then
    Exit Sub
End If

lSock.Connect MHOST, 80
Do
    Select Case lSock.State
        Case sckConnected, sckConnecting, sckConnectionPending, sckResolvingHost, sckHostResolved
        Case Else
            If lSock.State <> sckClosed Then lSock.Close
            AddLogEvent "Error while downloading picture"
            Exit Sub
    End Select
DoEvents
Loop Until lSock.State = sckConnected

'connected!

HEADER = "GET " & Mpath & " HTTP/1.1" & vbCrLf & _
    "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & _
    "Accept -Language: en -us" & vbCrLf & _
    "Accept -Encoding: gzip , deflate" & vbCrLf & _
    "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.1.4322)" & vbCrLf & _
    "Host: " & lSock.RemoteHost & vbCrLf & _
    "Content-Length: 0" & vbCrLf & _
    "Connection: Close" & vbCrLf & vbCrLf



DoEvents
lSock.SendData HEADER

Do
DoEvents
Loop While lSock.State = sckConnected

lSock.Parent.GotPicture MFile
'completed
End Sub

Public Function formatHTML(sInStr As String) As String
Dim SPCSPL() As String
Dim OUTSTRs As String
SPCSPL = Split(Trim(sInStr), " ", , vbTextCompare)

If UBound(SPCSPL) <= 0 Then
    'no spaces
    formatHTML = sInStr
    Exit Function
End If

For i = 0 To UBound(SPCSPL)
    If i <> UBound(SPCSPL) Then
        OUTSTRs = OUTSTRs & SPCSPL(i) & "%20"
    Else
        OUTSTRs = OUTSTRs & SPCSPL(i)
    End If
Next

formatHTML = OUTSTRs
End Function

Public Function FixHTMLChars(sInStr As String) As String
Dim S1() As String
Dim OSTR As String

S1 = Split(sInStr, "&#34;", , vbTextCompare)
If UBound(S1) > 0 Then
    'exists, quotes
    For i = 0 To UBound(S1)
        If i <> UBound(S1) Then
            OSTR = OSTR & S1(i) & """"
        Else
            OSTR = OSTR & S1(i)
        End If
    Next
Else
    OSTR = sInStr
End If

Erase S1
S1 = Split(OSTR, "&#228;", , vbTextCompare)
If UBound(S1) > 0 Then
    OSTR = ""
    'exists, quotes
    For i = 0 To UBound(S1)
        If i <> UBound(S1) Then
            OSTR = OSTR & S1(i) & "ä"
        Else
            OSTR = OSTR & S1(i)
        End If
    Next
Else
    'OSTR = sInStr
End If
    
FixHTMLChars = OSTR
End Function

Public Function FixPlot(sInPlot As String) As String
Dim S1() As String
Dim FF As String
S1 = Split(sInPlot, "<br>", , vbTextCompare)
If UBound(S1) > 0 Then
    FF = S1(0)
Else
    FF = sInPlot
End If
FixPlot = FF
End Function

