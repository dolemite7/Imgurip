VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Imgurip"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3975
      Begin VB.ComboBox cboLinks 
         Height          =   315
         Left            =   840
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cboFunction 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2880
         Width           =   2775
      End
      Begin VB.CheckBox chkRecent 
         Caption         =   "Rip"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtRecent 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Text            =   "5"
         Top             =   1440
         Width           =   375
      End
      Begin VB.ComboBox cboRedditUser 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkTimeStamps 
         Caption         =   "Timestamps"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.OptionButton optFunction 
         Caption         =   "Reddit - Subreddit"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton optFunction 
         Caption         =   "Reddit - User"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optFunction 
         Caption         =   "File (in.txt)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton optFunction 
         Caption         =   "Imgur - Single Album"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton optFunction 
         Caption         =   "Imgur - User Album Gallery"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chkDebug 
         Caption         =   "Debug"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   4440
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Text            =   "DeliriumTremens"
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "c:\temp"
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CheckBox chkPreview 
         Caption         =   "Preview"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CheckBox chkThumbs 
         Caption         =   "Thumbs"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   4440
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   6
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblLinks 
         Caption         =   "Links: "
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblRecent 
         Caption         =   "most recent albums"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblURL 
         Caption         =   "Code/Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblPath 
         Caption         =   "Save Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   975
      End
   End
   Begin VB.Frame fraProgress 
      Caption         =   "Download Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4200
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
      Begin VB.Label lblCountPage 
         Caption         =   "Page:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCountImg 
         Caption         =   "Image:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblCountDir 
         Caption         =   "Folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.TextBox txtStatus 
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4920
      Width           =   6615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'http://imgur.com/a/pD50I/embed
'1. get album list
'2. open first album
'3. get embed page
'4. create folder name from album list order and <title>
'5. download images, rename file with XXX- order prefix

'todo:
' - start at album X
' / check/skip/overwrite existing file - partial
' - picture captions/extended info
' - menubar
' - reddit user number albums asc/desc?, timestamp, comment link?
' - drop datafullnames into collection, stop when existing (looping)


'Changelog:
'1.0.7: fixed issue with album <title> on different line
'1.0.8: ?
'1.0.9: Fixed single image countdir/countimg display
'1.1.2: Added file capabilities, StripImgur function, changed form
'1.1.3: fixed file radio button, statistics, handling of periods (may need . to -)

Option Explicit

Private Enum BIFlagsConstants
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_RETURNFSANCESTORS = &H8
    BIF_EDITBOX = &H10
    BIF_VALIDATE = &H20
    BIF_NEWDIALOGSTYLE = &H40
    BIF_USENEWUI = (BIF_EDITBOX Or BIF_NEWDIALOGSTYLE)
    BIF_BROWSEINCLUDEURLS = &H80
    BIF_UAHINT = &H100
    BIF_NONEWFOLDERBUTTON = &H200
    BIF_NOTRANSLATETARGETS = &H400
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_SHAREABLE = &H8000&
    BIF_BROWSEFILEJUNCTIONS = &H10000
End Enum

Dim oFSO As New FileSystemObject
Dim bDebug As Boolean


Private Sub chkRecent_Click()
    If chkRecent Then
        txtRecent.Enabled = True
    Else
        txtRecent.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    
    Main
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If oFSO.FileExists(App.Path & "\download.htm") Then
        oFSO.DeleteFile App.Path & "\download.htm", True
    End If
    Set oFSO = Nothing
    StoreRegValues
    
End Sub

Public Sub Main()

    Dim x As Integer
    Dim sClipboard As String

    'more: http://www.angelfire.com/mi4/bvo/vb/vbconio.htm
    'Store command line arguments in this array
    Dim sArgs() As String
    
    Dim iLoop As Integer
    'Assuming that the arguments passed from command line will have space in between,
    'you can also use comma or otehr things...
    sArgs = Split(Command$, " ")
    For iLoop = 0 To UBound(sArgs)
        'this will print the command line arguments that are passed from the command line
        Debug.Print sArgs(iLoop)
    Next

    If UBound(sArgs) > -1 Then
        Select Case Trim$(sArgs(0))
        
            Case "s"
            
            Case "f"
            
            Case ""
                If chkDebug Then bDebug = True
                frmMain.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
                Me.Show 1
        End Select
    Else
        If chkDebug Then bDebug = True
        frmMain.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
        
        cboFunction.AddItem "Imgur - Single Album"
        cboFunction.AddItem "Imgur - User Album Gallery"
        cboFunction.AddItem "Reddit - Subreddit"
        cboFunction.AddItem "Reddit - User"
        cboFunction.AddItem "File (in.txt)"
        cboFunction.ListIndex = 0
        
        cboRedditUser.AddItem "Submitted"
        cboRedditUser.AddItem "Comments"
        cboRedditUser.AddItem "All"
        cboRedditUser.ListIndex = 0
        
        cboLinks.AddItem "All Links"
        cboLinks.AddItem "Imgur Links"
        cboLinks.AddItem "No Links"
        cboLinks.ListIndex = 0
        
        
        
        
        'clipboard check
        If Clipboard.GetFormat(vbCFText) Or Clipboard.GetFormat(vbCFRTF) Then
           If Clipboard.GetFormat(vbCFText) Then sClipboard = Clipboard.GetText(vbCFText)
           If Clipboard.GetFormat(vbCFRTF) Then sClipboard = Clipboard.GetText(vbCFRTF)
           If InStr(1, sClipboard, "imgur.com") > 0 And Len(sClipboard) < 200 Then txtCode.Text = sClipboard
        End If
        
        GetRegValues
        
        Me.Show 1
    End If



End Sub

Private Sub chkPreview_Click()
    If chkPreview Then
        Image1.Visible = True
    Else
        Image1.Visible = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    
On Error GoTo ErrorHandler
    
    Dim shlShell
    Dim shlFolder
    
    Set shlShell = New Shell32.Shell
    
    Set shlFolder = shlShell.BrowseForFolder(Me.hWnd, "Select a Folder", _
        BIF_RETURNONLYFSDIRS)
    
    If Not shlFolder Is Nothing Then
        txtPath = shlFolder.Items.Item.Path
    End If
        
    Exit Sub

ErrorHandler:
    UpdateStatus "cmdBrowse_Click Error " & Err.Number & ": " & Err.Description
        
End Sub

Private Sub cmdGo_Click()
    
'    On Error GoTo ErrorHandler

    Dim aList
    Dim x As Integer

    Dim oFile
    Dim sFile As String
    
    Dim aFile
    
    'validation
    If Not optFunction(2).Value Then
        If Trim$(txtCode.Text) = "" Then
            MsgBox "Invalid code, name or URL, please try again. ", vbApplicationModal + vbExclamation + vbOKOnly, App.EXEName
            txtCode.SetFocus
            Exit Sub
        End If
    End If
    
    If Not oFSO.FolderExists(txtPath.Text) Then
        MsgBox "Invalid Save Folder, please try again. ", vbApplicationModal + vbExclamation + vbOKOnly, App.EXEName
        txtPath.SetFocus
        Exit Sub
    End If
    
    If optFunction(0).Value Or optFunction(3).Value Or optFunction(4).Value And chkRecent.Value Then
        If Trim$(txtRecent) = "" Or Not IsNumeric(Trim$(txtRecent)) Then
            UpdateStatus "Error: Please enter a numeric value for Rip Most Recent."
            Exit Sub
        ElseIf Trim$(txtRecent) = "0" Or InStr(1, txtRecent, ".") > 0 Or InStr(1, txtRecent, "-") > 0 Then
            UpdateStatus "Error: Please enter a whole number for Rip Most Recent."
            Exit Sub
        Else
            txtRecent = Trim$(txtRecent)
        End If
    End If
    
    LockFormControls (True)
    
    StoreRegValues
        
    If InStr(1, txtCode.Text, "/i.imgur.com") > 0 Then
        'Clean URL: http://imgur.com/a/DcuU0#0
        txtCode.Text = StripImgur(txtCode.Text)

        optFunction(1).Value = True
        UpdateStatus "Downloading single image: " & Mid(txtCode.Text, InStrRev(txtCode.Text, "/") + 1)
        GetImage txtCode.Text, txtPath
        lblCountDir = "Folder: 1/1 (100%)"
        lblCountImg = "Image: 1/1 (100%)"
        UpdateStatus "Download complete."
    Else
        If optFunction(0).Value = True Then
            UpdateStatus "Getting single album: " & Mid(txtCode.Text, InStrRev(txtCode.Text, "/") + 1)
            txtCode.Text = StripImgur(txtCode.Text)
            GetAlbumList (txtCode.Text)
        ElseIf optFunction(1).Value = True Then
            txtCode.Text = StripImgur(txtCode.Text)
            If InStr(1, txtCode.Text, ",") > 0 Then
                aList = Split(txtCode.Text, ",")
                For x = 0 To UBound(aList)
                    lblCountDir = "Folder: " & x + 1 & "/" & UBound(aList) & " (" & Round((CInt(x + 1) / CInt(UBound(aList) - 1)) * 100, 1) & "%)"
                    GetAlbum (Trim(aList(x)))
                Next x
            Else
                lblCountDir = "Folder: 0/1 (0%)"
                GetAlbum (txtCode.Text)
                lblCountDir = "Folder: 1/1 (100%)"
            End If
        ElseIf optFunction(2).Value = True Then
            UpdateStatus "Getting input file: " & Mid(txtCode.Text, InStrRev(txtCode.Text, "/") + 1)
            Set oFile = oFSO.OpenTextFile(App.Path & "\in.txt", ForReading)
            sFile = oFile.ReadAll
            aFile = Split(sFile, vbCrLf)
            For x = 0 To UBound(aFile) - 1
                lblCountDir = "Folder: " & x + 1 & "/" & UBound(aFile) & " (" & Round((CInt(x + 1) / CInt(UBound(aFile) - 1)) * 100, 1) & "%)"
                If Trim$(aFile(x)) <> "" Then
                    If Left(aFile(x), 1) <> "#" Then
                        txtCode.Text = StripImgur(aFile(x))
                        If InStr(1, txtCode.Text, "/i.imgur.com") > 0 Then
                            UpdateStatus "Downloading single image: " & Mid(txtCode.Text, InStrRev(txtCode.Text, "/") + 1)
                            GetImage txtCode.Text, txtPath
                            lblCountImg = "Image: 1/1 (100%)"
                        Else
                            GetAlbum (txtCode.Text)
                        End If
                    End If
                End If
            Next x
            UpdateStatus "Download complete."
        ElseIf optFunction(3).Value = True Then
            GetRedditUser (txtCode.Text)
            UpdateStatus "Download complete."
        ElseIf optFunction(4).Value = True Then
            UpdateStatus "Getting subreddit: " & txtCode
            GetSubreddit (txtCode.Text)
            UpdateStatus "Download complete."
        End If
    End If
    
    LockFormControls (False)
    
'    Select Case PageType(txtCode)
'        Case 0
'
'        Case 1 'pD50I
'            GetAlbum (txtCode)
'        Case 2
'
'    End Select
    
    Exit Sub

ErrorHandler:
    UpdateStatus ("cmdGo Error " & Err.Number & ": " & Err.Description)
    LockFormControls (False)

End Sub
Private Function GetRedditUser(inputcode As String) As Boolean


    Dim sCode As String
    Dim sFile As String
    Dim sCodeImage As String
    Dim sTitle As String
    Dim sFileOut As String
    Dim sPath As String
    Dim iAlbumMax As Integer
    Dim iAlbum As Integer
    Dim aFile
    
    ReDim aLink(3, 1000) As String
    '0: reddit link
    '1: reddit comments
    '2: external link
    
    Dim x As Integer
    Dim sAlbumMax As String
    Dim iPos, iMax As Integer
    Dim sTitleAlbum As String
    Dim sUsername As String
    
    Dim sPostTypes As String
    
    Dim sAlbumNum, sAlbumDir As String
    Dim iPage As Integer
    Dim iCount As Integer
    
    Dim sDate, sTime As String
    Dim sDateTime, sDateTimeTemp As String
    Dim dDate, dTime, dDateTime As Date
    Dim sTimeZH, sTimeZM As String
    
    Dim cDataFullName As New Collection
    
    Call InitializeCollection(cDataFullName)
    
    sCode = Trim$(inputcode)
    
    If InStr(1, sCode, "reddit.com/") > 1 Then
        sCode = Mid(sCode, InStrRev(sCode, "/") + 1)
    End If
    
    'pages 1
    sUsername = sCode
    UpdateStatus "Getting reddit user: " & sUsername
    Select Case cboRedditUser.ListIndex
        Case 0
            sPostTypes = "submitted/"
        Case 1
            sPostTypes = "comments/"
        Case 2
            sPostTypes = ""
    End Select
    
    sCode = "http://www.reddit.com/user/" & sCode & "/" & sPostTypes
    sPath = txtPath & "\" & txtCode
    If Not oFSO.FolderExists(sPath) Then
        oFSO.CreateFolder (sPath)
    End If
    
    Dim bPageNext As Boolean
    Dim sDataFullName As String
    Dim sAlbum As String
            
    Dim bFoundLinkListing As Boolean
    Dim bFoundThingID As Boolean
    Dim sThingID As String
    Dim sThingIDLast As String
    Dim sThingIDCheck As String
    Dim sLinkExternal As String
    Dim sLinkReddit As String
    Dim sLinkComments As String
    Dim bFoundImgur As Boolean
    
    Dim bFoundImgurAlbum As Boolean
    Dim bFoundImgurSingle As Boolean
    
    Dim iLink As Integer
    
    'http://www.reddit.com/user/spif/submitted/?count=25&after=t3_21zkxv
    'http://www.reddit.com/user/DeliriumTremens/submitted/?count=25&after=t3_y6l4h
    'http://www.reddit.com/user/DeliriumTremens/submitted/?count=50&after=t3_eokjn
    'http://www.reddit.com/user/DeliriumTremens/submitted/?count=75&after=t3_c0qzj
    'http://www.reddit.com/user/DeliriumTremens/submitted/?count=100&after=t3_aejzk
    '<div class=" thing id-t3_y6l4h odd link " onclick="click_thing(this)" data-fullname="t3_y6l4h" data-ups="89" data-downs="15" >
    '<div class="nav-buttons"><span class="nextprev">view more:&#32;<a href="http://www.reddit.com/user/DeliriumTremens/submitted/?count=26&amp;before=t3_vwhcv" rel="nofollow prev" >&lsaquo; prev</a><span class="separator"></span><a href="http://www.reddit.com/user/DeliriumTremens/submitted/?count=50&amp;after=t3_eokjn" rel="nofollow next" >next &rsaquo;</a></span></div>
        
    'pages
    iPage = 0
    
    Do
        lblCountPage.Caption = "Page: " & iPage + 1
        bPageNext = False
        If iPage >= 1 Then
            sCode = sCode & "/" & sPostTypes & "?count=" & iPage * 25 & _
                "&after=t3_" & sDataFullName
        End If
        GetHTMLFile sCode
        sFile = OpenFile(App.Path & "\download.htm")
        aFile = Split(sFile, ">")
        
        Debug.Print UBound(aFile)
        
        bFoundLinkListing = False
        
        For x = 0 To UBound(aFile)
        
            If InStr(1, aFile(x), "the page you requested does not exist") > 0 Then
                UpdateStatus "User " & sUsername & " was not found."
                sPath = txtPath
                lblCountPage.Caption = "Page: "
                Exit Function
            End If
        
            If InStr(1, aFile(x), "sitetable linklisting") > 0 Then
                bFoundLinkListing = True
            End If
            
            If bFoundLinkListing Then
            
                If InStr(1, aFile(x), """ thing id-") > 0 Then
                    bFoundThingID = True
                    sThingIDCheck = Mid(aFile(x), InStr(1, aFile(x), "data-fullname") + 18, 6)
                    'sDataFullName = Mid(aFile(x), InStr(1, aFile(x), "data-fullname") + 18, 5)
                    
                    If sThingIDCheck <> sThingIDLast Then
                        'clear vars
                        sThingID = sThingIDCheck
                        sThingIDLast = sThingID
                        sLinkExternal = ""
                        sLinkReddit = ""
                        sLinkComments = ""
                        sDateTime = ""
                        sDateTimeTemp = ""
                        sDate = ""
                        sTime = ""
                        sTimeZH = ""
                        sTimeZM = ""
Dim bFoundImgurSingleI As Boolean

                        bFoundImgur = False
                        bFoundImgurSingle = False
                        bFoundImgurSingleI = False
                        bFoundImgurAlbum = False
                        
                        sAlbum = ""
                    End If
                End If
                
                '<div class=" thing id-t3_23l2sc odd link self" onclick="click_thing(this)" data-fullname="t3_23l2sc" data-ups="11" data-downs="4" >
                
                '<span class="rank">1</span>
                '><a class="thumbnail self may-blank loggedin" href="/r/StLouisCirclejerk/comments/23l2sc/i_stole_this_dog_do_the_twelve_of_you_know_who/" >
                '<a class="title may-blank loggedin" href="/r/StLouisCirclejerk/comments/23l2sc/i_stole_this_dog_do_the_twelve_of_you_know_who/" tabindex="1" >I stole this dog, do the twelve of you know who owns it?</a>
                '<span class="domain">(<a href="/r/StLouisCirclejerk/">self.StLouisCirclejerk</a>)</span>
                '<p class="tagline">submitted&#32;<time title="Mon Apr 21 12:29:45 2014 UTC" datetime="2014-04-21T12:29:45+00:00" class="live-timestamp">25 days ago</time>&#32;by&#32;<a href="http://www.reddit.com/user/DeliriumTremens" class="author may-blank id-t2_3ncqw" >DeliriumTremens</a><span class="userattrs"></span>&#32;to&#32;<a href="http://www.reddit.com/r/StLouisCirclejerk/" class="subreddit hover may-blank" >/r/StLouisCirclejerk</a></p><ul class="flat-list buttons"><li class="first"><a href="http://www.reddit.com/r/StLouisCirclejerk/comments/23l2sc/i_stole_this_dog_do_the_twelve_of_you_know_who/" class="comments may-blank" >2 comments</a>
                
                If InStr(1, aFile(x), "data-fullname") > 0 Then
                    sDataFullName = Mid(aFile(x), InStr(1, aFile(x), "data-fullname") + 18, 6)
                    cDataFullName.Add sDataFullName
                    If bDebug Then UpdateStatus "DataFullName: " & sDataFullName
                End If
                
                If InStr(1, aFile(x), "/comments/") > 0 Then
                    sLinkComments = Trim(Mid(aFile(x), InStr(1, aFile(x), "href=""") + 6))
                    sLinkComments = Left(sLinkComments, InStr(1, sLinkComments, """") - 1)
                    If bDebug Then UpdateStatus "LinkComments: " & sLinkComments
                    sLinkReddit = Mid(sLinkComments, InStr(1, sLinkComments, "/comments/") + 10, 6)
                    sLinkReddit = "http://redd.it/" & sLinkReddit & "/"
                    If bDebug Then UpdateStatus "LinkReddit: " & sLinkReddit
                End If
                
                
                If sThingID <> "" And sDateTime = "" And InStr(1, aFile(x), "time title") > 0 Then
                    sDateTimeTemp = Mid(aFile(x), InStr(1, aFile(x), "datetime=") + 10, 25)
                    sDate = Replace(Left(sDateTimeTemp, 10), "-", "")
                    dDate = DateSerial(CInt(Left(sDate, 4)), CInt(Mid(sDate, 5, 2)), CInt(Mid(sDate, 7, 2)))
                    sTime = Mid(sDateTimeTemp, 12, 8)

                    dTime = TimeSerial(CInt(Left(sTime, 2)), CInt(Mid(sTime, 4, 2)), CInt(Mid(sTime, 7, 2)))
                    'sTimeZH = Mid(sDateTimeTemp, 20, 3)
                    'sTimeZM = Mid(sDateTimeTemp, 20, 1) & Mid(sDateTimeTemp, 24, 2)
                    dDateTime = dDate & " " & dTime
                    UpdateStatus "Datetime: " & CStr(dDateTime)
                    'dDateTime = DateAdd("h", CInt(sTimeZH), dDateTime)
                    'If CInt(sTimeZM) > 0 Then dDateTime = DateAdd("n", CInt(sTimeZM), dDateTime)
                    sDateTime = CStr(Year(dDateTime))
                    If Month(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                    sDateTime = sDateTime & CStr(Month(dDateTime))
                    If Day(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                    sDateTime = sDateTime & CStr(Day(dDateTime))
                    If Hour(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                    sDateTime = sDateTime & CStr(Hour(dDateTime))
                    If Minute(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                    sDateTime = sDateTime & CStr(Minute(dDateTime))
                    If Second(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                    sDateTime = sDateTime & CStr(Second(dDateTime))
                    If bDebug Then UpdateStatus "Last Time: " & sDateTime
                End If
                
                'found imgur album
                If InStr(1, aFile(x), "imgur.com/a/") > 0 _
                    And InStr(1, aFile(x), "domain/") = 0 _
                    And InStr(1, aFile(x), "i.imgur.com<") = 0 Then
                    bFoundImgurAlbum = True
                    sAlbum = Mid(aFile(x), InStr(1, aFile(x), "imgur.com") + 12, 5)
                    If bDebug Then UpdateStatus "ImgurAlbum: " & aFile(x)
                    If bDebug Then UpdateStatus "Album: " & sAlbum
                    If bDebug Then UpdateStatus "TimestampPath: " & sPath & "\" & sDateTime & "-" & sAlbum
                End If
                
                'found single image post/comment
                If InStr(1, aFile(x), "i.imgur.com") > 0 _
                    And InStr(1, aFile(x), "domain/") = 0 _
                    And InStr(1, aFile(x), "i.imgur.com<") = 0 Then
                    bFoundImgurSingle = True
                    sAlbum = Mid(aFile(x), InStr(1, aFile(x), "imgur.com") + 10, 7)
                    If InStr(1, sAlbum, ".") > 1 Then sAlbum = Left(sAlbum, 5)
                    If bDebug Then UpdateStatus "ImgurSingle: " & aFile(x)
                    If bDebug Then UpdateStatus "Album: " & sAlbum
                    If bDebug Then UpdateStatus "TimestampPath: " & sPath & "\" & sDateTime & "-" & sAlbum
                End If
        
                'if found single image without i subdomain
                If InStr(1, aFile(x), "//imgur.com/") > 0 _
                    And InStr(1, aFile(x), "jpg") = 0 _
                    And InStr(1, aFile(x), "/a/") = 0 Then
                    bFoundImgurSingleI = True
                    sAlbum = Mid(aFile(x), InStr(1, aFile(x), "imgur.com") + 10, 7)
                    If InStr(1, sAlbum, ".") > 1 Then sAlbum = Left(sAlbum, 5)
                    If bDebug Then UpdateStatus "ImgurSingleI: " & aFile(x)
                    If bDebug Then UpdateStatus "Album: " & sAlbum
                    If bDebug Then UpdateStatus "TimestampPath: " & sPath & "\" & sDateTime & "-" & sAlbum
                End If
                    
                    
                    
                
                'end of thing, get files
                If InStr(1, aFile(x), "share-button") > 0 Then
                    If bDebug Then UpdateStatus "End of link found."
                    
                    'output related link text files?
                    
                    'get found imgur album
                    If bFoundImgurAlbum Then
                        If chkTimeStamps Then
                            GetAlbum sAlbum, sPath & "\" & sDateTime & "-" & sAlbum
                        Else
                            GetAlbum sAlbum, sPath & "\" & sAlbum
                        End If
                    End If
                    
                    'get found imgur single images
                    If bFoundImgurSingle Then
                        If Len(Trim(sAlbum)) > 4 Then
                            UpdateStatus "Downloading single image: " & sAlbum
                            If chkTimeStamps Then
                                GetImage "http://i.imgur.com/" & sAlbum & ".jpg", sPath, sDateTime & "-" & sAlbum & ".jpg"
                            Else
                                GetImage "http://i.imgur.com/" & sAlbum & ".jpg", sPath
                            End If
                            lblCountDir = "Folder: 1/1 (100%)"
                            lblCountImg = "Image: 1/1 (100%)"
                            UpdateStatus "Download complete."
                        Else
                            UpdateStatus "Rejecting single image code: " & sAlbum
                        End If
                    End If
                
                    If bFoundImgurSingleI Then
                        If Len(Trim(sAlbum)) > 4 Then
                            UpdateStatus "Downloading single image: " & sAlbum
                            If chkTimeStamps Then
                                GetImage "http://i.imgur.com/" & sAlbum & ".jpg", sPath, sDateTime & "-" & sAlbum & ".jpg"
                            Else
                                GetImage "http://i.imgur.com/" & sAlbum & ".jpg", sPath
                            End If
                            lblCountDir = "Folder: 1/1 (100%)"
                            lblCountImg = "Image: 1/1 (100%)"
                            UpdateStatus "Download complete."
                        Else
                            UpdateStatus "Rejecting single image code: " & sAlbum
                        End If
                    End If
                
                    If (bFoundImgurAlbum Or bFoundImgurSingle Or bFoundImgurSingleI) _
                        And cboLinks.ListIndex < 2 Then
                    
                    End If
                
                    sAlbum = ""
                
                End If
        
                If InStr(1, aFile(x), "next &rsaquo;") > 0 Then
                    If Not chkRecent.Value And (chkRecent.Value And (CInt(txtRecent.Text) < (iPage + 1))) Then
                        bPageNext = True
                    End If
                End If
        
        
            End If
        
        Next
        DumpCollection cDataFullName
        AppendFile App.Path & "\col.txt", StarDate & StarTime & ": Page " & iPage
        iPage = iPage + 1
        
    Loop Until Not bPageNext
    
    sPath = txtPath
    
    lblCountPage.Caption = "Page: "
    
    GetRedditUser = True
    
    Exit Function
    
    frmMain.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    
End Function

Private Function GetSubreddit(inputcode As String) As Boolean
    
    Dim sCode As String
    Dim sFile As String
    Dim sCodeImage As String
    Dim sTitle As String
    Dim sFileOut As String
    Dim sPath As String
    Dim iAlbumMax As Integer
    Dim iAlbum As Integer
    Dim aFile
    
    Dim x As Integer
    Dim sAlbumMax As String
    Dim iPos, iMax As Integer
    Dim sTitleAlbum As String
    Dim sUsername As String
    Dim sSubreddit As String
    
    Dim sPostTypes As String
    
    Dim sAlbumNum, sAlbumDir As String
    Dim iPage As Integer
    Dim iCount As Integer
    
    Dim sDate, sTime As String
    Dim sDateTime, sDateTimeTemp As String
    Dim dDate, dTime, dDateTime As Date
    Dim sTimeZH, sTimeZM As String
    
    sCode = Trim$(inputcode)
    
    If InStr(1, sCode, "reddit.com/") > 1 Then
        sCode = Mid(sCode, InStrRev(sCode, "/") + 1)
    End If
    
    'pages 1
    sUsername = sCode
    sSubreddit = sCode
    
    UpdateStatus "Getting subreddit: " & sSubreddit
    
    sCode = "http://www.reddit.com/r/" & sCode & "/"
    sPath = txtPath & "\" & txtCode
    If Not oFSO.FolderExists(sPath) Then
        oFSO.CreateFolder (sPath)
    End If
    
    Dim bPageNext As Boolean
    Dim sDataFullName As String
    Dim sAlbum As String
    
    'http://www.reddit.com/user/spif/submitted/?count=25&after=t3_21zkxv
    'http://www.reddit.com/user/DeliriumTremens/submitted/?count=25&after=t3_y6l4h
    'http://www.reddit.com/user/DeliriumTremens/submitted/?count=50&after=t3_eokjn
    'http://www.reddit.com/user/DeliriumTremens/submitted/?count=75&after=t3_c0qzj
    'http://www.reddit.com/user/DeliriumTremens/submitted/?count=100&after=t3_aejzk
    '<div class=" thing id-t3_y6l4h odd link " onclick="click_thing(this)" data-fullname="t3_y6l4h" data-ups="89" data-downs="15" >
    '<div class="nav-buttons"><span class="nextprev">view more:&#32;<a href="http://www.reddit.com/user/DeliriumTremens/submitted/?count=26&amp;before=t3_vwhcv" rel="nofollow prev" >&lsaquo; prev</a><span class="separator"></span><a href="http://www.reddit.com/user/DeliriumTremens/submitted/?count=50&amp;after=t3_eokjn" rel="nofollow next" >next &rsaquo;</a></span></div>
        
    'pages
    iPage = 0
    
    Do
        lblCountPage.Caption = "Page: " & iPage + 1
        bPageNext = False
        If iPage >= 1 Then
            sCode = sCode & "/" & sPostTypes & "?count=" & iPage * 25 & _
                "&after=t3_" & sDataFullName
        End If
        GetHTMLFile sCode
        sFile = OpenFile(App.Path & "\download.htm")
        aFile = Split(sFile, ">")
        Debug.Print UBound(aFile)
        
        For x = 0 To UBound(aFile)
            
            If InStr(1, aFile(x), "the page you requested does not exist") > 0 Then
                UpdateStatus "User " & sUsername & " was not found."
                sPath = txtPath
                lblCountPage.Caption = "Page: "
                Exit Function
            End If
            
            If InStr(1, aFile(x), "next &rsaquo;") > 0 Then
                If Not chkRecent.Value And (chkRecent.Value And (CInt(txtRecent.Text) < (iPage + 1))) Then
                    bPageNext = True
                End If
            End If
            
            If InStr(1, aFile(x), "data-fullname") > 0 Then
                sDataFullName = Mid(aFile(x), InStr(1, aFile(x), "data-fullname") + 18, 5)
                If bDebug Then UpdateStatus "Last DataFullName: " & sDataFullName
            End If
            
            If InStr(1, aFile(x), "time title") > 0 Then
                sDateTimeTemp = Mid(aFile(x), InStr(1, aFile(x), "datetime=") + 10, 25)
                sDate = Replace(Left(sDateTimeTemp, 10), "-", "")
                dDate = DateSerial(CInt(Left(sDate, 4)), CInt(Mid(sDate, 5, 2)), CInt(Mid(sDate, 7, 2)))
                sTime = Mid(sDateTimeTemp, 12, 8)
                dTime = TimeSerial(CInt(Left(sTime, 2)), CInt(Mid(sTime, 4, 2)), CInt(Mid(sTime, 7, 2)))
                sTimeZH = Mid(sDateTimeTemp, 20, 3)
                sTimeZM = Mid(sDateTimeTemp, 20, 1) & Mid(sDateTimeTemp, 24, 2)
                dDateTime = dDate & " " & dTime
                dDateTime = DateAdd("h", CInt(sTimeZH), dDateTime)
                If CInt(sTimeZM) > 0 Then dDateTime = DateAdd("n", CInt(sTimeZM), dDateTime)
                sDateTime = CStr(Year(dDateTime))
                If Month(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                sDateTime = sDateTime & CStr(Month(dDateTime))
                If Day(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                sDateTime = sDateTime & CStr(Day(dDateTime))
                If Hour(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                sDateTime = sDateTime & CStr(Hour(dDateTime))
                If Minute(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                sDateTime = sDateTime & CStr(Minute(dDateTime))
                If Second(dDateTime) < 10 Then sDateTime = sDateTime & "0"
                sDateTime = sDateTime & CStr(Second(dDateTime))
                If bDebug Then UpdateStatus "Last Time: " & sDateTime
            End If
            
            If InStr(1, aFile(x), "imgur.com/a/") > 0 Then
                sAlbum = Mid(aFile(x), InStr(1, aFile(x), "imgur.com") + 12, 5)
                If bDebug Then UpdateStatus "Imgur: " & aFile(x)
                If bDebug Then UpdateStatus "Album: " & sAlbum
                If bDebug Then UpdateStatus "TimestampPath: " & sPath & "\" & sDateTime & "-" & sAlbum
                If chkTimeStamps Then
                    GetAlbum sAlbum, sPath & "\" & sDateTime & "-" & sAlbum
                Else
                    GetAlbum sAlbum, sPath & "\" & sAlbum
                End If
            End If
            If InStr(1, aFile(x), "i.imgur.com") > 0 Then
                sAlbum = Mid(aFile(x), InStr(1, aFile(x), "imgur.com") + 10, 7)
                If InStr(1, sAlbum, ".") > 1 Then sAlbum = Left(sAlbum, 5)
                If bDebug Then UpdateStatus "Imgur: " & aFile(x)
                If bDebug Then UpdateStatus "Album: " & sAlbum
                If bDebug Then UpdateStatus "TimestampPath: " & sPath & "\" & sDateTime & "-" & sAlbum
                If Len(Trim(sAlbum)) > 4 Then
                    UpdateStatus "Downloading single image: " & sAlbum
                    If chkTimeStamps Then
                        GetImage "http://i.imgur.com/" & sAlbum & ".jpg", sPath, sDateTime & "-" & sAlbum & ".jpg"
                    Else
                        GetImage "http://i.imgur.com/" & sAlbum & ".jpg", sPath
                    End If
                    lblCountDir = "Folder: 1/1 (100%)"
                    lblCountImg = "Image: 1/1 (100%)"
                    UpdateStatus "Download complete."
                Else
                    UpdateStatus "Rejecting single image code: " & sAlbum
                End If
                
            End If
        
        Next
        
        iPage = iPage + 1
    Loop Until Not bPageNext
    
    sPath = txtPath
    
    lblCountPage.Caption = "Page: "
    
    GetRedditUser = True
    
    Exit Function
    
    frmMain.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    
End Function


Private Function GetAlbumList(inputcode As String)

    Dim sCode As String
    Dim sFile As String
    Dim sCodeImage As String
    Dim sTitle As String
    Dim sFileOut As String
    Dim sPath As String
    Dim iAlbumMax As Integer
    Dim iAlbum As Integer
    Dim aFile
    
    Dim x As Integer
    Dim sAlbumMax As String
    Dim iPos, iMax As Integer
    Dim sTitleAlbum As String
    
    
    Dim sAlbumNum, sAlbumDir As String
    
    sCode = Trim$(inputcode)
    GetHTMLFile sCode
    sFile = OpenFile(App.Path & "\download.htm")
    aFile = Split(sFile, vbLf)

    sPath = txtPath
    
    'count albums first to number them in reverse
    iAlbum = 0
    iAlbumMax = 0
    For x = 0 To UBound(aFile)
        If InStr(1, aFile(x), "<div id=""album") > 0 Then
            iAlbumMax = iAlbumMax + 1
        End If
    Next x
    sAlbumMax = iAlbumMax
    
    
    For x = 0 To UBound(aFile)
        If InStr(1, aFile(x), "<title>") > 0 Then
            'create download folder from album title if not exists
            iPos = InStr(1, aFile(x), "<title>") + 7
            sTitle = Trim(Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), "'s albums") - iPos))
            sTitle = Replace(sTitle, vbTab, "")
            If sTitle = "" Then
                'use temp folder, stardate & scode
                sTitle = StarDate & StarTime & "-" & sCode
            End If
            Debug.Print "Owner: " & sTitle
            If Not oFSO.FolderExists(sPath & "\" & sTitle) Then
                oFSO.CreateFolder (sPath & "\" & StripInvalidFolderChars(sTitle))
            End If
            sPath = sPath & "\" & StripInvalidFolderChars(sTitle)
        End If
        
        If InStr(1, aFile(x), "<div id=""album") > 0 Then
            
            '<div id="album-xNdHX" data-title="Good afternoon!" data-cover="H1VUV" data-layout="g" data-privacy="0" data-description="" class="album ">
            iPos = InStr(1, aFile(x), "album-") + 6
            If iPos > 6 Then
                sCode = Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos)
                iPos = InStr(1, aFile(x), "data-title") + 12
                sTitleAlbum = Trim$(Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos))
                sTitleAlbum = Replace(sTitleAlbum, vbTab, "")
                If iAlbumMax < 10 Then
                    sAlbumNum = "00" & iAlbumMax
                ElseIf iAlbumMax < 100 Then
                    sAlbumNum = "0" & iAlbumMax
                Else
                    sAlbumNum = iAlbumMax
                End If
                If InStr(1, sTitleAlbum, "data-cover") > 0 Then
                    sAlbumDir = sAlbumNum
                Else
                    sAlbumDir = sAlbumNum & "-" & sTitleAlbum
                End If
                sAlbumDir = StripInvalidFolderChars(sAlbumDir)
                If iAlbum + 1 <= CInt(sAlbumMax) Then
                    lblCountDir = "Folder: " & iAlbum + 1 & "/" & sAlbumMax & " (" & Round(((CInt(iAlbum) + 1) / CInt(sAlbumMax)) * 100, 1) & "%)"
                    frmMain.Caption = "" & Round(((CInt(iAlbum) + 1) / CInt(sAlbumMax)) * 100, 1) & "% - " & App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
                End If
                DoEvents
                GetAlbum sCode, sPath & "\" & sAlbumDir
                iAlbumMax = iAlbumMax - 1
                iAlbum = iAlbum + 1
            End If
        End If
        If chkRecent And iAlbum > txtRecent - 1 Then
            'fix counters and issue done msg
            Exit For
        End If
        
    Next x
    frmMain.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    
End Function

Private Function GetAlbum(inputcode As String, Optional altpath As String) As Boolean

    Dim sCode As String
    Dim sFile As String
    Dim sCodeImage As String
    Dim sTitle As String
    Dim sFileOut As String
    Dim sPath As String
    Dim x, Y As Integer
    Dim iPicMax As Integer
    Dim aFile
    Dim iPos As Integer
    Dim sUrl, sExt As String
    Dim sImageNum, sImageTitle As String
    Dim sFileOutThumb As String
    
    sCode = Trim$(inputcode)
    If bDebug Then UpdateStatus "GetAlbum: " & "http://imgur.com/a/" & sCode & "/embed"
    GetHTMLFile "http://imgur.com/a/" & sCode & "/embed"
    sFile = OpenFile(App.Path & "\download.htm")
    aFile = Split(sFile, vbLf)
    
    sPath = txtPath
    
    'UpdateStatus "Downloading album: " & sCode
    
    'get picture count
    For x = 0 To UBound(aFile)
        If InStr(1, aFile(x), "<img") > 0 Then
            If InStr(1, aFile(x), "thumb-") + 6 > 6 Then iPicMax = iPicMax + 1
        End If
    Next x
    iPicMax = iPicMax + 1
    
    For x = 0 To UBound(aFile)
        If Len(altpath) > 1 Then
            sPath = altpath
            If Not oFSO.FolderExists(sPath) Then
                oFSO.CreateFolder (sPath)
            End If
        ElseIf InStr(1, aFile(x), "<title>") > 0 Then
            If InStr(1, aFile(x), "- Imgur") = 0 Then
                For Y = 1 To 5
                    If InStr(1, aFile(x + Y), "- Imgur") > 0 Then
                        sTitle = Trim(Mid(aFile(x + Y), 1, InStr(1, aFile(x + Y), "- Imgur") - 1))
                        Exit For
                    End If
                Next Y
            
            Else
                'create download folder from album title if not exists
                iPos = InStr(1, aFile(x), "<title>") + 7
                sTitle = Trim(Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), "- Imgur") - iPos))
            End If
        
            sTitle = Replace(sTitle, vbTab, "")
            If sTitle = "" Then
                'use temp folder, stardate & scode
                sTitle = StarDate & StarTime & "-" & sCode
            End If
            UpdateStatus "Downloading album: " & sCode & " - " & sTitle
            If Not oFSO.FolderExists(sPath & "\" & StripInvalidFolderChars(sTitle)) Then
                oFSO.CreateFolder (sPath & "\" & StripInvalidFolderChars(sTitle))
            End If
            sPath = sPath & "\" & StripInvalidFolderChars(sTitle)
        
        End If
        
        If InStr(1, aFile(x), "<img") > 0 Then
            '<img id="thumb-rxRZ8" class="unloaded thumb-title-embed" title="" alt="" data-src="http://i.imgur.com/rxRZ8s.jpg" data-index="77" />
            iPos = InStr(1, aFile(x), "thumb-") + 6
            If iPos > 6 Then
                sCodeImage = Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos)
                iPos = InStr(1, aFile(x), "data-src") + 10
                sUrl = Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos)
                sExt = Mid(sUrl, InStrRev(sUrl, "."))
                If InStr(1, sExt, "?") > 0 Then
                    sExt = Left(sExt, InStr(1, sExt, "?") - 1)
                End If
                iPos = InStr(1, aFile(x), "data-index") + 12
                sImageNum = Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos)
                iPos = InStr(1, aFile(x), "title") + 5
                sImageTitle = Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos)
                
                If sImageNum < 10 Then
                    sImageNum = "00" & sImageNum
                ElseIf sImageNum < 100 Then
                    sImageNum = "0" & sImageNum
                End If
                sFileOut = sImageNum & "-" & sCodeImage & sExt
                Debug.Print "FileOut: " & sFileOut
                If CInt(sImageNum) + 1 <= iPicMax - 1 Then
                    lblCountImg = "Image: " & CInt(sImageNum) + 1 & "/" & iPicMax - 1 & " (" & Round((CInt(sImageNum + 1) / CInt(iPicMax - 1)) * 100, 1) & "%)"
                    If optFunction(0).Value = False Then
                        frmMain.Caption = "" & Round((CInt(sImageNum + 1) / CInt(iPicMax - 1)) * 100, 1) & "% - " & App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
                    End If
                End If
                DoEvents
                If chkThumbs Then
                    sFileOutThumb = sImageNum & "-" & sCodeImage & "s" & sExt
                    If Not oFSO.FolderExists(sPath & "\thumb") Then
                        oFSO.CreateFolder (sPath & "\thumb")
                    End If
                    If Not oFSO.FileExists(sPath & "\thumb" & "\" & sFileOutThumb) Then
                        GetImage "http://i.imgur.com/" & sCodeImage & "s" & sExt, sPath & "\thumb", sFileOutThumb
                    Else
                        'UpdateStatus "Image exists, skipping " & sCodeImage & " thumbnail."
                    End If
                End If
                If Not oFSO.FileExists(sPath & "\" & sFileOut) Then
                    GetImage "http://i.imgur.com/" & sCodeImage & sExt, sPath, sFileOut
                Else
                    UpdateStatus "Image exists, skipping " & sImageNum & " - " & sCodeImage & "."
                End If

            End If
        End If
    Next x
    
    If Trim$(sTitle) <> "" Then
        UpdateStatus "Album " & sCode & " - " & sTitle & " download complete."
    Else
        UpdateStatus "Album " & sCode & " download complete."
    End If
    frmMain.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    Exit Function
    
ErrorHandler:


    
    
End Function

Public Function GetImageList(Optional ByVal rootpath As String) As String
    
    Dim cNumber As Collection
    
    Dim x, iLast, iSpot As Integer
    Dim bFinished As Boolean
    Dim sNumber As String
    Dim sList As String
    Dim sFile As String
    Dim aFile
    Dim bExists As Boolean
    Dim cItem
    
    If Trim(rootpath) <> "" Then
        sFile = OpenFile(rootpath)
    Else
        'sFile = OpenFile(App.Path & "\download.htm")
        
    End If
    Call InitializeCollection(cNumber)
    aFile = Split(sFile, " ")
    For x = 0 To UBound(aFile)
        If InStr(1, aFile(x), "res/") > 0 Then
            'add threadnumber to list and prevent dupes
            bExists = False
            sNumber = Mid(aFile(x), InStr(1, aFile(x), "res/") + 4, 5)
            'TODO: fix this to not rely on a hardcoded number length
            For Each cItem In cNumber
                If cItem = sNumber Then
                    bExists = True
                    Exit For
                End If
            Next
            If Not bExists Then cNumber.Add sNumber
        End If
    Next x
    
    Exit Function
    
    GetImageList = sList

End Function


Public Function GetImageList2(albumname As String) As Boolean

'download embed page


'get al







End Function






Public Function GetFile(ByVal url As String, Optional ByVal DestinationPath As String) As Boolean
    
On Error GoTo ErrorHandler
   
    Dim oWinHTTP, oStream, sDownload
    Dim sFilename As String
    
    If url = "" Then Exit Function
            
    sFilename = Mid(url, InStrRev(url, "/") + 1)
    
    UpdateStatus "GetFile: " & sFilename
    DoEvents
    
    Set oWinHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    oWinHTTP.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.13) Gecko/20080311 Firefox/2.0.0.13"
    oWinHTTP.SetRequestHeader "Accept", "*/*"
    oWinHTTP.Open "GET", url, False
    oWinHTTP.Send
    If oWinHTTP.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write oWinHTTP.responseBody
        If Trim(DestinationPath) = "" Then
            If Not oFSO.FolderExists(App.Path & "\manual") Then oFSO.CreateFolder (App.Path & "\manual")
            sDownload = App.Path & "\manual\" & sFilename
        Else
            If Right(DestinationPath, 1) = "\" Then
                sDownload = DestinationPath & sFilename
            Else
                sDownload = DestinationPath & "\" & sFilename
            End If
        End If
        oStream.SaveToFile sDownload
        oStream.Close
    Else
        UpdateStatus "GetFile WinHttp Status: " & oWinHTTP.Status
        GetFile = False
    End If
    
    If IsImage(sFilename) And chkPreview.Value = 1 Then
        Set Image1.Picture = LoadPicture(sDownload)
        Image1.Refresh
        DoEvents
    End If
    
    GetFile = True
    
    Exit Function
ErrorHandler:
    UpdateStatus "GetFile Error " & Err.Number & ": " & Err.Description
    GetFile = False

End Function

Function GetHTMLFile(ByVal url As String, Optional ByVal rootpath As String) As String

    Dim oWinHTTP, oStream
    Dim sFile, sCopy, sPage As String
    If bDebug Then UpdateStatus "GetHTMLFile: " & url
        
    If oFSO.FileExists(App.Path & "\download.htm") Then oFSO.DeleteFile App.Path & "\download.htm"
    
    If Trim$(Left(url, 7)) <> "http://" Then
        url = "http://" & Trim$(url)
    End If
    
    Set oWinHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    oWinHTTP.Open "GET", url, False
    oWinHTTP.Send
        If oWinHTTP.Status = 200 Then
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write oWinHTTP.responseBody
            sFile = App.Path & "\download.htm"
            oStream.SaveToFile sFile
            oStream.Close
            If Len(rootpath) > 1 Then
                If Trim(Right(url, 1)) <> "/" Then
                    sCopy = rootpath & "\" & Mid(url, InStrRev(url, "/") + 1)
                    If InStr(1, sCopy, "?") > 0 Then
                        sCopy = Left(sCopy, InStrRev(sCopy, "?") - 1)
                    End If
                    If LCase(Right(sCopy, 3)) = "php" Then
                        If InStr(1, sCopy, "?b=") > 0 Then
                            sPage = Mid(url, InStr(1, url, "?b=") + 3)
                            sPage = Left(sPage, InStr(1, sPage, "&") - 1)
                            sCopy = Left(sCopy, InStr(1, sCopy, ".") - 1) & sPage & ".htm"
                        End If
                        'http://www.anonib.com/_pleasantlyplump/index.php?b=1&g=0
                    End If
                Else
                    sCopy = rootpath & "\download.htm"
                End If
                oFSO.CopyFile App.Path & "\download.htm", sCopy
            End If
        Else
            UpdateStatus "HTML WinHttp Status: " & oWinHTTP.Status
        End If
        If rootpath <> "" Then
            GetHTMLFile = sCopy
        Else
            GetHTMLFile = sFile
        End If
End Function

Public Function GetImage(ByVal url As String, Optional ByVal rootpath As String, Optional ByVal renamefile As String) As String
On Error GoTo ErrorHandler

    Dim oWinHTTP, oStream, sDownload
    Dim sFilename As String
    
    Dim sTemp, sTemp2 As String
    Dim urlprefix As String
    If Trim(url) = "" Then Exit Function
    sTemp = urlprefix & url
    sFilename = Mid(url, InStrRev(url, "/") + 1)
    
    If sFilename = "blank.gif" Then Exit Function
    
    If Not oFSO.FolderExists(rootpath) Then
        sTemp2 = App.Path
    Else
        sTemp2 = rootpath
    End If
    sDownload = sTemp2 & "\" & sFilename
    If oFSO.FileExists(sDownload) Then
        Debug.Print "GetImage found, exiting: " & sDownload
        GetImage = ""
        Exit Function
    End If
    
    Set oWinHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    oWinHTTP.Open "GET", sTemp, False
    oWinHTTP.Send
    If oWinHTTP.Status = 200 Or Not oFSO.FileExists(App.Path & "\dev\" & url) Then
        Debug.Print "GetImage: " & sDownload
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write oWinHTTP.responseBody
        oStream.SaveToFile sDownload
        oStream.Close
        If IsImage(sTemp2 & "\" & sFilename) And chkPreview.Value = 1 Then
            Set Image1.Picture = LoadPicture(sDownload)
            Image1.Refresh
            DoEvents
        End If
        If Len(renamefile) > 1 Then
            oFSO.CopyFile sDownload, sTemp2 & "\" & renamefile, True
            oFSO.DeleteFile sDownload, True
        End If
    Else
        UpdateStatus "GetImage WinHttp Status: " & oWinHTTP.Status
    End If
    GetImage = ""

Exit Function
ErrorHandler:
    If Err.Number = 3004 Then
        Debug.Print "GetImage: Write to file failed: " & sDownload
        Resume Next
    End If

End Function

Function OpenFile(FileName As String) As String

On Error GoTo ErrorHandler
    Dim oFile, sFile
    Set oFile = oFSO.OpenTextFile(FileName, 1)
    sFile = oFile.ReadAll
    oFile.Close
    OpenFile = sFile
    Set oFile = Nothing
    Exit Function
    
ErrorHandler:
    Select Case Err.Number
        Case 53 'File not found
            UpdateStatus "OpenFile Error: File not found: " & FileName
            Exit Function
        Case 62 'File empty
            UpdateStatus "Data file empty: " & FileName
        Case Else
            UpdateStatus "OpenFile Error " & Err.Number & ": " & Err.Description & ": " & FileName
    End Select

End Function

Function DumpCollection(inputcol As Collection)

    Dim cItem
    For Each cItem In inputcol
        AppendFile App.Path & "\col.txt", StarDate & StarTime & ": " & cItem
    Next

End Function

Function AppendFile(ByVal FileName As String, ByVal AppendText As String) As String
    Dim oTS, oStream
    If Not oFSO.FileExists(FileName) Then oFSO.CreateTextFile FileName, True
    Set oTS = oFSO.OpenTextFile(FileName, 8)
    oTS.WriteLine AppendText
    oTS.Close
    Set oTS = Nothing
End Function

Function LogFile(ByVal AppendText As String) As String
    Dim oTS, oStream
    If Not oFSO.FileExists(App.Path & "\" & App.EXEName & ".log") Then oFSO.CreateTextFile App.Path & "\" & App.EXEName & ".log", True
    Set oTS = oFSO.OpenTextFile(App.Path & "\" & App.EXEName & ".log", 8)
    oTS.WriteLine StarDate & " " & StarTime & ": " & AppendText
    oTS.Close
    Set oTS = Nothing
End Function

Public Function ParseFile() As String

    Dim x, iLast, iSpot As Integer
    Dim bFinished As Boolean
    Dim sImage As String
    Dim sList As String
    Dim sFile As String
      
    Dim cImage As Collection
   
    sFile = OpenFile(App.Path & "\download.htm")
    sFile = Replace(sFile, vbCrLf, "")
    iLast = 1
    bFinished = False
    '<a href="/members/afullimage.php?cam=deviantdawl&fav=20070510030528-deviantdawl.jpg" target="big">
    Do Until bFinished = True
        iSpot = InStr(iLast, sFile, "&fav=")
        iLast = iSpot + 5
        sImage = Mid(sFile, iLast, 26)
        'If LCase(Right(sImage, "jpg")) = "jpg" Then
        cImage.Add sImage
        'UpdateStatus "ParseFile: " & sImage
        iLast = iLast + 26
        If cImage.Count > 25 Then bFinished = True
        'If iLast > Len(sFile) Then bFinished = True
    Loop
    
    ParseFile = sList

End Function

Function StarDate()
    Dim sMonth, sDay As String
    sMonth = Month(Now)
    sDay = Day(Now)
    If sMonth < 10 Then sMonth = "0" & sMonth
    If sDay < 10 Then sDay = "0" & sDay
    StarDate = Year(Now) & sMonth & sDay
End Function

Function StarTime()
    Dim sHr, sMin, sSec As String
    sHr = Hour(Now)
    sMin = Minute(Now)
    sSec = Second(Now)
    If sHr < 10 Then sHr = "0" & sHr
    If sMin < 10 Then sMin = "0" & sMin
    If sSec < 10 Then sSec = "0" & sSec
    StarTime = sHr & sMin & sSec
End Function

Function UpdateStatus(updatetext As String)
    txtStatus = txtStatus & StarTime & ": " & updatetext & vbCrLf
    If Len(txtStatus) > 32768 Then txtStatus = Right(txtStatus, 16384)
    txtStatus.SelStart = Len(txtStatus)
    DoEvents
    If bDebug Then LogFile (StarDate & StarTime & ": " & updatetext)
End Function

'Function UpdateBar(updatetext As String)
'    sBar.SimpleText = updatetext
'    DoEvents
'End Function

Private Sub InitializeCollection(ByVal inputcollection As Collection)
    Dim cItem
    For Each cItem In inputcollection
        If inputcollection.Count > 0 Then inputcollection.Remove 1
    Next
End Sub

Private Function StripHTML(ByVal inputstring) As String
    Dim lngStart, lngEnd, strHTML
    inputstring = Replace(inputstring, vbTab, "")
    inputstring = Replace(inputstring, vbCrLf, "")
    inputstring = Trim(inputstring)
    Do
        lngStart = InStr(inputstring, "<")
        lngEnd = InStr(inputstring, ">")
        strHTML = Mid(inputstring, lngStart, lngEnd - lngStart + 1)
        inputstring = Trim(Replace(inputstring, strHTML, ""))
    Loop Until Not InStr(inputstring, "<") And Not InStr(inputstring, ">")
    If InStr(inputstring, "<") Then inputstring = StripHTML(Trim(inputstring))
    StripHTML = Trim(inputstring)
End Function

Function IsImage(FileName As String) As Boolean
    Select Case LCase(Right(FileName, 3))
        Case "jpg", "peg", "gif", "bmp", "png"
            IsImage = True
        Case Else
            IsImage = False
    End Select
End Function

Private Sub LockFormControls(bInput As Boolean)

    Dim x As Integer

    fraOptions.Enabled = Not bInput
    For x = 0 To optFunction.Count - 1
        optFunction(x).Enabled = Not bInput
    Next x
    
    txtCode.Enabled = Not bInput
    txtPath.Enabled = Not bInput
    cmdBrowse.Enabled = Not bInput
    chkPreview.Enabled = Not bInput
    chkThumbs.Enabled = Not bInput
    cmdGo.Enabled = Not bInput

    cboLinks.Enabled = Not bInput
    

End Sub

Function PageType(inputurl As String) As Integer
    'detect whether album list or album
    
    PageType = 0
    Exit Function
    '0 = AlbumList
    '1 = Album
    '2 = Image

    ' 's album string in title

    Dim sUrl, sCode As String
    Dim sFile As String
    Dim aFile
    Dim x As Integer
    Dim iPos As Integer
    Dim sPath As String
    Dim sTitle As String
        Dim sCodeImage, sExt As String
    Dim sImageNum, sFileOut, sFileOutThumb As String
    
    
    If Len(Trim$(sUrl)) < 8 Then
        'code: album or image
    
    End If
    
    sUrl = inputurl
    sCode = Trim$(sUrl)
    GetHTMLFile "http://imgur.com/a/" & sCode & "/embed"
    sFile = OpenFile(App.Path & "\download.htm")
    aFile = Split(sFile, vbLf)

    sPath = txtPath
    
    For x = 0 To UBound(aFile)
        If InStr(1, aFile(x), "<title>") > 0 Then
            'create download folder from album title if not exists
            iPos = InStr(1, aFile(x), "<title>") + 7
            sTitle = Trim(Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), "- Imgur") - iPos))
            sTitle = Replace(sTitle, vbTab, "")
            If sTitle = "" Then
                'use temp folder, stardate & scode
                sTitle = StarDate & StarTime & "-" & sCode
            End If
            Debug.Print "Title: " & sTitle
            If Not oFSO.FolderExists(sPath & "\" & sTitle) Then
                oFSO.CreateFolder (sPath & "\" & sTitle)
            End If
            sPath = sPath & "\" & sTitle
        End If
        If InStr(1, aFile(x), "<img") > 0 Then
            '<img id="thumb-rxRZ8" class="unloaded thumb-title-embed" title="" alt="" data-src="http://i.imgur.com/rxRZ8s.jpg" data-index="77" />
            iPos = InStr(1, aFile(x), "thumb-") + 6
            If iPos > 6 Then
                sCodeImage = Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos)
                iPos = InStr(1, aFile(x), "data-src") + 10
                sUrl = Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos)
                sExt = Mid(sUrl, InStrRev(sUrl, "."))
                iPos = InStr(1, aFile(x), "data-index") + 12
                sImageNum = Mid(aFile(x), iPos, InStr(iPos + 1, aFile(x), """") - iPos)
                If sImageNum < 10 Then
                    sImageNum = "00" & sImageNum
                ElseIf sImageNum < 100 Then
                    sImageNum = "0" & sImageNum
                End If
                sFileOut = sImageNum & "-" & sCodeImage & sExt
                Debug.Print "FileOut: " & sFileOut
                'http://i.imgur.com/rxRZ8s.jpg
                If chkThumbs Then
                    sFileOutThumb = sImageNum & "-" & sCodeImage & "s" & sExt
                    If Not oFSO.FolderExists(sPath & "\thumb") Then
                        oFSO.CreateFolder (sPath & "\thumb")
                    End If
                    GetImage "http://i.imgur.com/" & sCodeImage & "s" & sExt, sPath & "\thumb", sFileOutThumb
                End If

                GetImage "http://i.imgur.com/" & sCodeImage & sExt, sPath, sFileOut
            End If
        End If
    Next x

End Function

Private Function StripInvalidFolderChars(inputstring)

    '/\:?*"<>|
    Dim sOut As String
    
    If Trim$(inputstring) = "" Then Exit Function
    sOut = inputstring
    sOut = Replace(sOut, "/", "-")
    sOut = Replace(sOut, "\", "-")
    sOut = Replace(sOut, ":", "-")
    sOut = Replace(sOut, "?", "-")
    sOut = Replace(sOut, "*", "-")
    sOut = Replace(sOut, """", "-")
    sOut = Replace(sOut, "<", "-")
    sOut = Replace(sOut, ">", "-")
    sOut = Replace(sOut, "|", "-")
    sOut = Replace(sOut, ",", "-")
    sOut = Replace(sOut, "..", "-")
'    sOut = Replace(sOut, "&", "-")
'    sOut = Replace(sOut, "#", "-")
'    sOut = Replace(sOut, ";", "-")
    sOut = Replace(sOut, vbTab, "-")
    
    If Right(sOut, 3) = "..." Then sOut = Left(sOut, Len(sOut) - 3)
    
    StripInvalidFolderChars = sOut

End Function


Function StripImgur(inputstring) As String
    Dim sOut As String
    sOut = Trim$(inputstring)
    If InStr(1, sOut, "http://imgur.com/a/") > 0 Then
        sOut = Replace(sOut, "http://imgur.com/a/", "")
    End If
    If InStr(1, sOut, "#") > 0 Then
        sOut = Left(sOut, InStr(1, sOut, "#") - 1)
    End If
    StripImgur = sOut
End Function


Private Sub optFunction_Click(Index As Integer)

    Dim x As Integer

    DoEvents

    If Index = 1 Or Index = 3 Or Index = 4 Then
        chkRecent.Enabled = True
        If chkRecent Then
            txtRecent.Enabled = True
        Else
            txtRecent.Enabled = False
        End If
        lblRecent.Enabled = True
        If Index = 3 Then
            cboRedditUser.Enabled = True
        Else
            cboRedditUser.Enabled = False
        End If
        If Index = 1 Then
            lblRecent.Caption = "most recent albums"
            lblLinks.Enabled = False
            cboLinks.Enabled = False
        ElseIf Index = 3 Or Index = 4 Then
            lblRecent.Caption = "most recent pages"
            lblLinks.Enabled = True
            cboLinks.Enabled = True
        End If
        
    Else
        chkRecent.Enabled = False
        txtRecent.Enabled = False
        lblRecent.Enabled = False
        cboRedditUser.Enabled = False
        lblLinks.Enabled = False
        cboLinks.Enabled = False
    End If


End Sub

Public Sub GetRegValues()
    
    Dim x As Integer
    Dim bTemp As Boolean
    
    If oFSO.FolderExists(RegGet("txtPath")) Then
        txtPath.Text = RegGet("txtPath")
    Else
        txtPath.Text = "c:\"
    End If
    If IsNumeric(RegGet("chkTimeStamps")) Then chkTimeStamps.Value = RegGet("chkTimeStamps")
    If IsNumeric(RegGet("chkPreview")) Then chkPreview.Value = RegGet("chkPreview")
    If IsNumeric(RegGet("chkDebug")) Then chkDebug.Value = RegGet("chkDebug")
    
    'If IsNumeric(RegGet("optFunction")) Then optFunction(CInt(RegGet("optFunction"))).Value = True
    If IsNumeric(RegGet("optFunction")) Then
        Debug.Print "optFunction: " & CStr(RegGet("optFunction"))
        DoEvents
        'optFunction_Click (CInt(RegGet("optFunction")))
        optFunction(CInt(RegGet("optFunction"))).Value = True
    End If
    'bTemp = False
    'For x = 0 To optFunction.Count - 1
    '    If optFunction(x).Value Then bTemp = True
    'Next x
    'If Not bTemp Then optFunction(0).Value = True
    
    If IsNumeric(RegGet("cboRedditUser")) Then cboRedditUser.ListIndex = CInt(RegGet("cboRedditUser"))
    If IsNumeric(RegGet("chkRecent")) Then chkRecent.Value = RegGet("chkRecent")
    If IsNumeric(RegGet("txtRecent")) Then
        txtRecent.Text = RegGet("txtRecent")
    Else
        txtRecent.Text = "5"
    End If
    
    If chkRecent Then
        txtRecent.Enabled = True
    Else
        txtRecent.Enabled = False
    End If
    
    If IsNumeric(RegGet("cboLinks")) Then cboLinks.ListIndex = CInt(RegGet("cboLinks"))
    
End Sub

Public Sub StoreRegValues()
    
    Dim x As Integer
    
    If Trim$(txtPath.Text) <> "" Then RegSet "txtPath", txtPath.Text
    RegSet "chkTimeStamps", chkTimeStamps.Value
    RegSet "chkPreview", chkPreview.Value
    RegSet "chkDebug", chkDebug.Value
    For x = 0 To optFunction.Count - 1
        If optFunction(x).Value = True Then
            RegSet "optFunction", CStr(x)
            Exit For
        End If
    Next x
    RegSet "cboRedditUser", cboRedditUser.ListIndex
    RegSet "chkRecent", chkRecent.Value
    If Trim$(txtRecent.Text) <> "" Then RegSet "txtRecent", txtRecent.Text
    RegSet "cboLinks", cboLinks.ListIndex
    
End Sub

Public Function RegSet(inKeyName As String, inKeyValue As String) As Variant
    On Error GoTo ErrHand
    Dim c As New cRegistry
    With c
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\" & App.EXEName
        .ValueKey = inKeyName
        .ValueType = REG_SZ
        .Value = inKeyValue
    End With
    RegSet = True
    Exit Function
ErrHand:
    UpdateStatus "RegSet Error " & Err.Number & ": " & Err.Description
    RegSet = False
End Function

Public Function RegGet(inKeyName As String) As String
    On Error GoTo ErrHand
    Dim c As New cRegistry
    With c
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\" & App.EXEName
        .ValueKey = inKeyName
        .ValueType = REG_SZ
        RegGet = .Value
    End With
    Exit Function
ErrHand:
    UpdateStatus "RegGet Error " & Err.Number & ": " & Err.Description
End Function

