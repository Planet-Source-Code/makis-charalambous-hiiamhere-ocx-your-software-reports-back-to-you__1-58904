VERSION 5.00
Object = "*\AmsHiIamHerevbp.vbp"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mp3 Utility (Demo for HiIamHere)"
   ClientHeight    =   5820
   ClientLeft      =   3000
   ClientTop       =   4455
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9825
   Begin Project2.msHiIamHere msHiIamHere1 
      Left            =   9000
      Top             =   60
      _ExtentX        =   582
      _ExtentY        =   582
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1170
      MaxLength       =   255
      TabIndex        =   17
      Top             =   4650
      Width           =   8445
   End
   Begin VB.TextBox txtGenre 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   15
      Top             =   4230
      Width           =   375
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1170
      MaxLength       =   30
      TabIndex        =   14
      Top             =   3990
      Width           =   2895
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3750
      Width           =   795
   End
   Begin VB.TextBox txtAlbum 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1170
      MaxLength       =   30
      TabIndex        =   12
      Top             =   3510
      Width           =   2895
   End
   Begin VB.TextBox txtArtist 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1170
      MaxLength       =   30
      TabIndex        =   11
      Top             =   3270
      Width           =   2895
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1170
      MaxLength       =   30
      TabIndex        =   10
      Top             =   3030
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   4350
      Pattern         =   "*.mp3"
      TabIndex        =   4
      Top             =   420
      Width           =   5205
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   4275
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   90
      TabIndex        =   2
      Top             =   780
      Width           =   4245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rename File from ID tag"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   4845
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Rename ID from Filename and other info"
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Label lblID 
      Caption         =   "FileName"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   4650
      Width           =   945
   End
   Begin VB.Label lblID 
      Caption         =   "Genre"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   4230
      Width           =   945
   End
   Begin VB.Label lblID 
      Caption         =   "Comments"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   3990
      Width           =   945
   End
   Begin VB.Label lblID 
      Caption         =   "Year"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   3750
      Width           =   945
   End
   Begin VB.Label lblID 
      Caption         =   "Album"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   3510
      Width           =   945
   End
   Begin VB.Label lblID 
      Caption         =   "Artist"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   3270
      Width           =   945
   End
   Begin VB.Label lblID 
      Caption         =   "Title"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3030
      Width           =   945
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tag1 As ID3Tag
Dim sFileName As String


Private Sub Command1_Click()
  
    On Error Resume Next
    
    ' Rename filename according to ID tag info
    
    If Len(txtFileName.Text) <> 0 Then
       Name Dir1 & "\" & File1 As Dir1 & "\" & txtFileName.Text & ".mp3"
    End If
    File1.SetFocus
  
End Sub

Private Sub Command2_Click()
     
     On Error Resume Next
     
     ' Modify or add an ID tag to the mp3 file
     
     tag1.SongTitle = txtTitle.Text
     tag1.Artist = txtArtist.Text
     tag1.Year = txtYear.Text
     tag1.Comment = txtComments.Text
     tag1.Album = txtAlbum.Text
     
     SetID3Tag Dir1 & "\" & File1, tag1
     
     File1.SetFocus
     
End Sub

Private Sub Dir1_Change()
     File1 = Dir1
End Sub

Private Sub Drive1_Change()
      
      Dir1.Path = Drive1
      File1 = Dir1
      
End Sub

Private Sub File1_Click()
   
   Dim dd$
   
   ClearInfo
   
   ' Here we try to create a filename as "Actor - Songtitle - Year"
   If GetID3Tag(File1.Path & "\" & File1, tag1) Then
      
      sFileName = ""
      
      dd$ = StripZero$((tag1.Artist))
      
      sFileName = Trim(dd$)
            
      dd$ = StripZero$(tag1.SongTitle)
      
      If sFileName <> "" Then
         sFileName = sFileName & " - " & Trim(dd$)
      Else
         sFileName = Trim(dd$)
      End If
    
      dd$ = StripZero$(tag1.Year)
      If Trim(dd$) <> "" Then
         If sFileName <> "" Then
            sFileName = sFileName & " - " & Trim(dd$)
         Else
            sFileName = Trim(dd$)
         End If
      End If
      
      sFileName = StripZero$(sFileName)
      txtFileName.Text = sFileName
      
      ShowInfo
      
   Else
      lblID(0).Caption = "FileName"
      txtTitle.Text = File1
   End If
      
End Sub

Private Sub Form_Load()
    
    Dir1.Path = Drive1
    File1 = Dir1
    
    '-------- msHiIamHere -------------------------------------------------
    '
    ' MAKE SURE THAT YOU PUT THE RIGHT EMAILS AND MAIL SERVER
    
    msHiIamHere1.AppName = "Mp3Test" ' Used to store some info in the registry
    
    msHiIamHere1.CheckEveryMinutes = 30  ' How often to check for internet connection
    
    msHiIamHere1.Period1 = 7    ' Try to talk to me at one week from installation
    msHiIamHere1.Period2 = 30   ' After one month do the same
    msHiIamHere1.Period3 = 90   ' After that try to contact me every 3 months.
    
    msHiIamHere1.EmailFrom = "Mp3Test@AutoInfo.com"  ' Dummy really but needed
    msHiIamHere1.EmailServer = "mail.someserver.com" ' Your mailserver
    msHiIamHere1.EmailTo = "info@somecompany.com"    ' mail to send the info
    
    ' Info that will be included in the auto mail. It depents of what you want to send back
    ' Some info is automatically gathered by the activex itself.
    
    msHiIamHere1.InfoApplication = "Mp3Test"        ' Here use info from your
    msHiIamHere1.InfoUserName = "Makis Charles"     ' software user setup etc.
    msHiIamHere1.InfoAddress = ""                   ' The values on the right
    msHiIamHere1.InfoSerial = "AGC-434-HHFW"        ' are "hardcoded" values
    msHiIamHere1.InfoTelephones = ""                ' to demonstrate the
    msHiIamHere1.InfoUserEmail = ""                 ' control.
    msHiIamHere1.InfoSpare1 = ""                    '
    msHiIamHere1.InfoSpare2 = ""                    '
    msHiIamHere1.InfoSpare3 = ""                    '
    
    msHiIamHere1.Activate = True ' Now activate our buddy
    
    ' msHiIamHere1.DeActivate  ' on Unload form
    
    '-------- msHiIamHere -------------------------------------------------
          
End Sub

Function StripZero$(t$)
   
   Dim d$
   Dim i As Integer
   d$ = ""
   For i = 1 To Len(t$)
      If Asc(Mid$(t$, i, 1)) <> 0 And Mid$(t$, i, 1) <> "?" And Mid$(t$, i, 1) <> "/" Then
         d$ = d$ & Mid$(t$, i, 1)
      End If
   Next
   StripZero$ = d$
   
End Function

Sub ShowInfo()
   
   txtTitle.Text = tag1.SongTitle
   txtArtist.Text = tag1.Artist
   txtAlbum.Text = tag1.Album
   txtYear.Text = tag1.Year
   txtComments.Text = tag1.Comment
   txtGenre.Text = tag1.Genre
   
End Sub
Sub ClearInfo()
   
   lblID(0).Caption = "Title"
   txtTitle.Text = ""
   
   txtArtist.Text = ""
   txtAlbum.Text = ""
   txtYear.Text = ""
   txtComments.Text = ""
   txtGenre.Text = ""
   
   txtFileName.Text = ""
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    msHiIamHere1.Deactivate
End Sub
