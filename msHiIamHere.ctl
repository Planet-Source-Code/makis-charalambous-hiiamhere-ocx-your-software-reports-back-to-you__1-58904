VERSION 5.00
Begin VB.UserControl msHiIamHere 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1140
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "msHiIamHere.ctx":0000
   ScaleHeight     =   885
   ScaleWidth      =   1140
   ToolboxBitmap   =   "msHiIamHere.ctx":0016
   Begin VB.TextBox DataArrival 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   630
      Top             =   330
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "msHiIamHere.ctx":0328
      Top             =   30
      Width           =   285
   End
End
Attribute VB_Name = "msHiIamHere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim lblflag0 As Long
Dim lblflag1 As Long

Private sText As String
Private bSent As Boolean
Private iStatus As Integer
Private SmellySock As Integer
Private Rc As Integer
Private Bytes As Integer
Private ResponseCode As Integer

'This is for the WaitforResponse Routine
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'Default Property Values:
Const m_def_Activate = 0
Const m_def_Period1 = 7
Const m_def_Period2 = 30
Const m_def_Period3 = 90
Const m_def_InfoApplication = ""
Const m_def_InfoSerial = ""
Const m_def_InfoUserNAme = ""
Const m_def_InfoAddress = ""
Const m_def_InfoTelephones = ""
Const m_def_InfoUserEmail = ""
Const m_def_InfoSpare1 = ""
Const m_def_InfoSpare2 = ""
Const m_def_InfoSpare3 = ""
Const m_def_EmailFrom = ""
Const m_def_EmailTo = ""
Const m_def_EmailServer = ""
Const m_def_EmailSubject = ""
Const m_def_CheckEveryMinutes = 30
Const m_def_AppName = ""

'Property Variables:

Dim m_Period1 As Integer
Dim m_Period2 As Integer
Dim m_Period3 As Integer
Dim m_InfoApplication As String
Dim m_InfoSerial As String
Dim m_InfoUserNAme As String
Dim m_InfoAddress As String
Dim m_InfoTelephones As String
Dim m_InfoUserEmail As String
Dim m_InfoSpare1 As String
Dim m_InfoSpare2 As String
Dim m_InfoSpare3 As String
Dim m_EmailFrom As String
Dim m_EmailTo As String
Dim m_EmailServer As String
Dim m_EmailSubject As String
Dim m_CheckEveryMinutes As Integer
Dim m_AppName As Variant

'***************************************************************
'Connect to the server
'***************************************************************

Private Sub SndMe()

Dim StartupData As WSADataType
Dim SocketBuffer As sockaddr
Dim IpAddr As Long

' Here check if we want to send the info

'Ini the WinSocket
Rc = WSAStartup(&H101, StartupData)
Rc = WSAStartup(&H101, StartupData)
    
'Open a free Socket (with this source code you can also
'open several connections! Very useful for E-Mail Applications...)

SmellySock = socket(AF_INET, SOCK_STREAM, 0)
If SmellySock = SOCKET_ERROR Then
'    MsgBox "Cannot Create Socket."
    Exit Sub
End If

'Checks if the Hostname exists
If Rc = SOCKET_ERROR Then
        Exit Sub
End If

IpAddr = GetHostByNameAlias(m_EmailServer)

If IpAddr = -1 Then
'    MsgBox "Problem in sending report. Are you connected to the internet?"
    Exit Sub
End If

'This part is responsible for the connection
SocketBuffer.sin_family = AF_INET
SocketBuffer.sin_port = htons(25)
SocketBuffer.sin_addr = IpAddr
SocketBuffer.sin_zero = String$(8, 0)
    
Rc = connect(SmellySock, SocketBuffer, Len(SocketBuffer))

'If an error occured close the connection and
'send an error message to the text window
If Rc = SOCKET_ERROR Then
'        MsgBox "Cannot Connect to " + mailserver + _
'                            Chr$(13) + Chr$(10) + _
'                            GetWSAErrorString(WSAGetLastError())
'        MsgBox "Problem in sending report. Are you connected to the internet?"
        closesocket SmellySock
        Rc = WSACleanup()
        Exit Sub
End If

'Select Receive Window
Rc = WSAAsyncSelect(SmellySock, UserControl.DataArrival.hWnd, ByVal &H202, ByVal FD_READ Or FD_CLOSE)
    
    If Rc = SOCKET_ERROR Then
'        MsgBox "Cannot Process Asynchronously."
        closesocket SmellySock
        Rc = WSACleanup()
        Exit Sub
    End If

bSent = True
iStatus = 0

ResponseCode = 220
Call WaitForResponse

End Sub

'***************************************************************
'Transmit the E-Mail
'***************************************************************

Private Sub Transmit(iStage As Integer)

Dim sHelo As String, temp As String
Dim pos As Integer

On Error Resume Next

Select Case iStatus

Case 1:
    sHelo = m_EmailFrom
    pos = Len(sHelo) - InStr(sHelo, "@")
    sHelo = Right$(sHelo, pos)
    
    ResponseCode = 250
    WinSmellySockSendData ("HELO " & sHelo & vbCrLf)
    Call WaitForResponse

Case 2:
    ResponseCode = 250
    WinSmellySockSendData ("MAIL FROM: <" & Trim(m_EmailFrom) & ">" & vbCrLf)
    Call WaitForResponse

Case 3:
    ResponseCode = 250
    WinSmellySockSendData ("RCPT TO: <" & Trim(m_EmailTo) & ">" & vbCrLf)
    Call WaitForResponse

Case 4:
    ResponseCode = 354
    WinSmellySockSendData ("DATA" & vbCrLf)
    Call WaitForResponse

Case 5:

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If you want additional Headers like Date,etc.               !
'simply add them below                                       !
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    temp = temp & "From: " & m_EmailFrom & vbNewLine
    temp = temp & "To: " & m_EmailTo & vbNewLine
    temp = temp & "Subject: " & m_EmailSubject & vbNewLine

    'Header + Message
     temp = temp & vbCrLf & _
          "APPLICATION" & vbCrLf & _
           "--------------------------------------" & vbCrLf & _
           "Application : " & m_InfoApplication & vbCrLf & _
           "Version     : " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
           "Serial      : " & m_InfoSerial & vbCrLf & _
           "Name        : " & m_InfoUserNAme & vbCrLf & _
           "Address     : " & m_InfoAddress & vbCrLf & _
           "Telephones  : " & m_InfoTelephones & vbCrLf & _
           "email       : " & m_InfoUserEmail & vbCrLf & _
           "Info1       : " & m_InfoSpare1 & vbCrLf & _
           "Info2       : " & m_InfoSpare2 & vbCrLf & _
           "Info3       : " & m_InfoSpare3 & vbCrLf & _
           vbCrLf & _
           "COMPUTER" & vbCrLf & _
           "--------------------------------------" & vbCrLf & _
           "CompUserName : " & FindUserName & vbCrLf & _
           "NetUserName : " & FindNetUserName & vbCrLf & _
           "ComputerName : " & FindComputerName & vbCrLf & _
            vbCrLf & ""
    
    'Send the Message & close connection
    WinSmellySockSendData (temp)
    WinSmellySockSendData (vbCrLf & "." & vbCrLf)
    ResponseCode = 250
    Call WaitForResponse

Case 6:

' MsgBox "E-Mail was successfuly send!"
' Here update the variables in our file
    
    UpdateInfoCount

    WinSmellySockSendData ("QUIT" & vbCrLf)
    ResponseCode = 221
    Call WaitForResponse
    iStatus = 0
    bSent = False
End Select

End Sub


'**************************************************************
' Waits until time out, while waiting for response
'**************************************************************

Private Sub WaitForResponse()
Dim Start As Long
Dim Tmr As Long

'Works with an Api Declaration because it's more precious

Start = timeGetTime
While Bytes > 0
    Tmr = timeGetTime - Start
    DoEvents ' Let System keep checking for incoming response
    'Wait 50 seconds for response
    If Tmr > 50000 Then
       ' MsgBox "SMTP service error, timed out while waiting for response", 64, "Error!"
        Exit Sub
    End If
Wend
End Sub

Private Sub WinSmellySockSendData(sInfoToSend As String)

Dim Rc As Integer
Dim MsgBuffer As String * 2048

    MsgBuffer = sInfoToSend
    
    Rc = send(SmellySock, ByVal MsgBuffer, Len(sInfoToSend), 0)
        
    'If an error occurs send an error message and
    'reset the winSmellySock
    If Rc = SOCKET_ERROR Then
     '   MsgBox "Cannot Send Request." + _
     '           Chr$(13) + Chr$(10) + _
     '           Str$(WSAGetLastError()) + _
     '           GetWSAErrorString(WSAGetLastError())
    '    MsgBox "Problem in sending report. Are you connected to the internet?"
        closesocket SmellySock
        Rc = WSACleanup()
        Unload Me
        Exit Sub
    End If


End Sub

Public Sub DeActivate()

    On Error Resume Next
    
    closesocket SmellySock
    Rc = WSACleanup()
        
End Sub
Private Function FindUserName() As String
' Function that returns the name of the currently logged on user
' Example - MyString = FindUserName
    sText = Space(512)
    GetUserName sText, Len(sText)
    FindUserName = Trim$(sText)
    
End Function

Private Function FindNetUserName() As String
' Function that returns the netword name of the currently logged on user
' Example - MyString = FindNetUserName
    sText = Space(512)
    WNetGetUser vbNullString, sText, Len(sText)
    FindNetUserName = Trim$(sText)
End Function

Private Function FindComputerName() As String
' Function that returns the network name of the run time machine
' Example - MyString = FindComputerName
    sText = Space(512)
    GetComputerName sText, Len(sText)
    FindComputerName = Trim$(sText)
End Function

Private Sub UpdateInfoCount()
  
    If m_AppName <> "" Then
        
        SaveSetting m_AppName, "General", "lblflag0", (Val(GetSetting(m_AppName, "General", "lblflag0", "0")) + 1)
        SaveSetting m_AppName, "General", "lblflag1", CLng(Fix(Now))
        lblflag0 = Val(GetSetting(m_AppName, "General", "lblflag0", "0"))
        lblflag1 = Val(GetSetting(m_AppName, "General", "lblflag1", CLng(Fix(Now))))
        
    End If
    
End Sub
Private Sub GetInfo()
     
End Sub

Private Sub ContactMe()

     If m_AppName <> "" Then
        
        lblflag0 = Val(GetSetting(m_AppName, "General", "lblflag0", "0"))
        lblflag1 = Val(GetSetting(m_AppName, "General", "lblflag1", CLng(Fix(Now))))
        
        Select Case lblflag0
          
          Case 0 ' Send straight away
                   SndMe
          Case 1
                If lblflag1 < CLng(Now - m_Period1) Then SndMe
          Case 2
                If lblflag1 < CLng(Now - m_Period2) Then SndMe
          Case Is > 2
                If lblflag1 < CLng(Now - m_Period3) Then SndMe
        End Select
        
    End If
    
End Sub

'***************************************************************
'Routine for arriving Data
'***************************************************************

Private Sub DataArrival_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MsgBuffer As String * 2048

On Error Resume Next

    If SmellySock > 0 Then
        'Receive up to 2048 chars
        Bytes = recv(SmellySock, ByVal MsgBuffer, 2048, 0)
        
        If Bytes > 0 Then
                
        If bSent Then
            If ResponseCode = Left(MsgBuffer, 3) Then
            MsgBuffer = vbNullString
            iStatus = iStatus + 1
            Transmit iStatus
            Else
                closesocket (SmellySock)
                Rc = WSACleanup()
                SmellySock = 0
'                MsgBox "The Server responds with an unexpected Response Code!", vbOKOnly, "Error!"
                Unload Me
                Exit Sub
            End If
        End If

        ElseIf WSAGetLastError() <> WSAEWOULDBLOCK Then
            closesocket (SmellySock)
            Rc = WSACleanup()
            SmellySock = 0
        End If
    End If

End Sub


Private Sub UserControl_Resize()
  UserControl.Width = 330
  UserControl.Height = 330
End Sub
Private Sub Timer1_Timer()
    
    Static iMinutesPassed As Integer
        
    If iMinutesPassed < m_CheckEveryMinutes Then
       iMinutesPassed = iMinutesPassed + 1
    Else
       ContactMe
       iMinutesPassed = 0
    End If
    
End Sub

Public Property Get AppName() As Variant
Attribute AppName.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    AppName = m_AppName
    
End Property

Public Property Let AppName(ByVal New_AppName As Variant)
    m_AppName = New_AppName
    PropertyChanged "AppName"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    m_AppName = m_def_AppName
    m_CheckEveryMinutes = m_def_CheckEveryMinutes
    m_EmailFrom = m_def_EmailFrom
    m_EmailTo = m_def_EmailTo
    m_EmailServer = m_def_EmailServer
    m_EmailSubject = m_def_EmailSubject
    m_InfoApplication = m_def_InfoApplication
    m_InfoSerial = m_def_InfoSerial
    m_InfoUserNAme = m_def_InfoUserNAme
    m_InfoAddress = m_def_InfoAddress
    m_InfoTelephones = m_def_InfoTelephones
    m_InfoUserEmail = m_def_InfoUserEmail
    m_InfoSpare1 = m_def_InfoSpare1
    m_InfoSpare2 = m_def_InfoSpare2
    m_InfoSpare3 = m_def_InfoSpare3
    m_Period1 = m_def_Period1
    m_Period2 = m_def_Period2
    m_Period3 = m_def_Period3
        
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_AppName = PropBag.ReadProperty("AppName", m_def_AppName)
    m_CheckEveryMinutes = PropBag.ReadProperty("CheckEveryMinutes", m_def_CheckEveryMinutes)
    m_EmailFrom = PropBag.ReadProperty("EmailFrom", m_def_EmailFrom)
    m_EmailTo = PropBag.ReadProperty("EmailTo", m_def_EmailTo)
    m_EmailServer = PropBag.ReadProperty("EmailServer", m_def_EmailServer)
    m_EmailSubject = PropBag.ReadProperty("EmailSubject", m_def_EmailSubject)
    m_InfoApplication = PropBag.ReadProperty("InfoApplication", m_def_InfoApplication)
    m_InfoSerial = PropBag.ReadProperty("InfoSerial", m_def_InfoSerial)
    m_InfoUserNAme = PropBag.ReadProperty("InfoUserNAme", m_def_InfoUserNAme)
    m_InfoAddress = PropBag.ReadProperty("InfoAddress", m_def_InfoAddress)
    m_InfoTelephones = PropBag.ReadProperty("InfoTelephones", m_def_InfoTelephones)
    m_InfoUserEmail = PropBag.ReadProperty("InfoUserEmail", m_def_InfoUserEmail)
    m_InfoSpare1 = PropBag.ReadProperty("InfoSpare1", m_def_InfoSpare1)
    m_InfoSpare2 = PropBag.ReadProperty("InfoSpare2", m_def_InfoSpare2)
    m_InfoSpare3 = PropBag.ReadProperty("InfoSpare3", m_def_InfoSpare3)
    m_Period1 = PropBag.ReadProperty("Period1", m_def_Period1)
    m_Period2 = PropBag.ReadProperty("Period2", m_def_Period2)
    m_Period3 = PropBag.ReadProperty("Period3", m_def_Period3)
        
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AppName", m_AppName, m_def_AppName)
    Call PropBag.WriteProperty("CheckEveryMinutes", m_CheckEveryMinutes, m_def_CheckEveryMinutes)
    Call PropBag.WriteProperty("EmailFrom", m_EmailFrom, m_def_EmailFrom)
    Call PropBag.WriteProperty("EmailTo", m_EmailTo, m_def_EmailTo)
    Call PropBag.WriteProperty("EmailServer", m_EmailServer, m_def_EmailServer)
    Call PropBag.WriteProperty("EmailSubject", m_EmailSubject, m_def_EmailSubject)
    Call PropBag.WriteProperty("InfoApplication", m_InfoApplication, m_def_InfoApplication)
    Call PropBag.WriteProperty("InfoSerial", m_InfoSerial, m_def_InfoSerial)
    Call PropBag.WriteProperty("InfoUserNAme", m_InfoUserNAme, m_def_InfoUserNAme)
    Call PropBag.WriteProperty("InfoAddress", m_InfoAddress, m_def_InfoAddress)
    Call PropBag.WriteProperty("InfoTelephones", m_InfoTelephones, m_def_InfoTelephones)
    Call PropBag.WriteProperty("InfoUserEmail", m_InfoUserEmail, m_def_InfoUserEmail)
    Call PropBag.WriteProperty("InfoSpare1", m_InfoSpare1, m_def_InfoSpare1)
    Call PropBag.WriteProperty("InfoSpare2", m_InfoSpare2, m_def_InfoSpare2)
    Call PropBag.WriteProperty("InfoSpare3", m_InfoSpare3, m_def_InfoSpare3)
    Call PropBag.WriteProperty("Period1", m_Period1, m_def_Period1)
    Call PropBag.WriteProperty("Period2", m_Period2, m_def_Period2)
    Call PropBag.WriteProperty("Period3", m_Period3, m_def_Period3)
    
End Sub

Public Property Get CheckEveryMinutes() As Integer
Attribute CheckEveryMinutes.VB_Description = "Period on minutes to check it need to send and for an internet connection"
Attribute CheckEveryMinutes.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    CheckEveryMinutes = m_CheckEveryMinutes
End Property

Public Property Let CheckEveryMinutes(ByVal New_CheckEveryMinutes As Integer)
    m_CheckEveryMinutes = New_CheckEveryMinutes
    
    ' Minimum period every 10 minutes
    If m_CheckEveryMinutes < 10 Then
      m_CheckEveryMinutes = 10
    End If
    
    PropertyChanged "CheckEveryMinutes"
End Property
Public Property Get EmailFrom() As String
Attribute EmailFrom.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    EmailFrom = m_EmailFrom
End Property

Public Property Let EmailFrom(ByVal New_EmailFrom As String)
    m_EmailFrom = New_EmailFrom
    PropertyChanged "EmailFrom"
End Property

Public Property Get EmailTo() As String
Attribute EmailTo.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    EmailTo = m_EmailTo
End Property

Public Property Let EmailTo(ByVal New_EmailTo As String)
    m_EmailTo = New_EmailTo
    PropertyChanged "EmailTo"
End Property
Public Property Get EmailServer() As String
Attribute EmailServer.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    EmailServer = m_EmailServer
End Property

Public Property Let EmailServer(ByVal New_EmailServer As String)
    m_EmailServer = New_EmailServer
    PropertyChanged "EmailServer"
End Property

Public Property Get InfoApplication() As String
Attribute InfoApplication.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoApplication = m_InfoApplication
End Property

Public Property Let InfoApplication(ByVal New_InfoApplication As String)
    m_InfoApplication = New_InfoApplication
    PropertyChanged "InfoApplication"
End Property

Public Property Get InfoSerial() As String
Attribute InfoSerial.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoSerial = m_InfoSerial
End Property

Public Property Let InfoSerial(ByVal New_InfoSerial As String)
    m_InfoSerial = New_InfoSerial
    PropertyChanged "InfoSerial"
End Property

Public Property Get InfoUserNAme() As String
Attribute InfoUserNAme.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoUserNAme = m_InfoUserNAme
End Property

Public Property Let InfoUserNAme(ByVal New_InfoUserNAme As String)
    m_InfoUserNAme = New_InfoUserNAme
    PropertyChanged "InfoUserNAme"
End Property

Public Property Get InfoAddress() As String
Attribute InfoAddress.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoAddress = m_InfoAddress
End Property

Public Property Let InfoAddress(ByVal New_InfoAddress As String)
    m_InfoAddress = New_InfoAddress
    PropertyChanged "InfoAddress"
End Property

Public Property Get InfoTelephones() As String
Attribute InfoTelephones.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoTelephones = m_InfoTelephones
End Property

Public Property Let InfoTelephones(ByVal New_InfoTelephones As String)
    m_InfoTelephones = New_InfoTelephones
    PropertyChanged "InfoTelephones"
End Property

Public Property Get InfoUserEmail() As String
Attribute InfoUserEmail.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoUserEmail = m_InfoUserEmail
End Property

Public Property Let InfoUserEmail(ByVal New_InfoUserEmail As String)
    m_InfoUserEmail = New_InfoUserEmail
    PropertyChanged "InfoUserEmail"
End Property
Public Property Get InfoSpare1() As String
Attribute InfoSpare1.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoSpare1 = m_InfoSpare1
End Property

Public Property Let InfoSpare1(ByVal New_InfoSpare1 As String)
    m_InfoSpare1 = New_InfoSpare1
    PropertyChanged "InfoSpare1"
End Property
Public Property Get InfoSpare2() As String
Attribute InfoSpare2.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoSpare2 = m_InfoSpare2
End Property

Public Property Let InfoSpare2(ByVal New_InfoSpare2 As String)
    m_InfoSpare2 = New_InfoSpare2
    PropertyChanged "InfoSpare2"
End Property
Public Property Get InfoSpare3() As String
Attribute InfoSpare3.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InfoSpare3 = m_InfoSpare3
End Property

Public Property Let InfoSpare3(ByVal New_InfoSpare3 As String)
    m_InfoSpare3 = New_InfoSpare3
    PropertyChanged "InfoSpare3"
End Property
Public Property Get Period1() As Integer
Attribute Period1.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Period1 = m_Period1
End Property

Public Property Let Period1(ByVal New_Period1 As Integer)
    m_Period1 = New_Period1
    PropertyChanged "Period1"
End Property

Public Property Get Period2() As Integer
Attribute Period2.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Period2 = m_Period2
End Property

Public Property Let Period2(ByVal New_Period2 As Integer)
    m_Period2 = New_Period2
    PropertyChanged "Period2"
End Property
Public Property Get Period3() As Integer
Attribute Period3.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Period3 = m_Period3
End Property

Public Property Let Period3(ByVal New_Period3 As Integer)
    m_Period3 = New_Period3
    PropertyChanged "Period3"
End Property
Public Property Get Activate() As Boolean
    Activate = Timer1.Enabled
End Property

Public Property Let Activate(ByVal New_Activate As Boolean)
    Timer1.Enabled = New_Activate
End Property

