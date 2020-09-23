VERSION 5.00
Begin VB.Form Lm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "FreeMem"
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   Icon            =   "FreeMem.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FreeMem.frx":030A
   ScaleHeight     =   750
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1800
      Top             =   240
   End
   Begin VB.PictureBox Dm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      Picture         =   "FreeMem.frx":2259
      ScaleHeight     =   750
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Lm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Roundform As New clsRounder
Dim mdown As Boolean
Dim mprevx As Integer
Dim mprevy As Integer


Private Sub Form_Load()
Dm.CurrentX = 10
Dm.CurrentY = 10
Lm.CurrentX = 10
Lm.CurrentY = 10
'Dm.FontBold = True
'Lm.FontBold = True
Lm.ForeColor = &HFFFFFF
Dm.ForeColor = &HFFFFFF
Dm.Left = 0
Dm.Height = Lm.Height
Dm.Top = 0
Me.Show
Dm.Refresh
Lm.Refresh
Call Timer1_Timer

End Sub

Private Sub Form_Paint()
Roundform.RoundedBorder Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Roundform = Nothing
End
End Sub

Private Sub Timer1_Timer()
Lm.Refresh
Dm.Refresh

Dim sngTotalPhys As Single, lngAvailPhys As Long
Dim lngTotalVir As Long, lngAvailVir As Long
Dim lngTotalPF As Long, lngAvailPF As Long
Dim MemStat As MEMORYSTATUS
Call GlobalMemoryStatus(MemStat)
sngTotalPhys = Round(MemStat.dwTotalPhys / 1024 / 1024, 1)
lngAvailPhys = Round(MemStat.dwAvailPhys / 1024 / 1024)
If Right(sngTotalPhys, 1) >= 5 Then sngTotalPhys = sngTotalPhys + 1
'Me.Line (0, 0)-(lngAvailPhys / sngTotalPhys * Me.Width, Me.Height), RGB(0, 0, 0), BF

Dm.Width = lngAvailPhys / sngTotalPhys * Me.Width


Dm.CurrentX = 500
Dm.CurrentY = 75
Lm.CurrentX = 500
Lm.CurrentY = 75
Lm.Print "Physical Memory - " & Int((lngAvailPhys / sngTotalPhys) * 100) & "% Free"
Dm.Print "Physical Memory - " & Int((lngAvailPhys / sngTotalPhys) * 100) & "% Free"

Dm.CurrentX = 500
Dm.CurrentY = 530
Lm.CurrentX = 500
Lm.CurrentY = 530
Lm.Print lngAvailPhys & "mb Available  " & Int(sngTotalPhys) & "mb Total"
Dm.Print lngAvailPhys & "mb Available  " & Int(sngTotalPhys) & "mb Total"

Me.Caption = "FreeMem - " & Int((lngAvailPhys / sngTotalPhys) * 100) & "% Free"
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Call butdown
Else
mdown = True
mprevx = X
mprevy = Y
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdown = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mdown Then
Me.Move (Me.Left + X - mprevx), (Me.Top + Y - mprevy)
End If
End Sub

Private Sub Dm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Call butdown
Else
mdown = True
mprevx = X
mprevy = Y
End If
End Sub

Private Sub Dm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mdown Then Me.Move (Me.Left + X - mprevx), (Me.Top + Y - mprevy)
End Sub
Private Sub Dm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdown = False
End Sub

Private Sub butdown()
Dim thestring As String
thestring = ""
thestring = thestring & "                      FreeMem                " & vbCrLf
thestring = thestring & "---------------------------------------------" & vbCrLf
thestring = thestring & " Coded by Spinflip@graalmail.com             " & vbCrLf
thestring = thestring & " With special thanks going out to...         " & vbCrLf
thestring = thestring & "    Neil Crosby                  " & vbCrLf
thestring = thestring & "    Chris  Neuner                 " & vbCrLf
thestring = thestring & "    Eric Osterheldt                " & vbCrLf
thestring = thestring & "          and                      " & vbCrLf
thestring = thestring & "    Everyone else :)          " & vbCrLf
thestring = thestring & vbCrLf & vbCrLf & vbCrLf & "             Would you like to exit? "
If MsgBox(thestring, vbYesNo, "FreeMem") = vbYes Then Unload Me
End Sub
