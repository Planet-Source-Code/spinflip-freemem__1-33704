VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRounder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Chris Neuner
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim ret As Long
Public Sub RoundedBorder(ByVal Frm As Form)
    Dim PriorScaleMode As Long
    PriorScaleMode = Frm.ScaleMode
    If PriorScaleMode <> vbPixels Then Frm.ScaleMode = vbPixels
    ret = CreateRoundRectRgn((Frm.ScaleWidth / 10), (Frm.ScaleHeight / 10), Frm.ScaleWidth, Frm.ScaleHeight, (Frm.ScaleWidth / 10), (Frm.ScaleHeight / 1))
    SelectClipRgn Frm.hdc, ret
    GetClipRgn Frm.hdc, ret
    SetWindowRgn Frm.hWnd, ret, True
    If PriorScaleMode <> vbPixels Then Frm.ScaleMode = PriorScaleMode
End Sub
Private Sub Class_Terminate()
    Call DeleteObject(ret)
End Sub
