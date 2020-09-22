Attribute VB_Name = "Module1"
Option Explicit
'
'                 Dichiarazioni Mouse
'   Ritorna la posizione Assoluta del mouse in PIXEL  X e Y
'
Private Declare Function M_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINT) As Long
Private Type POINT
      X       As Long
      Y       As Long
End Type
Public pt      As POINT
'
'            Dichiarazioni per Form Trasparente
'
Public Const LWA_COLORKEY = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const BM_SETSTATE = &HF3

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'
'
'                         Sezione Mouse
'
Public Sub GetCursorPos(xX As Long, xY As Long)
  Call M_GetCursorPos(pt)
  xX = pt.X
  xY = pt.Y
End Sub
