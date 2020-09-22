VERSION 5.00
Begin VB.UserControl Led 
   BackColor       =   &H00C8B4AC&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H00C8B4AC&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Led.ctx":0000
   Begin VB.Image ImgLedAcceso 
      Height          =   360
      Index           =   3
      Left            =   600
      Picture         =   "Led.ctx":0312
      Top             =   2160
      Width           =   390
   End
   Begin VB.Image ImgLedSpento 
      Height          =   360
      Index           =   3
      Left            =   120
      Picture         =   "Led.ctx":0AD6
      Top             =   2160
      Width           =   390
   End
   Begin VB.Image ImgLedAcceso 
      Height          =   360
      Index           =   2
      Left            =   600
      Picture         =   "Led.ctx":129A
      Top             =   1800
      Width           =   390
   End
   Begin VB.Image ImgLedAcceso 
      Height          =   360
      Index           =   1
      Left            =   600
      Picture         =   "Led.ctx":1A5E
      Top             =   1320
      Width           =   390
   End
   Begin VB.Image ImgLedSpento 
      Height          =   360
      Index           =   2
      Left            =   120
      Picture         =   "Led.ctx":2222
      Top             =   1800
      Width           =   390
   End
   Begin VB.Image ImgLedSpento 
      Height          =   360
      Index           =   1
      Left            =   120
      Picture         =   "Led.ctx":29E6
      Top             =   1320
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   0
      Picture         =   "Led.ctx":31AA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   390
   End
   Begin VB.Image ImgLedAcceso 
      Height          =   360
      Index           =   0
      Left            =   600
      Picture         =   "Led.ctx":396E
      Top             =   840
      Width           =   390
   End
   Begin VB.Image ImgLedSpento 
      Height          =   360
      Index           =   0
      Left            =   120
      Picture         =   "Led.ctx":4132
      Top             =   840
      Width           =   390
   End
End
Attribute VB_Name = "Led"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Led
' Nome del File ..: Led
' Data............: 27/11/2004
' Versione........: 1.0
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
'
Option Explicit

Private M_Status As Boolean
Private M_Colore As Long
'                                Dichiarazione Eventi
Public Event Change(Value As Integer)

'
'      Inizializza le Variabili ( Solo Progetazione )
'
Private Sub UserControl_InitProperties()
     
     M_Status = False
     M_Colore = 0
     
     UserControl.Height = 360
     UserControl.Width = 390
End Sub
'
'                        Resizing
'
Private Sub UserControl_Resize()
    Image1.Left = 0
    Image1.Top = 0
    
  '  UserControl.Height = 360
  '  UserControl.Width = 390
    
    Image1.Width = ScaleWidth
    Image1.Height = ScaleHeight
End Sub
'
'                       inizializa
'
Private Sub UserControl_Initialize()
  UserControl.Height = 360
  UserControl.Width = 390
  
End Sub
'
'                         Eventi
'
Private Sub ChangeEvent(Valore As Integer)
    RaiseEvent Change(Valore)
End Sub
'
'                                Property
'
'
'
'
'
Public Property Get Status() As Boolean
   Status = M_Status
End Property
Public Property Let Status(ByVal NewValue As Boolean)
   M_Status = NewValue
   PropertyChanged "Status"
   Call Stato(M_Status, M_Colore)
End Property
Public Property Get Colore() As Long
   Colore = M_Colore
End Property
Public Property Let Colore(ByVal NewValue As Long)
   M_Colore = NewValue
   PropertyChanged "Colore"
   Call Stato(M_Status, M_Colore)
End Property

'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  M_Status = PropBag.ReadProperty("Status", False)
  M_Colore = PropBag.ReadProperty("Colore", 0)
  Call Stato(M_Status, M_Colore)
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Status", M_Status, False)
  Call PropBag.WriteProperty("Colore", M_Colore, False)
End Sub
'
'
'         Inizio Routine Led
'
'
Private Sub Stato(Status As Boolean, Colore As Long)
 If Status = True Then
   Image1.Picture = ImgLedAcceso(Colore)
 Else
   Image1.Picture = ImgLedSpento(Colore)
 End If
End Sub
