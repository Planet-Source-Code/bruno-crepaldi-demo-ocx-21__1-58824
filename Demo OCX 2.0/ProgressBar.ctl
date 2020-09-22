VERSION 5.00
Begin VB.UserControl ProgressBar 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   ScaleHeight     =   1320
   ScaleWidth      =   6465
   ToolboxBitmap   =   "ProgressBar.ctx":0000
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image ImgCur 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   3240
      Picture         =   "ProgressBar.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3060
   End
   Begin VB.Image ImgBack 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   0
      Picture         =   "ProgressBar.ctx":549E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3060
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: ProgressBar
' Nome del File ..: ProgressBar
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
'
Private M_Value As Long
Private M_MinValue As Long
Private M_MaxValue As Long
Private M_ViewValue As Boolean
'
Private CursRaporto As Double
Private CursRange As Long
'                                Dichiarazione Eventi
Public Event Change(Value As Long)
'
'                      Inizializza
'
Private Sub UserControl_Initialize()
     Call Sposta((M_Value - M_MinValue) * CursRaporto)
End Sub

'
'      Inizializza le Variabili ( Solo Progetazione )
'
Private Sub UserControl_InitProperties()
     M_Value = 0                   ' Valore Iniziale
     M_MinValue = 0                ' Valore Iniziale
     M_MaxValue = 100              ' Valore Iniziale
     M_ViewValue = True            ' Visualizza Numero
     
     UserControl.Height = 255      ' Altezza
     UserControl.Width = 1830      ' Larghezza
End Sub
'
'                                Property
'
'
Public Property Get Value() As Long
   Value = M_Value
End Property
Public Property Let Value(ByVal NewValue As Long)
   
   If NewValue > M_MaxValue Then NewValue = M_MaxValue
   If NewValue < M_MinValue Then NewValue = M_MinValue
   
   M_Value = NewValue
   PropertyChanged "Value"
   Call Sposta((M_Value - M_MinValue) * CursRaporto)
End Property
'
Public Property Get MinValue() As Long
   MinValue = M_MinValue
End Property
Public Property Let MinValue(ByVal NewValue As Long)
   M_MinValue = NewValue
   PropertyChanged "MinValue"
   CursRaporto = Raporto(M_MinValue, M_MaxValue)
End Property
'
Public Property Get MaxValue() As Long
   MaxValue = M_MaxValue
End Property
Public Property Let MaxValue(ByVal NewValue As Long)
   M_MaxValue = NewValue
   PropertyChanged "MaxValue"
   CursRaporto = Raporto(M_MinValue, M_MaxValue)
End Property
'
Public Property Get ViewValue() As Boolean
   ViewValue = M_ViewValue
End Property
Public Property Let ViewValue(ByVal NewValue As Boolean)
   M_ViewValue = NewValue
   PropertyChanged "ViewValue"
   LblValue.Visible = M_ViewValue
End Property
'
Public Property Get PictureBackG() As Picture
   Set PictureBackG = ImgBack.Picture
End Property

Public Property Set PictureBackG(ByVal NewPic As Picture)
   Set ImgBack.Picture = NewPic
   PropertyChanged "PictureBackG"
End Property
'
Public Property Get PictureForG() As Picture
   Set PictureForG = ImgCur.Picture
End Property

Public Property Set PictureForG(ByVal NewPic As Picture)
   Set ImgCur.Picture = NewPic
   PropertyChanged "PictureForG"
End Property
'
Public Property Get ValueColor() As OLE_COLOR
ValueColor = LblValue.ForeColor
End Property
Public Property Let ValueColor(ByVal NewValue As OLE_COLOR)
LblValue.ForeColor = NewValue
PropertyChanged "ValueColor"
End Property
'
'                 Read/Write Properties
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  M_Value = PropBag.ReadProperty("Value", 0)
  M_MinValue = PropBag.ReadProperty("MinValue", 0)
  M_MaxValue = PropBag.ReadProperty("MaxValue", 100)
  M_ViewValue = PropBag.ReadProperty("ViewValue", True)
  '
  CursRaporto = Raporto(M_MinValue, M_MaxValue)
  Call Sposta((M_Value - M_MinValue) * CursRaporto)
  Set ImgBack.Picture = PropBag.ReadProperty("PictureBackG", Nothing)
  Set ImgCur.Picture = PropBag.ReadProperty("PictureForG", Nothing)
  LblValue.ForeColor = PropBag.ReadProperty("ValueColor", &H0)
  LblValue.Visible = M_ViewValue
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Value", M_Value, 0)
  Call PropBag.WriteProperty("MinValue", M_MinValue, 0)
  Call PropBag.WriteProperty("MaxValue", M_MaxValue, 100)
  Call PropBag.WriteProperty("ViewValue", M_ViewValue, True)
  Call PropBag.WriteProperty("ValueColor", LblValue.ForeColor, &H0)
  Call PropBag.WriteProperty("PictureBackG", ImgBack.Picture, Nothing)
  Call PropBag.WriteProperty("PictureForG", ImgCur.Picture, Nothing)
End Sub
'
'                        Eventi
'
Private Sub ChangeEvent(Valore As Long)
    RaiseEvent Change(Valore)
End Sub
'
'                        Resizing
'
Private Sub UserControl_Resize()
    ImgBack.Left = 0
    ImgBack.Top = 0
    ImgBack.Width = ScaleWidth
    ImgBack.Height = ScaleHeight

    ImgCur.Width = ScaleWidth
    ImgCur.Height = ScaleHeight
    ImgCur.Left = -ImgCur.Width
    
    LblValue.Top = 0
    LblValue.Left = 0
    LblValue.Height = ScaleHeight
    LblValue.Width = ScaleWidth
    
    LblValue.FontSize = ScaleHeight / 22
    
End Sub
'
'
'
Private Sub Sposta(Posizione As Long)
    LblValue.Caption = M_Value
    ImgCur.Left = Posizione - ImgCur.Width
End Sub

Private Function Raporto(Min As Long, Max As Long) As Single
  CursRange = Max - Min
  Raporto = ScaleWidth / CursRange
End Function
