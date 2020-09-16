VERSION 5.00
Object = "*\AAxioTextBox.vbp"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin AxioTextBox.AxTextBox AxTextBox1 
      Height          =   360
      Left            =   150
      TabIndex        =   47
      Top             =   315
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BorderColorOnFocus=   0
      CueTextColor    =   0
   End
   Begin VB.OptionButton OptionColor 
      Alignment       =   1  'Right Justify
      Caption         =   "CueText Color"
      Height          =   210
      Index           =   5
      Left            =   750
      TabIndex        =   46
      Top             =   3435
      Width           =   1335
   End
   Begin VB.TextBox CueText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   975
      TabIndex        =   44
      Top             =   4620
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vía Text"
      Height          =   285
      Left            =   2940
      TabIndex        =   43
      Top             =   5550
      Width           =   1110
   End
   Begin VB.CommandButton cmdSetText 
      Caption         =   "Vía SetText"
      Height          =   285
      Left            =   1785
      TabIndex        =   42
      Top             =   5550
      Width           =   1110
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   615
      TabIndex        =   40
      Text            =   "123456"
      Top             =   5565
      Width           =   1125
   End
   Begin VB.OptionButton OptionColor 
      Alignment       =   1  'Right Justify
      Caption         =   "BorderColorOnFocus"
      Height          =   210
      Index           =   4
      Left            =   300
      TabIndex        =   39
      Top             =   2964
      Value           =   -1  'True
      Width           =   1785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Form 2"
      Height          =   360
      Left            =   6600
      TabIndex        =   38
      Top             =   5445
      Width           =   990
   End
   Begin VB.OptionButton OptionColor 
      Alignment       =   1  'Right Justify
      Caption         =   "BackColorOnFocus"
      Height          =   210
      Index           =   3
      Left            =   450
      TabIndex        =   37
      Top             =   2508
      Width           =   1635
   End
   Begin VB.Frame Frame4 
      Caption         =   "RoundCorner"
      Height          =   645
      Left            =   180
      TabIndex        =   34
      Top             =   1320
      Width           =   3105
      Begin VB.HScrollBar hRCorner 
         Height          =   270
         Left            =   495
         Max             =   50
         TabIndex        =   35
         Top             =   240
         Value           =   1
         Width           =   2355
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   270
         TabIndex        =   36
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdAutoSize 
      Caption         =   "AutoSize TextBox"
      Height          =   360
      Left            =   5595
      TabIndex        =   33
      Top             =   4125
      Width           =   1770
   End
   Begin VB.HScrollBar cSizeFont 
      Height          =   270
      Left            =   5595
      TabIndex        =   32
      Top             =   3720
      Width           =   2355
   End
   Begin VB.ListBox ListFont 
      Height          =   1230
      Left            =   3450
      TabIndex        =   31
      Top             =   3330
      Width           =   1980
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   900
      Left            =   5580
      ScaleHeight     =   840
      ScaleWidth      =   2265
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2805
      Width           =   2325
   End
   Begin VB.Frame Frame3 
      Caption         =   "Align"
      Height          =   960
      Left            =   5580
      TabIndex        =   25
      Top             =   1800
      Width           =   2220
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   75
         TabIndex        =   26
         Top             =   225
         Width           =   2040
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   810
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      Caption         =   "CaseText && Other"
      Height          =   1575
      Left            =   5580
      TabIndex        =   13
      Top             =   180
      Width           =   2220
      Begin VB.CheckBox Check1 
         Caption         =   "SelText on Focus"
         Height          =   195
         Left            =   195
         TabIndex        =   29
         Top             =   330
         Width           =   1830
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Normal"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   1245
         Width           =   1515
      End
      Begin VB.OptionButton Option2 
         Caption         =   "LoCase Letters"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   940
         Width           =   1515
      End
      Begin VB.OptionButton Option2 
         Caption         =   "UpCase Letters"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   635
         Width           =   1515
      End
   End
   Begin VB.OptionButton OptionColor 
      Alignment       =   1  'Right Justify
      Caption         =   "BorderColor"
      Height          =   210
      Index           =   2
      Left            =   930
      TabIndex        =   12
      Top             =   2736
      Width           =   1155
   End
   Begin VB.OptionButton OptionColor 
      Alignment       =   1  'Right Justify
      Caption         =   "ForeColor"
      Height          =   210
      Index           =   1
      Left            =   1080
      TabIndex        =   11
      Top             =   3195
      Width           =   1005
   End
   Begin VB.OptionButton OptionColor 
      Alignment       =   1  'Right Justify
      Caption         =   "BackColor"
      Height          =   210
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   2280
      Width           =   1005
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   2145
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1800
      ScaleWidth      =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2250
      Width           =   990
   End
   Begin VB.PictureBox PicMuestra1 
      Height          =   255
      Left            =   1725
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3765
      Width           =   270
   End
   Begin VB.PictureBox PicMuestra2 
      Height          =   255
      Left            =   1725
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4005
      Width           =   270
   End
   Begin VB.Frame Frame1 
      Caption         =   "FormatToString"
      Height          =   3105
      Left            =   3450
      TabIndex        =   1
      Top             =   180
      Width           =   2010
      Begin VB.OptionButton Option1 
         Caption         =   "Decimals"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   28
         Top             =   1140
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "IP Address"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   27
         Top             =   2790
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ChileanRUT"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   22
         Top             =   2515
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Dates"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   2240
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Money"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   18
         Top             =   1415
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Letters && Numbers"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   1690
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Numbers Only"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   865
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Percent"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   4
         Top             =   1965
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Letters Only"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   590
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Chars"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   315
         Width           =   1755
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CueText"
      Height          =   195
      Left            =   240
      TabIndex        =   45
      Top             =   4650
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TEXT"
      Height          =   195
      Left            =   180
      TabIndex        =   41
      Top             =   5595
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "La Propiedad SetText formatea automaticamente el texto recibido desde una función u otro control."
      Height          =   420
      Left            =   195
      TabIndex        =   23
      Top             =   5085
      Width           =   3780
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value :"
      Height          =   195
      Left            =   1290
      TabIndex        =   21
      Top             =   855
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Actual"
      Height          =   195
      Left            =   735
      TabIndex        =   9
      Top             =   3810
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Nuevo"
      Height          =   195
      Left            =   735
      TabIndex        =   8
      Top             =   4035
      Width           =   885
   End
   Begin VB.Label lblTextVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AxTextBox 1.x"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
'Función Api SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
  
'Constante para usar como mensaje con SendMessage.
'Recuperar el Item a partir de una coordenada
Private Const LB_ITEMFROMPOINT = &H1A9


Private Sub AxTextbox1_Change()
Text1.Text = CStr(AxTextBox1.Value)
End Sub

Private Sub AxTextBox1_LostFocus()
Text1.Text = CStr(AxTextBox1.Value)
End Sub

Sub Mostrar_Fuente(Fuente As String)
  
    With Picture1
     ' Si se posicionó en otro item redibuja la _
      nueva fuente, si no solo cambia la posición x / y
     If Picture1.FontName = Fuente Then
        '.Move Pos_x, Pos_y - 500
     Else
        'Limpiar el picture con cls
        .Cls
        'Asignar la nueva fuente , el tamaño y el color
        .FontName = Fuente
        .FontSize = 16
        .ForeColor = vbBlue
        ' establecer Ancho y alto del Picture
        .Width = Picture1.TextWidth(Fuente) + 250
        .Height = 600
        ' Dibujar la fuente con Print
        .CurrentX = 100
        .CurrentY = 100
        Picture1.Print CStr(Fuente)
        'Dibujar un borde de color al picture
        Picture1.Line (0, 0)-(Picture1.ScaleWidth, Picture1.ScaleHeight), &HFFC0C0, B
        'Hacer visible y posicionar al costado del mouse
        '.Move Pos_x, Pos_y - 500
          
'        If Not .Visible Then
'            .Visible = True
'        End If
  
    End If
    End With
  
End Sub

Private Sub Check1_Click()
AxTextBox1.SelTextOnFocus = Check1.Value
End Sub

Private Sub cmdAutoSize_Click()
  AxTextBox1.AutoHeight
End Sub

Private Sub cmdSetText_Click()
AxTextBox1.SetText = Text2.Text
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
AxTextBox1.Text = Text2.Text
End Sub

Private Sub cSizeFont_Change()
AxTextBox1.Font.Size = cSizeFont.Value
End Sub

Private Sub CueText_Change()
AxTextBox1.CueText = CueText.Text
End Sub

Private Sub Form_Load()
Dim i_Font As Integer
  
    'Bucle para cargar las fuentes instaladas en el control ListBox
    For i_Font = Screen.FontCount - 1 To 0 Step -1
        ListFont.AddItem Screen.Fonts(i_Font)
    Next
      
    ' Propiedades para el picturebox
    With Picture1
        .ScaleMode = vbTwips
        .BorderStyle = 0 ' sin borde
        .BackColor = &HC0FFFF ' color de fondo
        .DrawWidth = 5 ' ancho de la linea de dibujo
        .AutoRedraw = True ' habilitar la persistencia
        '.Visible = False ' ocultar el control
        List1.ZOrder 1
        .ZOrder 0
    End With
  
    With List1
      .AddItem "Left", 0
      .AddItem "Right", 1
      .AddItem "Center", 2
      .ListIndex = 1
    End With

  With cSizeFont
    .Max = 50
    .Min = 8
    .Value = 12
  End With
  
  lblTextVersion.Caption = "AxTextBox v" & AxTextBox1.Version
End Sub

Private Sub hRCorner_Change()
Label3.Caption = hRCorner.Value
AxTextBox1.CornerCurve = hRCorner.Value
End Sub

Private Sub List1_Click()
AxTextBox1.Alignment = List1.ListIndex
End Sub

Private Sub ListFont_Click()

Call Mostrar_Fuente(ListFont.List(ListFont.ListIndex))
AxTextBox1.Font = ListFont.Text
End Sub

'Private Sub ListFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''Para el Indice del Item seleccionado
'Dim indice As Long
''Paraobtener las coordenadas del item seleccionado ( .. en Pixeles )
'Dim XPoint As Long
'Dim YPoint As Long
'Dim ZPoint As Long
'
'
'    If ListFont.ListCount <> -1 Then
'        'Pasar las medidas a Pixeles
'        XPoint = CLng(X / Screen.TwipsPerPixelX)
'        YPoint = CLng(Y / Screen.TwipsPerPixelY)
'        ZPoint = CLng(YPoint * &H10000 + XPoint)
'
'        With ListFont
'            'Recuperar el número de índice del Item seleccionado
'            indice = SendMessage(.hWnd, LB_ITEMFROMPOINT, 0, ByVal ZPoint)
'
'            If indice >= 0 And indice <= .ListCount Then
'                'Selecionar el Item
'                .Selected(indice) = True
'                'Previsualizar la fuente en el PictureBox
'                Call Mostrar_Fuente(ListFont.List(indice))
'            End If
'        End With
'    End If
'End Sub

Private Sub Option1_Click(Index As Integer)
Option2(0).Enabled = False
Option2(1).Enabled = False
Option2(2).Enabled = False

Select Case Index
  Case 0
    AxTextBox1.FormatToString = AllChars
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    Option2(2).Enabled = True
  Case 1
    AxTextBox1.FormatToString = LettersOnly
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    Option2(2).Enabled = True
  Case 2
    AxTextBox1.FormatToString = NumbersOnly
  Case 3
    AxTextBox1.FormatToString = LettersAndNumbers
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    Option2(2).Enabled = True
  Case 4
    AxTextBox1.FormatToString = Money
    Check1.Enabled = True
  Case 5
    AxTextBox1.FormatToString = Percent
  Case 6
    AxTextBox1.FormatToString = Dates
  Case 7
    AxTextBox1.FormatToString = ChileanRUT
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    Option2(2).Enabled = True
  Case 8
    AxTextBox1.FormatToString = IPAddress
  Case 9
    AxTextBox1.FormatToString = Decimals
    Check1.Enabled = True
End Select

End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
  Case 0
    AxTextBox1.CaseText = Normal
  Case 1
    AxTextBox1.CaseText = LowerCase
  Case 2
    AxTextBox1.CaseText = UpperCase
End Select
End Sub

Private Sub OptionColor_Click(Index As Integer)
Select Case Index
  Case Is = 0
    PicMuestra1.BackColor = AxTextBox1.BackColor
  Case Is = 1
    PicMuestra1.BackColor = AxTextBox1.ForeColor
  Case Is = 2
    PicMuestra1.BackColor = AxTextBox1.BorderColor
  Case Is = 3
    PicMuestra1.BackColor = AxTextBox1.BackColorOnFocus
End Select
End Sub

Private Sub picColor_Click()
If OptionColor(0).Value = True Then
   'BackColor
    AxTextBox1.BackColor = PicMuestra2.BackColor
    PicMuestra1.BackColor = AxTextBox1.BackColor
ElseIf OptionColor(1).Value = True Then
    'ForeColor
    AxTextBox1.ForeColor = PicMuestra2.BackColor
    PicMuestra1.BackColor = AxTextBox1.ForeColor
ElseIf OptionColor(2).Value = True Then
    'BorderColor
    AxTextBox1.BorderColor = PicMuestra2.BackColor
    PicMuestra1.BackColor = AxTextBox1.BorderColor
ElseIf OptionColor(3).Value = True Then
    'BackColorOnFocus
    AxTextBox1.BackColorOnFocus = PicMuestra2.BackColor
    PicMuestra1.BackColor = AxTextBox1.BackColorOnFocus
ElseIf OptionColor(4).Value = True Then
    'BorderColorOnFocus
    AxTextBox1.BorderColorOnFocus = PicMuestra2.BackColor
    PicMuestra1.BackColor = AxTextBox1.BorderColorOnFocus
ElseIf OptionColor(5).Value = True Then
    'CueTextColor
    AxTextBox1.CueTextColor = PicMuestra2.BackColor
    PicMuestra1.BackColor = AxTextBox1.CueTextColor

End If

End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Color As Long
Color = picColor.Point(X, Y)
PicMuestra2.BackColor = Color
End Sub

