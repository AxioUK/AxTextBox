VERSION 5.00
Begin VB.UserControl AxTextBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   840
   ScaleWidth      =   4875
   ToolboxBitmap   =   "AxTextbox.ctx":0000
   Begin VB.Timer TMouse 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4635
      Top             =   510
   End
   Begin VB.PictureBox pB 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   75
      ScaleHeight     =   540
      ScaleWidth      =   4620
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   4620
      Begin VB.TextBox txtRaiz 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   150
         TabIndex        =   1
         Top             =   150
         Width           =   3915
      End
   End
End
Attribute VB_Name = "AxTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxTextBox
'Version  : 1.30RC
'Editor   : David Rojas [AxioUK]
'Date     : 07/08/2020
'Description : Another X TextBox with multiple string validations
'------------------------------------
Option Explicit

'Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
'-------------------------------------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
''------------------------------------------------------------
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'''-----------------------------------------------------------
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
''''----------------------------------------------------------
'Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
'''''----------------------------------------------------------
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
'Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
'Private Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
''''''----------------------------------------------------------
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As tTrackMouseEvent) As Long ' Win98 or later
Private Declare Function TrackMouseEvent2 Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As tTrackMouseEvent) As Long ' Win95 w/ IE 3.0
'Private Declare Function GetCapture Lib "user32.dll" () As Long
'Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
'Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
' para saber si el puntero se encuentra dentro de un rectángulo ( para las opciones del menú )
'Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
' Determines if the control's parent form/window is an MDI child window
'Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32.dll" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'''''''---------------------------------------------------
 
'TYPES----------------------
Private Type RECTF
  Left As Long
  Top As Long
  Width As Long
  Height As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type tTrackMouseEvent
    cbSize      As Long
    dwFlags     As Long
    hwndTrack   As Long
    dwHoverTime As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

'ENUMS----------------------
Public Enum RegionalConstant
  LOCALE_SCURRENCY = &H14
  LOCALE_SCOUNTRY = &H6
  LOCALE_SDATE = &H1D
  LOCALE_SDECIMAL = &HE
  LOCALE_SLANGUAGE = &H2
  LOCALE_SLONGDATE = &H20
  LOCALE_SMONDECIMALSEP = &H16
  LOCALE_SMONGROUPING = &H18
  LOCALE_SMONTHOUSANDSEP = &H17
  LOCALE_SNATIVECTRYNAME = &H8
  LOCALE_SNATIVECURRNAME = &H1008
  LOCALE_SNATIVEDIGITS = &H13
  LOCALE_SNEGATIVESIGN = &H51
  LOCALE_SSHORTDATE = &H1F
  LOCALE_STIME = &H1E
  LOCALE_STIMEFORMAT = &H1003
End Enum

Public Enum CharacterType
    AllChars
    LettersOnly
    NumbersOnly
    LettersAndNumbers
    Money
    Percent
    Fraction
    Decimals
    Dates
    ChileanRUT
    IPAddress
End Enum

Public Enum CaseType
    Normal
    UpperCase
    LowerCase
End Enum

Public Enum eAlignConst
   [Left Justify] = 0
   [Right Justify] = 1
   Center = 2
End Enum

Public Enum eEnterKeyBehavior
    [eNone] = 0
    [eKeyTab] = 1
    '[Validate] = 2
End Enum

Public Enum eSizeMode
    eTextWidth
    eTextHeight
    eBothSizes
End Enum


'CONSTANTS----------------------
Private Const EM_LINELENGTH As Long = &HC1
Private Const WrapModeTileFlipXY As Long = &H3
Private Const UnitPixel          As Long = &H2&
Private Const LOGPIXELSX         As Long = 88
Private Const LOGPIXELSY         As Long = 90
Private Const SmoothingModeAntiAlias As Long = 4
Private Const TME_LEAVE     As Long = &H2

'Default Property Values:
Private Const m_def_Alignment = 0
Private Const m_def_BackColor = vbWhite
Private Const m_def_BorderColor = vbBlack
Private Const m_def_FocusColor = vbYellow
Private Const m_def_FormatToString = 0
Private Const m_def_CaseText = 0

'VARIABLES----------------------
'You have to have MSScripting Runtime referenced : WshShell.SendKeys "{Tab}"
Dim WshShell   As Object
Dim FlechasTab As Boolean     'Usar las Flechas del cursor como Tabulador
Dim EnterTab   As Boolean     'Usar Enter como Tabulador
Dim nScale     As Single
Dim GdipToken  As Long

Private lRect             As RECT
Private m_bIsTracking     As Boolean
Private m_bTrackHandler32 As Boolean
Private m_bSuppMouseTrack As Boolean

'Property Variables:
Private m_FormatToString As CharacterType
Private m_CaseText       As CaseType
Private m_KeyBehavior    As eEnterKeyBehavior

Private m_BackColor     As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_FocusColor    As OLE_COLOR
Private m_ForeColor     As OLE_COLOR
Private m_BorderOnFocus As OLE_COLOR
Private m_CueTextColor  As OLE_COLOR

Private m_Alignment     As Integer
Private m_SelTextFocus  As Boolean
Private HaveFocus       As Boolean
Private m_SetText       As String
Private m_CornerCurve   As Long
Private m_CueText       As String
Private m_Text          As String

Private sDecimal        As String
Private sThousand       As String
Private sDateDiv        As String
Private sMoney          As String
Private iCount          As Integer
Private bCancel         As Boolean
Private txtBaseString   As String
Private SetSize         As Boolean
Private m_UseCue        As Boolean
'------------------

'Event Declarations:
Public Event Change()
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event EnterKeyPress()
Public Event Click()
Public Event DblClick()


'---------------------------------
Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
 MsgBox "AxTextBox v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "Another X TextBox with more features" & vbCrLf & vbCrLf & "______________________________________by AxioUK"
End Sub

'---------------------------------
Public Sub Refresh()
    txtRaiz_Validate bCancel
End Sub

Private Function IsFunctionSupported(sFunction As String, sModule As String) As Boolean
'/Determina si la Funcion es Soportada por la Libreria
Dim hModule As Long

    ' GetModuleHandle?
    hModule = GetModuleHandleA(sModule)
    If (hModule = 0) Then
        hModule = LoadLibrary(sModule)
    End If
    
    If (hModule) Then
        If (GetProcAddress(hModule, sFunction)) Then
            IsFunctionSupported = True
        End If
        FreeLibrary hModule
    End If
End Function

'/Start tracking of mouse leave event
Private Sub TrackMouseTracking(hwnd As Long)
Dim tEventTrack As tTrackMouseEvent
    
    With tEventTrack
        .cbSize = Len(tEventTrack)
        .dwFlags = TME_LEAVE
        .hwndTrack = hwnd
    End With
    If (m_bTrackHandler32) Then
        TrackMouseEvent tEventTrack
    Else
        TrackMouseEvent2 tEventTrack
    End If
End Sub

Private Function IsMouseOver(hwnd As Long) As Boolean
    Dim PT As POINTAPI
    GetCursorPos PT
    IsMouseOver = (WindowFromPoint(PT.X, PT.Y) = hwnd)
End Function

Public Property Get Version() As String
Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Function fGetLocaleInfo(Valor As RegionalConstant) As String
   Dim Simbolo As String
   Dim r1 As Long
   Dim r2 As Long
   Dim p As Integer
   Dim Locale As Long
     
   Locale = GetUserDefaultLCID()
   r1 = GetLocaleInfo(Locale, Valor, vbNullString, 0)
   'buffer
   Simbolo = String$(r1, 0)
   'En esta llamada devuelve el símbolo en el Buffer
   r2 = GetLocaleInfo(Locale, Valor, Simbolo, r1)
   'Localiza el espacio nulo de la cadena para eliminarla
   p = InStr(Simbolo, Chr$(0))
     
   If p > 0 Then
      'Elimina los nulos
      fGetLocaleInfo = Left$(Simbolo, p - 1)
   End If
     
End Function

Private Sub DrawBorders(lBorderColor As OLE_COLOR)
Dim Rgn As Long, uRect As RECTF

With UserControl
    .Cls
    .ScaleMode = 3
    .AutoRedraw = True
    
    uRect.Left = 1: uRect.Top = 1: uRect.Width = .ScaleWidth - 2: uRect.Height = .ScaleHeight - 2
    GDIpRoundRect .hdc, uRect, RGBtoARGB(m_BackColor, 100), RGBtoARGB(lBorderColor, 100), m_CornerCurve
    'pB
    Rgn = CreateRoundRectRgn(2, 2, pB.Width, pB.Height, m_CornerCurve, m_CornerCurve)
    SetWindowRgn pB.hwnd, Rgn, True
    'DeleteObject Rgn
    'UserControl
    Rgn = CreateRoundRectRgn(0, 0, .Width, .Height, m_CornerCurve, m_CornerCurve)
    SetWindowRgn .hwnd, Rgn, True
    'DeleteObject Rgn
    .ScaleMode = 1
End With

'txtRaiz.BackColor = m_BackColor
'pB.BackColor = m_BackColor

End Sub

Private Function EsRut(CadenA As String) As Boolean
Dim i As Byte
Dim Z As Byte
Dim CadenaLimpiA As String
Dim DiG As String
Dim XXXX As Byte

If CadenA <> Empty And Val(CadenA) <> 0 Then
    'Limpia Cadena
    For i = 1 To Len(CadenA)
        If (Mid(CadenA, i, 1)) = "-" Or (Mid(CadenA, i, 1)) = "." Then
            'pasa al siguiente espacio
        Else
            CadenaLimpiA = CadenaLimpiA + Mid(CadenA, i, 1)
        End If
    Next
    
    'Prepara Variables
    CadenA = CadenaLimpiA
    DiG = (Mid(CadenaLimpiA, (Len(CadenaLimpiA)), 1))
    If Asc(DiG) <= 47 Or Asc(DiG) >= 58 Then
        If DiG = "K" Or DiG = "k" Then
            DiG = "10"
        Else
           DiG = "12"
        End If
    End If
    
    CadenaLimpiA = Empty
    
    For i = 1 To (Len(CadenA) - 1)
        CadenaLimpiA = CadenaLimpiA + (Mid(CadenA, i, 1))
    Next
    
    CadenA = Empty
    i = Empty
    i = (Len(CadenaLimpiA))
    Z = 2
    While i <> 0
        If Z <> 8 Then
            CadenA = Val(CadenA) + (Val((Mid(CadenaLimpiA, i, 1))) * Z)
            Z = Z + 1
        Else
            Z = 2
            CadenA = Val(CadenA) + (Val((Mid(CadenaLimpiA, i, 1))) * Z)
            Z = Z + 1
        End If
        i = i - 1
    Wend
    
    Z = 11 - (Val(CadenA) - Int((Val(CadenA)) / 11) * 11)
    
    XXXX = Asc(DiG)
        If DiG = 0 And Z = 11 Then
            EsRut = True
        Else
                If Z = DiG Then
                    EsRut = True
                Else
                    EsRut = False
                End If
        End If
Else
    EsRut = False
End If
CadenA = Empty
CadenaLimpiA = Empty
End Function

Private Function FormatoRUT(sRUT As String) As String
Dim strRut As String, gPos As Integer

If Trim$(sRUT) = "" Then Exit Function

   If Len(sRUT) = 10 Then
      strRut = Mid$(sRUT, 1, 2) & "." & Mid$(sRUT, 3, 3) & "." & Mid$(sRUT, 6, 5)
   ElseIf Len(sRUT) <= 9 Then
      If Mid$(sRUT, Len(sRUT) - 1, 1) = "-" Then
        If Len(sRUT) = 8 Then
          strRut = Mid$(sRUT, 1, 3) & "." & Mid$(sRUT, 4, 5)
        Else
          strRut = Mid$(sRUT, 1, 1) & "." & Mid$(sRUT, 2, 3) & "." & Mid$(sRUT, 5, 5)
        End If
      Else
        If Len(sRUT) = 7 Then
          strRut = Mid$(sRUT, 1, 3) & "." & Mid$(sRUT, 4, 3) & "-" & Right$(sRUT, 1)
        ElseIf Len(sRUT) = 8 Then
          strRut = Mid$(sRUT, 1, 1) & "." & Mid$(sRUT, 2, 3) & "." & Mid$(sRUT, 5, 3) & "-" & Right$(sRUT, 1)
        Else
          strRut = Mid$(sRUT, 1, 2) & "." & Mid$(sRUT, 3, 3) & "." & Mid$(sRUT, 6, 3) & "-" & Right$(sRUT, 1)
        End If
      End If
   Else
      strRut = sRUT
   End If

   FormatoRUT = UCase$(strRut)
End Function

Private Function InsertStr(ByVal InsertTo As String, ByVal Str As String, ByVal Position As Integer) As String
    Dim Str1 As String
    Dim Str2 As String
    
    Str1 = Mid$(InsertTo, 1, Position - 1)
    Str2 = Mid$(InsertTo, Position, Len(InsertTo) - Len(Str1))
    
    InsertStr = Str1 & Str & Str2
End Function

Public Function fCleanValue(sValor As String) As String
Dim sValue As Variant, i As Integer

   For i = 1 To Len(sValor)
      If IsNumeric(Mid(sValor, i, 1)) Or Mid(sValor, i, 1) = sDecimal Then
         sValue = sValue & Mid(sValor, i, 1)
      End If
   Next i

fCleanValue = Trim(sValue)
End Function

Private Sub pB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DrawBorders m_BorderOnFocus
TMouse.Enabled = True
End Sub

Private Sub TMouse_Timer()
If Not IsMouseOver(UserControl.hwnd) And Not IsMouseOver(pB.hwnd) And Not IsMouseOver(txtRaiz.hwnd) Then
  TMouse.Enabled = False
  DrawBorders m_BorderColor
End If
DoEvents
End Sub

Private Sub txtRaiz_Change()
If Not m_UseCue Then RaiseEvent Change
End Sub

Private Sub txtRaiz_Click()
    RaiseEvent Click
    If m_SelTextFocus = True Then
      txtRaiz.SelStart = 0
      txtRaiz.SelLength = Len(txtRaiz.Text)
    End If
End Sub

Private Sub txtRaiz_DblClick()
    RaiseEvent DblClick
    'DbleClick Selecciona el contenido
    txtRaiz.SelStart = 0
    txtRaiz.SelLength = Len(txtRaiz.Text)
End Sub

Private Sub txtRaiz_GotFocus()
With txtRaiz
  If LenB(Trim$(m_CueText)) <> 0& And LenB(Trim$(m_Text)) = 0& Then
    If LenB(.Tag) = 0& Then
      .Tag = .Text
      .Text = vbNullString
      .ForeColor = m_ForeColor '&HC0C0C0
    End If
  End If
  
  .SelStart = 0
  .SelLength = Len(.Text)
  .BackColor = m_FocusColor
  pB.BackColor = m_FocusColor
  iCount = 0
End With

End Sub

Private Sub txtRaiz_LostFocus()
With txtRaiz
    If LenB(Trim$(.Text)) = 0& Then  'LenB(Trim$(.Text)) = 0&
        .ForeColor = m_CueTextColor
        .Text = m_CueText  '.Tag
        .Tag = vbNullString
        m_Text = ""
        Exit Sub
    End If
    
  .BackColor = m_BackColor
  pB.BackColor = m_BackColor
  
  txtRaiz_Validate bCancel
End With
    
    Debug.Print "TextLenght:" & Len(txtRaiz.Text) & "/" & LenB(txtRaiz.Text)

End Sub

Private Sub txtRaiz_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    
    m_Text = txtRaiz.Text
    
    Select Case KeyCode
        Case 13 '39, 40, 13  Next Control: right arrow, down arrow and Enter
            WshShell.SendKeys "{Tab}"
        Case 37, 38 'Previous Control: left and up arrows
            WshShell.SendKeys "+{Tab}"
    End Select

End Sub

Private Sub txtRaiz_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    Dim lCurPos As Long
    Dim lLineLength As Long
    Dim i As Integer
    
    Dim MoneyDolB As Boolean
    Dim MoneyDotB As Boolean
    Dim MoneyDot As String
    Dim MoneyDolLoc As Long
    Dim MoneyDotLoc As Long
    
    Dim PercentDotB As Boolean
    Dim PercentPerB As Boolean
    Dim PercentNum As String
    Dim PercentDot As String
    Dim PercentLoc As Long
    Dim PercentDotLoc As Long
    
    Dim DecimalDotB As Boolean
    
    Dim Space As Boolean
    Dim FractionSlash As Boolean
    Dim SpaceLoc As Long
    Dim FractionLoc As Long
    
    Dim ipPoint As Integer
    
    txtRaiz.ForeColor = m_ForeColor
    
    Select Case FormatToString
    Case LettersOnly
        If Not (KeyAscii > 64 And KeyAscii < 91) And Not (KeyAscii > 96 And KeyAscii < 123) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    
    Case NumbersOnly
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    
    Case LettersAndNumbers
        If IsNumeric(Chr$(KeyAscii)) = False And (Not (KeyAscii > 64 And KeyAscii < 91) And Not (KeyAscii > 96 And KeyAscii < 123)) And KeyAscii <> 8 And KeyAscii <> 32 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    
    Case Money
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 36 And KeyAscii <> 44 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If txtRaiz.SelLength <> 0 Then
                Exit Sub
            End If
            
            ' Determine cursor position
            If txtRaiz.SelLength = 0 Then
                lCurPos = txtRaiz.SelStart
            Else
                lCurPos = txtRaiz.SelStart + txtRaiz.SelLength
            End If
            
            ' Determine textbox length
            lLineLength = SendMessage(txtRaiz.hwnd, EM_LINELENGTH, lCurPos, 0)
            
            ' Determine location/existance of "$" and ","
            For i = 1 To lLineLength
                If Mid$(txtRaiz.Text, i, 1) = "$" Then
                    MoneyDolB = True
                    MoneyDolLoc = i
                    Exit For
                End If
            Next i
            For i = 1 To lLineLength
                If Mid$(txtRaiz.Text, i, 1) = "," Then
                    MoneyDotB = True
                    MoneyDotLoc = i
                    Exit For
                End If
            Next i
                        
            ' Make sure number only goes to 2 decimal places
            If MoneyDotB = True Then
                'MoneyDot = Mid$(txtRaiz.Text, InStr(1, txtRaiz.Text, ",") + 1, Len(txtRaiz.Text) + InStr(1, txtRaiz.Text, ",") + 1)
                MoneyDot = Mid$(txtRaiz.Text, InStr(1, txtRaiz.Text, ",") + 1, Len(txtRaiz.Text) + 1)
                     
                If Len(MoneyDot) = 2 And lCurPos = MoneyDotLoc + 1 And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If Len(MoneyDot) = 2 And lCurPos = MoneyDotLoc And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If lCurPos = MoneyDotLoc + 2 And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
            End If
                
            ' Make sure "," and "$" is only typed once
            If KeyAscii = 36 And MoneyDolB = False Then
                MoneyDolB = True
            ElseIf KeyAscii = 36 And MoneyDolB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            If lCurPos <> 0 And MoneyDolB <> False And KeyAscii = 36 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If

            If KeyAscii = 44 And MoneyDotB = False Then
                MoneyDotB = True
            ElseIf KeyAscii = 44 And MoneyDotB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
        End If
    
    Case Percent
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 37 And KeyAscii <> Asc(sDecimal) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If txtRaiz.SelLength <> 0 Then
                Exit Sub
            End If
            
            ' Determine cursor position
            If txtRaiz.SelLength = 0 Then
                lCurPos = txtRaiz.SelStart
            Else
                lCurPos = txtRaiz.SelStart + txtRaiz.SelLength
            End If
            
            ' Determine textbox length
            lLineLength = SendMessage(txtRaiz.hwnd, EM_LINELENGTH, lCurPos, 0)
            
            ' Determine location of "%" and ","
            For i = 1 To lLineLength
                If Mid$(txtRaiz.Text, i, 1) = "%" Then
                    PercentPerB = True
                    PercentLoc = i
                    Exit For
                End If
            Next i
            For i = 1 To lLineLength
                If Mid$(txtRaiz.Text, i, 1) = sDecimal Then
                    PercentDotB = True
                    PercentDotLoc = i
                    Exit For
                End If
            Next i

            ' Make sure number only goes to 2 decimal places
            If PercentDotB = True Then
                PercentDot = Mid$(txtRaiz.Text, InStr(1, txtRaiz.Text, sDecimal) + 1, Len(txtRaiz.Text) + InStr(1, txtRaiz.Text, sDecimal) + 1)
        
                If InStr(1, PercentDot, "%") <> 0 Then
                    PercentDot = Mid$(PercentDot, 1, Len(PercentDot) - 1)
                End If
        
                If Len(PercentDot) = 2 And lCurPos = PercentDotLoc + 1 And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If Len(PercentDot) = 2 And lCurPos = PercentDotLoc And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If lCurPos = PercentDotLoc + 2 And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
            End If

            ' Make sure "%" and "," is only typed once
            If KeyAscii = 37 And PercentPerB = False Then
                PercentPerB = True
            ElseIf KeyAscii = 37 And PercentPerB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            If lCurPos <> Len(txtRaiz.Text) And PercentPerB <> False And KeyAscii = 37 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If KeyAscii = Asc(sDecimal) And PercentDotB = False Then
                MoneyDotB = True
            ElseIf KeyAscii = Asc(sDecimal) And PercentDotB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            ' Make sure numbers are not written after the "%"
            If KeyAscii <> 37 And KeyAscii <> 8 And PercentPerB = True And lCurPos = PercentLoc Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            ' Determine if the percentage is >100
            If IsNumeric(Chr$(KeyAscii)) = True Then
                PercentNum = txtRaiz.Text
                PercentNum = InsertStr(PercentNum, Chr$(KeyAscii), lCurPos + 1)
                If InStr(1, PercentNum, "%") <> 0 Then
                    If Val(Mid$(PercentNum, 1, Len(PercentNum) - 1)) > 100 Then
                        KeyAscii = 0
                        Beep
                        Exit Sub
                    End If
                Else
                    'If Val(PercentNum) > 100 Then
                    '    KeyAscii = 0
                    '    Beep
                    '    Exit Sub
                    'End If
                End If
            End If
        End If
    
    Case Fraction
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 47 And KeyAscii <> 32 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If txtRaiz.SelLength <> 0 Then
                Exit Sub
            End If
            ' Determine cursor position
            If txtRaiz.SelLength = 0 Then
                lCurPos = txtRaiz.SelStart
            Else
                lCurPos = txtRaiz.SelStart + txtRaiz.SelLength
            End If
            
            ' Determine textbox length
            lLineLength = SendMessage(txtRaiz.hwnd, EM_LINELENGTH, lCurPos, 0)
            
            ' Determine location of " " and "/"
            For i = 1 To lLineLength
                If Mid$(txtRaiz.Text, i, 1) = "/" Then
                    FractionLoc = i
                    Exit For
                End If
            Next i
    
            For i = 1 To lLineLength
                If Mid$(txtRaiz.Text, i, 1) = " " Then
                    SpaceLoc = i
                    Exit For
                End If
            Next i
            
            If FractionLoc <> 0 Then
                FractionSlash = True
            End If
            If SpaceLoc <> 0 Then
                Space = True
            End If
            
            ' Don't allow more then 1 space in the field
            If (Space = True Or Fraction = True) And KeyAscii = 32 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If Space = False And KeyAscii = 32 Then
                Space = True
            End If
            
            ' Check if " " is being used correctly
            If lCurPos = 0 And KeyAscii = 32 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            ' Don't allow more then 1 "/" in the field
            If FractionSlash = True And KeyAscii = 47 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If FractionSlash = False And KeyAscii = 47 Then
                FractionSlash = True
            End If
            
            ' Check if "/" is being used correctly
            If lLineLength >= 1 Then
                If lCurPos > 0 Then
                    If KeyAscii = 47 And IsNumeric(Mid$(txtRaiz.Text, lCurPos, 1)) = False Then
                        KeyAscii = 0
                        Beep
                        Exit Sub
                    End If
                End If
            ElseIf KeyAscii = 47 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
        End If
    
    Case Decimals
        If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 44 And KeyAscii <> 46 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            ' Determine textbox length
            lLineLength = SendMessage(txtRaiz.hwnd, EM_LINELENGTH, lCurPos, 0)
        
            ' Determine existance of ","
            For i = 1 To lLineLength
                If Mid$(txtRaiz.Text, i, 1) = sDecimal Then
                    DecimalDotB = True
                    Exit For
                End If
            Next i
                        
            ' Make sure Decimal separator is only typed once
            If sDecimal = Chr$(44) Then
              If KeyAscii = 44 And DecimalDotB = False Then
                  DecimalDotB = True
              ElseIf KeyAscii = 44 And DecimalDotB = True Then
                  KeyAscii = 0
                  Beep
                  Exit Sub
              End If
            ElseIf sDecimal = Chr$(46) Then
              If KeyAscii = 46 And DecimalDotB = False Then
                  DecimalDotB = True
              ElseIf KeyAscii = 46 And DecimalDotB = True Then
                  KeyAscii = 0
                  Beep
                  Exit Sub
              End If
            End If
        End If
    
    Case IPAddress
          ipPoint = 0
          If iCount >= 15 Then
            KeyAscii = 0
            Exit Sub
          End If
          
          If Len(txtRaiz.Text) = 0 Then iCount = 0
          
          For i = 1 To Len(txtRaiz.Text)
            If Mid$(txtRaiz.Text, i, 1) = "." Then ipPoint = ipPoint + 1
          Next i
          
          Select Case KeyAscii
            Case 8 'Borrar
              iCount = iCount - 1
              
            Case 48 To 57
              
            Case 46
              If Len(txtRaiz.Text) = 0 Then KeyAscii = 0
              If ipPoint = 3 Then KeyAscii = 0
              
            Case Else
              KeyAscii = 0
          End Select
    
    End Select
    
    Select Case CaseText
    Case UpperCase
        KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    Case LowerCase
        KeyAscii = Asc(LCase(Chr$(KeyAscii)))
    End Select
    
    If LenB(Trim$(m_Text)) <> 0& Then
      m_Text = txtRaiz.Text
    Else
      m_Text = vbNullString
    End If
    
End Sub

Private Sub txtRaiz_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtRaiz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DrawBorders m_BorderOnFocus
TMouse.Enabled = True
End Sub

Private Sub txtRaiz_Validate(Cancel As Boolean)
Dim sValor As String, i As Integer, sCar As String


  Select Case FormatToString
      Case Is = Money
        txtRaiz.Text = fCleanValue(txtRaiz.Text)
        txtRaiz.Text = sMoney & " " & Format$(txtRaiz.Text, "Standard")
        
    Case Is = Percent
        txtRaiz.Text = fCleanValue(txtRaiz.Text)

        If txtRaiz.Text = "" Then
            txtRaiz.Text = "0 %"
        End If
        If InStr(1, txtRaiz.Text, "%") = 0 Then
            txtRaiz.Text = txtRaiz.Text & " %"
        End If
        If InStr(1, txtRaiz.Text, "%") <> 0 Then
            If Len(txtRaiz.Text) = 1 Then
                txtRaiz.Text = "0 %"
            End If
        End If
        If InStr(1, txtRaiz.Text, sDecimal) <> 0 Then
            If Mid$(txtRaiz.Text, 1, Len(txtRaiz.Text) - 1) = sDecimal Then
                txtRaiz.Text = Mid$(txtRaiz.Text, 1, Len(txtRaiz.Text) - 2) & " %"
            End If
            If Mid$(txtRaiz.Text, 1, 1) = sDecimal Then
                txtRaiz.Text = "0" & Mid$(txtRaiz.Text, 1, Len(txtRaiz.Text))
            End If
        End If
        
    Case Is = NumbersOnly
      sValor = ""
      If txtRaiz.Text = "" Then txtRaiz.Text = "0"
      For i = 1 To Len(txtRaiz.Text)
           sCar = Mid$(txtRaiz.Text, i, 1)
           If IsNumeric(sCar) Then
              sValor = sValor & sCar
           ElseIf sCar = sDecimal Then
              Exit For
           End If
      Next i
      
      txtRaiz.Text = sValor
      
    Case Is = LettersAndNumbers
      sValor = ""
      If txtRaiz.Text = "" Then txtRaiz.Text = "0"
      For i = 1 To Len(txtRaiz.Text)
           sCar = Mid$(txtRaiz.Text, i, 1)
          If IsNumeric(sCar) Then
              sValor = sValor & sCar
          ElseIf (Asc(sCar) >= Asc("a") And Asc(sCar) <= Asc("z")) Or (Asc(sCar) >= Asc("A") And Asc(sCar) <= Asc("Z")) Then
              sValor = sValor & sCar
          End If
      Next i
      
      txtRaiz.Text = sValor
    
    Case Is = Fraction
        If txtRaiz.Text = "" Then
            txtRaiz.Text = "0"
        End If
        
        ' if the user inputs a fractional number
        If InStr(1, txtRaiz.Text, "/") <> 0 Then
            ' if / is the first character in the text box then set to 0
            If InStr(1, txtRaiz.Text, "/") = 1 Then
                txtRaiz.Text = "0"
            ' make sure there are numbers before and after the /
            ElseIf (IsNumeric(Mid$(txtRaiz.Text, InStr(1, txtRaiz.Text, "/") - 1, 1)) = False) Or (IsNumeric(Mid$(txtRaiz.Text, InStr(1, txtRaiz.Text, "/") + 1, 1)) = False) Then
                txtRaiz.Text = "0"
            End If
        End If
        txtRaiz.Text = Trim(txtRaiz.Text)
        
    Case Is = Decimals
        txtRaiz.Text = fCleanValue(txtRaiz.Text)
        If Trim$(txtRaiz.Text) = "" Then txtRaiz.Text = "0"
        txtRaiz.Text = FormatNumber(txtRaiz.Text, 2, vbTrue)
          
    Case Is = Dates
      If txtRaiz.Text = "" Or txtRaiz.Text = "00/00/0000" Then Exit Sub
      If Not IsDate(txtRaiz.Text) Then
          For i = 1 To Len(txtRaiz.Text)
            If i = 3 Or i = 5 Then
              sValor = sValor & "/" & Mid$(txtRaiz.Text, i, 1)
            Else
              sValor = sValor & Mid$(txtRaiz.Text, i, 1)
            End If
          Next i
      Else
          txtRaiz.Text = Format(txtRaiz, "Short Date")
          txtRaiz.ForeColor = vbBlack
      End If
      
      txtRaiz.Text = sValor
      
    Case Is = ChileanRUT
        txtRaiz.Text = FormatoRUT(txtRaiz)
        If EsRut(txtRaiz.Text) = False Then MsgBox "RUT no Válido...!", vbInformation + vbOKOnly, "Error!"
    
    Case Is = IPAddress
        Dim arIP() As String
        
        arIP = Split(txtRaiz.Text, ".")
        If UBound(arIP) <> 3 Then
          txtRaiz.ForeColor = vbRed
        Else
          For i = 0 To 3
           If (CInt(arIP(i)) > 255) Or (CInt(arIP(i)) < 0) Then
                  txtRaiz.ForeColor = vbRed
                  Exit For
           Else
              txtRaiz.Text = CInt(arIP(0)) & "." & CInt(arIP(1)) & "." & CInt(arIP(2)) & "." & CInt(arIP(3))
           End If
          Next i
        End If
  End Select
    
    txtRaiz.BackColor = m_BackColor
    pB.BackColor = m_BackColor
End Sub

Private Sub UserControl_Initialize()
Set WshShell = CreateObject("WScript.Shell")
InitGDI
nScale = GetWindowsDPI
'InitCommonControls
End Sub

Private Sub UserControl_InitProperties()
Set txtRaiz.Font = Ambient.Font
m_BackColor = m_def_BackColor
m_BorderColor = m_def_BorderColor
m_FocusColor = m_def_FocusColor
m_ForeColor = vbBlack
m_Alignment = m_def_Alignment
m_FormatToString = m_def_FormatToString
m_CaseText = m_def_CaseText
m_CornerCurve = 1
FlechasTab = False
HaveFocus = False
m_SelTextFocus = False
SetSize = False
m_SetText = ""
m_CueText = ""
m_Text = ""
txtRaiz.Text = ""
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DrawBorders m_BorderOnFocus
TMouse.Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_KeyBehavior = .ReadProperty("EnterKeyBehavior", eNone)
  txtRaiz.PasswordChar = .ReadProperty("PasswordChar", vbNullString)
  Set txtRaiz.Font = .ReadProperty("Font", Ambient.Font)
  txtRaiz.ForeColor = .ReadProperty("ForeColor", vbButtonText)
  txtRaiz.MaxLength = .ReadProperty("MaxLength", 0)
  txtRaiz.Locked = .ReadProperty("Locked", False)
  txtRaiz.Enabled = .ReadProperty("Enabled", True)
  txtRaiz.Text = .ReadProperty("Text", "")
  m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
  m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
  m_FocusColor = .ReadProperty("BackColorOnFocus", m_def_FocusColor)
  Alignment = .ReadProperty("Alignment", m_def_Alignment)
  m_SelTextFocus = .ReadProperty("SelTextOnFocus", False)
  m_FormatToString = .ReadProperty("FormatToString", m_def_FormatToString)
  m_CaseText = .ReadProperty("CaseText", m_def_CaseText)
  txtRaiz.Text = .ReadProperty("SetText", "")
  m_CornerCurve = .ReadProperty("CornerCurve", 1)
  m_BorderOnFocus = .ReadProperty("BorderColorOnFocus", vbRed)
  m_CueText = .ReadProperty("CueText", "")
  m_CueTextColor = .ReadProperty("CueTextColor", &HC0C0C0)
End With

sDecimal = fGetLocaleInfo(LOCALE_SDECIMAL)
sThousand = fGetLocaleInfo(LOCALE_SMONTHOUSANDSEP)
sDateDiv = fGetLocaleInfo(LOCALE_SDATE)
sMoney = fGetLocaleInfo(LOCALE_SCURRENCY)

m_bTrackHandler32 = IsFunctionSupported("TrackMouseEvent", "User32")
m_bSuppMouseTrack = m_bTrackHandler32
If Not m_bSuppMouseTrack Then m_bSuppMouseTrack = IsFunctionSupported("_TrackMouseEvent", "Comctl32")

If Trim$(m_Text) = "" Or Trim$(m_Text) = vbNullString And Trim$(m_CueText) <> "" Then
  txtRaiz.Text = m_CueText
  txtRaiz.ForeColor = m_CueTextColor
End If
End Sub

Private Sub UserControl_Resize()
'On Error Resume Next
'''''''''''''''''''''''''''''''''''''''''''''
If SetSize = True Then GoTo LResize

With pB
  .Top = 7
  .Left = 7
  .Height = UserControl.ScaleHeight - 5
  .Width = UserControl.ScaleWidth - 5
End With

'single line
With txtRaiz
  .Left = 50
  .Top = 10
  .Width = pB.ScaleWidth - 100
  .Height = pB.ScaleHeight - 5
  .Appearance = 0
End With

LResize:
DrawBorders m_BorderColor

Debug.Print "PictureH:" & pB.Height
Debug.Print "TextH:" & txtRaiz.Height

End Sub

Public Sub AutoHeight()
SetSize = True
  
txtRaiz.Height = TextHeight("Wq") + TextHeight(txtRaiz.Text) '+ 5
'txtRaiz.Width = TextWidth("WW") + TextWidth(txtRaiz.text)

pB.Height = ((txtRaiz.Height + txtRaiz.Top)) * (1.1 * nScale)
'pB.Width = txtRaiz.Width + 12

'UserControl.Width = pB.Width
UserControl.Height = (pB.Height + pB.Top + 1.4) * nScale
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
   .WriteProperty "EnterKeyBehavior", m_KeyBehavior, eNone
   .WriteProperty "PasswordChar", txtRaiz.PasswordChar, vbNullString
   .WriteProperty "Font", txtRaiz.Font, Ambient.Font
   .WriteProperty "ForeColor", txtRaiz.ForeColor, vbButtonText
   .WriteProperty "MaxLength", txtRaiz.MaxLength, 0
   .WriteProperty "Locked", txtRaiz.Locked, False
   .WriteProperty "Enabled", txtRaiz.Enabled, True
   .WriteProperty "Text", txtRaiz.Text, ""
   .WriteProperty "BackColor", m_BackColor, m_def_BackColor
   .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
   .WriteProperty "BackColorOnFocus", m_FocusColor, m_def_FocusColor
   .WriteProperty "Alignment", m_Alignment, m_def_Alignment
   .WriteProperty "SelTextOnFocus", m_SelTextFocus, False
   .WriteProperty "FormatToString", m_FormatToString, m_def_FormatToString
   .WriteProperty "CaseText", m_CaseText, m_def_CaseText
   .WriteProperty "SetText", txtRaiz.Text, ""
   .WriteProperty "CornerCurve", m_CornerCurve, 1
   .WriteProperty "BorderColorOnFocus", m_BorderOnFocus, vbRed
   .WriteProperty "CueText", m_CueText, ""
   .WriteProperty "CueTextColor", m_CueTextColor, &HC0C0C0
End With

End Sub

Public Property Get Alignment() As eAlignConst
Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal NewAlignment As eAlignConst)
m_Alignment = NewAlignment
txtRaiz.Alignment = m_Alignment
PropertyChanged "Alignment"
'UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
m_BackColor = NewBackColor
txtRaiz.BackColor = NewBackColor
pB.BackColor = NewBackColor
PropertyChanged "BackColor"
UserControl_Resize
End Property

Public Property Get BackColorOnFocus() As OLE_COLOR
BackColorOnFocus = m_FocusColor
End Property

Public Property Let BackColorOnFocus(ByVal NewColor As OLE_COLOR)
m_FocusColor = NewColor
PropertyChanged "BackColorOnFocus"
UserControl_Resize
End Property

Public Property Get BorderColor() As OLE_COLOR
BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
m_BorderColor = NewBorderColor
PropertyChanged "BorderColor"
DrawBorders m_BorderColor
End Property

Public Property Get BorderColorOnFocus() As OLE_COLOR
  BorderColorOnFocus = m_BorderOnFocus
End Property

Public Property Let BorderColorOnFocus(ByVal NewBorderColorOnFocus As OLE_COLOR)
  m_BorderOnFocus = NewBorderColorOnFocus
  PropertyChanged "BorderColorOnFocus"
End Property

Public Property Get CaseText() As CaseType
    CaseText = m_CaseText
End Property

Public Property Let CaseText(ByVal New_CaseText As CaseType)
    m_CaseText = New_CaseText
    PropertyChanged "CaseText"
End Property

Public Property Get CueText() As String
  CueText = m_CueText
End Property

Public Property Let CueText(ByVal NewCueText As String)
  m_CueText = NewCueText
  PropertyChanged "CueText"
End Property

Public Property Get CueTextColor() As OLE_COLOR
  CueTextColor = m_CueTextColor
End Property

Public Property Let CueTextColor(ByVal NewCueTextColor As OLE_COLOR)
  m_CueTextColor = NewCueTextColor
  PropertyChanged "CueTextColor"
If Trim$(m_CueText) <> "" And Trim$(m_Text) = "" Then
  txtRaiz.ForeColor = m_CueTextColor
End If
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  DrawBorders m_BorderColor
End Property

Public Property Get Enabled() As Boolean
Enabled = txtRaiz.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
txtRaiz.Enabled = New_Enabled
PropertyChanged "Enabled"
End Property

Public Property Get EnterKeyBehavior() As eEnterKeyBehavior
EnterKeyBehavior = m_KeyBehavior
End Property

Public Property Let EnterKeyBehavior(ByVal NewBehavior As eEnterKeyBehavior)
m_KeyBehavior = NewBehavior
PropertyChanged "EnterKeyBehavior"
End Property

Public Property Get Font() As Font
Set Font = txtRaiz.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Set txtRaiz.Font = New_Font
PropertyChanged "Font"
UserControl_Resize
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = txtRaiz.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
m_ForeColor = New_ForeColor
txtRaiz.ForeColor = New_ForeColor
PropertyChanged "ForeColor"
End Property

Public Property Get FormatToString() As CharacterType
    FormatToString = m_FormatToString
End Property

Public Property Let FormatToString(ByVal New_FormatToString As CharacterType)
    m_FormatToString = New_FormatToString
    txtRaiz_LostFocus
    PropertyChanged "FormatToString"
End Property

Public Property Get SetText() As String
    SetText = m_SetText
End Property

Public Property Let SetText(ByVal New_SetText As String)
    txtRaiz.Text = New_SetText
    Call txtRaiz_GotFocus
    Call txtRaiz_LostFocus
    PropertyChanged "SetText"
End Property

Public Property Get Locked() As Boolean
Locked = txtRaiz.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
txtRaiz.Locked = New_Locked
PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Single
MaxLength = txtRaiz.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Single)
txtRaiz.MaxLength = New_MaxLength
PropertyChanged "MaxLength"
End Property

Public Property Get PasswordChar() As String
PasswordChar = txtRaiz.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_Char As String)
txtRaiz.PasswordChar = New_Char
PropertyChanged "PasswordChar"
End Property

Public Property Let SelLength(ByVal iLength As Integer)
txtRaiz.SelLength = iLength
End Property

Public Property Let SelStart(ByVal iStart As Integer)
txtRaiz.SelStart = iStart
End Property

Public Property Get SelText() As String
SelText = txtRaiz.SelText
End Property

Public Property Get SelTextOnFocus() As Boolean
SelTextOnFocus = m_SelTextFocus
End Property

Public Property Let SelTextOnFocus(ByVal bSel As Boolean)
m_SelTextFocus = bSel
PropertyChanged "SelTextOnFocus"
End Property

Public Property Let SetFocus(ByVal Focus As Boolean)
HaveFocus = Focus
If Focus Then txtRaiz.SetFocus
End Property

Public Property Get Text() As String
Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
m_Text = New_Text
txtRaiz.Text = New_Text
PropertyChanged "Text"
End Property

Public Function Value() As Variant
Dim sValue As Variant, i As Integer
Dim sValor As String

sValor = txtRaiz.Text

On Error GoTo ErrVal

Select Case m_FormatToString
         
  Case Dates
ErrDATE:
     sValue = DateSerial(Year(sValor), Month(sValor), Day(sValor))
     
  Case ChileanRUT
    'Limpio Guión y Letras
     For i = 1 To Len(sValor)
        If IsNumeric(Mid$(sValor, i, 1)) = True Then
           sValue = sValue & Mid$(sValor, i, 1)
        ElseIf Mid$(sValor, i, 1) = "-" Then
          Exit For
        End If
     Next i
  
  Case AllChars, LettersOnly, LettersAndNumbers
    'Cuento caracteres
    sValue = Len(txtRaiz.Text)
    
  Case Else
    'Limpio Simbolo moneda y Puntos
     For i = 1 To Len(sValor)
        If IsNumeric(Mid$(sValor, i, 1)) Or Mid$(sValor, i, 1) = sDecimal Then
           sValue = sValue & Mid$(sValor, i, 1)
        End If
     Next i
    
End Select
  
  Value = sValue

ErrVal:
If Err.Number = 13 Then
  Value = "Error Date Value!"
Else
  Value = 0
End If
End Function

'- AllChars          - LettersOnly - NumbersOnly
'- LettersAndNumbers - Money       - Percent     - Fraction
'- Decimals          - Dates       - ChileanRUT

Private Sub GDIpRoundRect(ByVal hdc As Long, RECT As RECTF, ByVal BackColor, ByVal BorderColor As Long, Round As Long)
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim hGraphics As Long
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipCreateSolidFill BackColor, hBrush
    GdipCreatePen1 BorderColor, &H1 * nScale, &H2, hPen

     GdipCreatePath &H0, mPath   '&H0
        With RECT
            GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
            GdipAddPathArcI mPath, .Left + .Width - Round, .Top, Round, Round, 270, 90
            GdipAddPathArcI mPath, .Left + .Width - Round, .Top + .Height - Round, Round, Round, 0, 90
            GdipAddPathArcI mPath, .Left, .Top + .Height - Round, Round, Round, 90, 90
            GdipClosePathFigures mPath
        End With
        GdipFillPath hGraphics, hBrush, mPath
        GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
    
    GdipDeleteGraphics hGraphics
    
End Sub

Public Function RGBtoARGB(ByVal RGBColor As Long, Optional ByVal Opacity As Long = 100) As Long
    'By LaVople
    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format
    ' Passing VB system color constants is allowed, i.e., vbButtonFace
    ' Pass Opacity as a value from 0 to 255

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

'Inicia GDI+
Private Sub InitGDI()
    Dim GdipStartupInput As GdiplusStartupInput
    GdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

'' Ordinal #1
'Private Sub WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, _
'       ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
'       ByVal lParam As Long, ByRef lParamUser As Long)
'
'    Select Case uMsg
'        Case WM_MOUSELEAVE
'
'        Case WM_MOUSEHOVER
'
'        Case WM_MOUSEMOVE
'
'    End Select
'End Sub

