Attribute VB_Name = "API"
Option Explicit

'AUTHOR: Ricardo Camisa
'Email: rich.7.2014@gmail.com

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const GWL_STYLE As Long = (-16)
Private Const WS_CAPTION = 55000000
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const IDC_HAND = 32649&
Public MeuForm As Long
Public ESTILO As Long
Public Const ESTILO_ATUAL As Long = (-16)
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Colbtn                   As New Collection
Public cbtStyle                 As NavBar
Public useStateWidth            As Boolean
Public useStateToggle            As Boolean

'-----------------------------------------------------------------------------------------------------------
#If VBA7 Then
    'Author: Ricardo Camisa
    'Api que permite fazermos um mover o formulário e darmos o release ---------------------------------
    Public Declare PtrSafe Sub ReleaseCapture Lib "User32" ()
    Public Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            lParam As Any) As Long
    'Api para determinar o theme atual
    Public Declare PtrSafe Function GetClassName Lib "User32" _
        Alias "GetClassNameA" (ByVal hWnd As Long, _
        ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Public Declare PtrSafe Function SetSysColors Lib "User32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
    Public Declare PtrSafe Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, _
                                                                                     Optional ByVal pszSubAppName As Long = 0, _
                                                                                     Optional ByVal pszSubIdList As Long = 0) As Long
    Public Declare PtrSafe Sub InitCommonControls Lib "comctl32" ()
    
    '----------------------------------------------------------------------------------------------------
        'API DO ICON
    Public Declare PtrSafe Function IconApp Lib "User32" Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
    Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function FindWindowA Lib "User32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function LoadCursorBynum Lib "User32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare PtrSafe Function SetCursor Lib "User32" (ByVal hCursor As Long) As Long
    Public Declare PtrSafe Function MoveJanela Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    'Author: Ricardo Camisa
    'Api que permite fazermos um mover o formulário e darmos o release ---------------------------------
    Public Declare Sub ReleaseCapture Lib "User32" ()
    Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            lParam As Any) As Long
    'Api para determinar o theme atual
    Public Declare Function GetClassName Lib "User32" _
        Alias "GetClassNameA" (ByVal hWnd As Long, _
        ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
        
    Public Declare Function SetSysColors Lib "User32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
    Public Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, _
                                                                                     Optional ByVal pszSubAppName As Long = 0, _
                                                                                     Optional ByVal pszSubIdList As Long = 0) As Long
    Public Declare Sub InitCommonControls Lib "comctl32" ()
    '----------------------------------------------------------------------------------------------------
        'API DO ICON
    Public Declare Function IconApp Lib "User32" Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
    Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long
    Public Declare Function FindWindowA Lib "User32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function LoadCursorBynum Lib "User32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare Function SetCursor Lib "User32" (ByVal hCursor As Long) As Long
    Public Declare Function MoveJanela Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
Private lngPixelsX As Long
Private lngPixelsY As Long
Private strThunder As String
Private blnCreate As Boolean
Private lnghWnd_Form As Long
Private lnghWnd_Sub As Long
Private colBaseCtrl As Collection
Private Const cstMask As Long = &H7FFFFFFF
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const FOCO_ICONE = &H80
Private Const ICONE = 0&
Private Const GRANDE_ICONE = 1&

'Constante que representa a mensagem
Private Const WM_SETICON = &H80
'Constante que representa o índice da cor activa da barra de títuto
Private Const ICON_SMALL = 0
'Constante que representa o tamanho do ícone a ser definido
Private Const COLOR_ACTIVECAPTION = 2

Private hIcone As Long


Public Function MouseCursor(CursorType As Long)
  Dim lngRet As Long
  lngRet = LoadCursorBynum(0&, CursorType)
  lngRet = SetCursor(lngRet)
End Function




Public Function GetCurrentThemeWin(ByVal form As Object)
    Dim Result As Integer
    Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long

    If Val(Application.Version) < 9 Then
        lngMyHandle = FindWindowA("ThunderXFrame", form.Caption)
    Else
        lngMyHandle = FindWindowA("ThunderDFrame", form.Caption)
    End If

    Call InitCommonControls



    Dim theme As String
    theme = "Aero"
    Dim className As String

    className = Space(256)
    Dim nameClass As Long

    nameClass = GetClassName(lngMyHandle, className, Len(className))
    className = Left(className, InStr(className, vbNullChar) - 1)

    Result = SetWindowTheme(lngMyHandle, nameClass, StrPtr(theme))
    MsgBox Result
End Function

'Metodo que permite movimentar o formulário
Public Sub moverForm(form As Object, obj As Object, Button As Integer)
    Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
    If Val(Application.Version) < 9 Then
        lngMyHandle = FindWindowA("ThunderXFrame", form.Caption)
    Else
        lngMyHandle = FindWindowA("ThunderDFrame", form.Caption)
    End If
    
    If Button = 1 Then
        With obj
            Call ReleaseCapture
            Call SendMessage(lngMyHandle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End With
    End If
End Sub

'No evento mover mouse da caixa de texto, digite:
'=MouseCursor(32649) => altera para formato de mão
Public Function Maozinha()
    Call MouseCursor(IDC_HAND)
End Function

Public Sub removeTudo(objForm As Object)
    MeuForm = FindWindowA(vbNullString, objForm.Caption)
    ESTILO = ESTILO Or WS_CAPTION
    MoveJanela MeuForm, ESTILO_ATUAL, (ESTILO)
    
End Sub

Public Sub removeCaption(objForm As Object, Optional color As Variant)
'    On Error Resume Next
    Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
    
    objForm.Caption = "My Portifolio"
    
    If Val(Application.Version) < 9 Then
        lngMyHandle = FindWindowA("ThunderXFrame", objForm.Caption)
    Else
        lngMyHandle = FindWindowA("ThunderDFrame", objForm.Caption)
    End If

    MeuForm = FindWindowA(vbNullString, objForm.Caption)
    ESTILO = ESTILO Or &HCCCCC0 '&HDC47BE '&HDCCEBE
    MoveJanela MeuForm, ESTILO_ATUAL, (ESTILO)
    
    lngCurrentStyle = GetWindowLong(lngMyHandle, GWL_STYLE)
    lngNewStyle = lngCurrentStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    SetWindowLong lngMyHandle, GWL_STYLE, lngNewStyle
    
    Dim ImageIcon As MSForms.Image
    Set ImageIcon = objForm.Controls.Add("Forms.Image.1", "ImageIcon", False)
    
    ImageIcon.Picture = LoadPicture(ThisWorkbook.Path & "\logo.ico")
    
    hIcone = ImageIcon.Picture.Handle
    Call IconApp(MeuForm, FOCO_ICONE, ICONE, ByVal hIcone)
    
    
End Sub

Sub MostrarExcel()
Application.Visible = True
Application.ScreenUpdating = True
End Sub

Public Sub Maocursor(myForm As Object)
    Dim ctrl As Control
    
    For Each ctrl In myForm.Controls
        If ctrl.Tag = "btn" And TypeName(ctrl) = "CommandButton" Then
            Set cbtStyle = New NavBar
            Set cbtStyle.btnButton = ctrl
            Colbtn.Add cbtStyle
        ElseIf ctrl.Tag = "btn" And TypeName(ctrl) = "Label" Then
            Set cbtStyle = New NavBar
            Set cbtStyle.btnLabel = ctrl
            Colbtn.Add cbtStyle
        ElseIf ctrl.Tag = "btn" And TypeName(ctrl) = "Image" Then
            Set cbtStyle = New NavBar
            Set cbtStyle.btnImage = ctrl
            Colbtn.Add cbtStyle
        ElseIf ctrl.Tag = "btn" And TypeName(ctrl) = "CheckBox" Then
            Set cbtStyle = New NavBar
            Set cbtStyle.btnCheckbox = ctrl
            Colbtn.Add cbtStyle
        ElseIf ctrl.Tag = "btn-social-media" And TypeName(ctrl) = "Label" Then
            Set cbtStyle = New NavBar
            Set cbtStyle.btnLabelSocialMedia = ctrl
            Colbtn.Add cbtStyle
        End If
    Next
End Sub

