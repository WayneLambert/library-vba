'Source: https://wellsr.com/vba/2018/excel/vba-fade-userform-in-and-out/

'PLACE IN YOUR USERFORM CODE
Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long                
Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "USER32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'Constants for title bar
Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20)         'The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar
Private Const WS_EX_DLGMODALFRAME As Long = &H1   'Controls if the window has an icon
 
'Constants for transparency
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1                  'Chroma key for fading a certain color on your Form
Private Const LWA_ALPHA = &H2                     'Only needed if you want to fade the entire userform

'sleep
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64-Bit versions of Excel
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32-Bit versions of Excel
#End If
Dim formhandle As Long

Private Sub UserForm_Initialize()
    'force the form to fully transparent before it even loads
    formhandle = FindWindow(vbNullString, Me.Caption)
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetOpacity (0)
End Sub

'***********************************************************************************************************************

Private Sub UserForm_Activate()
    'HideTitleBarAndBorder Me 'hide the titlebar and border
    FadeUserform Me, True 'Fade your userform in
End Sub

'***********************************************************************************************************************

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    FadeUserform Me, False 'Fade your userform in
End Sub

'***********************************************************************************************************************

Sub FadeUserform(frm As Object, Optional FadeIn As Boolean = True)
    'Defaults to fade your userform in.
    'Set the 2nd argument to False to Fade Out.
    Dim iOpacity As Integer
    
    formhandle = FindWindow(vbNullString, Me.Caption)
    
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    'The following line sets the userform opacity equal to whatever value you have in iOpacity (0 to 255).
    If FadeIn = True Then 'fade in
        For iOpacity = 0 To 255 Step 15
            Call SetOpacity(iOpacity)
        Next
    Else 'fade out
        For iOpacity = 255 To 0 Step -15
            Call SetOpacity(iOpacity)
        Next
        Unload Me 'unload form once faded out
    End If
End Sub

'***********************************************************************************************************************

Sub SetOpacity(Opacity As Integer)
    SetLayeredWindowAttributes formhandle, Me.BackColor, Opacity, LWA_ALPHA
    Me.Repaint
    Sleep 50
End Sub

'***********************************************************************************************************************

''Source: https://wellsr.com/vba/2017/excel/remove-window-border-title-bar-around-userform-vba/
''Hides title bar and border around userform
Sub HideTitleBarAndBorder(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
'Build window and set window until you remove the caption, title bar and frame around the window
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl
End Sub