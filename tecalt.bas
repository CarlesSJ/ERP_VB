Attribute VB_Name = "Module3"
'------------------------------------------------------------------------------
' Para bloquear algunas teclas en Windows NT/2000/XP                (08/Mar/03)
' Para NT debe tener el SP3 como mínimo
'
' ¡¡¡ NO FUNCIONA para Ctrl+Alt+Supr !!!
'
' En este ejemplo se bloquean las siguientes teclas:
'   Ctrl+Esc, Alt+Tab y Alt+Esc
'
' ©Guillermo 'guille' Som, 2003
'------------------------------------------------------------------------------
Option Explicit

' para guardar el gancho creado con SetWindowsHookEx
Private mHook As Long

'
' para indicar a SetWindowsHookEx que tipo de gancho queremos instalar
Private Const WH_KEYBOARD_LL As Long = 13&
' este es para el ratón
'Private Const WH_MOUSE_LL As Long = 14&
'
Private Type tagKBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
'
Private Const VK_TAB As Long = &H9
Private Const VK_CONTROL As Long = &H11     ' tecla Ctrl
'Private Const VK_MENU As Long = &H12        ' tecla Alt
Private Const VK_ESCAPE As Long = &H1B
'Private Const VK_DELETE As Long = &H2E      ' tecla Supr (Del)
'
Private Const LLKHF_ALTDOWN As Long = &H20&
'
' códigos para los ganchos (la acción a tomar en el gancho del teclado)
Private Const HC_ACTION As Long = 0&


'-----------------------------
' Funciones del API de Windows
'-----------------------------

' para asignar un gancho (hook)
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
    (ByVal idHook As Long, ByVal lpfn As Long, _
    ByVal hMod As Long, ByVal dwThreadId As Long) As Long

' para quitar el gancho creado con SetWindowsHookEx
Private Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Long) As Long
    
' para llamar al siguiente gancho
Private Declare Function CallNextHookEx Lib "user32" _
    (ByVal hHook As Long, ByVal nCode As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

' para saber si se ha pulsado en una tecla
Private Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Integer

' para copiar la estructura en un long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)



' La función a usar para el gancho del teclado
Public Function LLKeyBoardProc(ByVal nCode As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long _
                               ) As Long
    Dim pkbhs As tagKBDLLHOOKSTRUCT
    Dim ret As Long
    '
    ret = 0
    '
    ' copiar el parámetro en la estructura
    CopyMemory pkbhs, ByVal lParam, Len(pkbhs)
    '
    If nCode = HC_ACTION Then
        '
        ' si se pulsa Ctrl+Esc
        tecla = pkbhs.vkCode
        If pkbhs.vkCode = VK_ESCAPE Then
            If (GetAsyncKeyState(VK_CONTROL) And &H8000) Then
                ret = 1
            End If
        End If
        '
        ' si se pulsa Alt+Tab
        If pkbhs.vkCode = VK_TAB Then
            If (pkbhs.flags And LLKHF_ALTDOWN) <> 0 Then
                ret = 1
            End If
        End If
        '
        ' si se pulsa Alt+Esc
        If pkbhs.vkCode = VK_ESCAPE Then
            If (pkbhs.flags And LLKHF_ALTDOWN) <> 0 Then
                ret = 1
            End If
        End If
        '
    End If
    '
    If ret = 0 Then
        ret = CallNextHookEx(mHook, nCode, wParam, lParam)
    End If
    '
    LLKeyBoardProc = ret
    '
    '

    ' El código C++ en el que he basado (o casi) el de VB

'LRESULT CALLBACK LowLevelKeyboardProc (INT nCode, WPARAM wParam, LPARAM lParam)
'{
'    // By returning a non-zero value from the hook procedure, the
'    // message does not get passed to the target window
'    KBDLLHOOKSTRUCT *pkbhs = (KBDLLHOOKSTRUCT *) lParam;
'    BOOL bControlKeyDown = 0;
'
'    switch (nCode)
'    {
'        case HC_ACTION:
'        {
'            // Check to see if the CTRL key is pressed
'            bControlKeyDown = GetAsyncKeyState (VK_CONTROL) >> ((sizeof(SHORT) * 8) - 1);
'
'            // Disable CTRL+ESC
'            if (pkbhs->vkCode == VK_ESCAPE && bControlKeyDown)
'                return 1;
'
'            // Disable ALT+TAB
'            if (pkbhs->vkCode == VK_TAB && pkbhs->flags & LLKHF_ALTDOWN)
'                return 1;
'
'            // Disable ALT+ESC
'            if (pkbhs->vkCode == VK_ESCAPE && pkbhs->flags & LLKHF_ALTDOWN)
'                return 1;
'
'            break;
'        }
'
'default:
'            break;
'    } return CallNextHookEx (hHook, nCode, wParam, lParam);
'}

End Function



Public Sub HookKeyB(ByVal hMod As Long)
    ' instalar el gancho para el teclado
    ' hMod será el valor de App.hInstance de la aplicación
    mHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LLKeyBoardProc, hMod, 0&)
End Sub

Public Sub UnHookKeyB()
    ' desinstalar el gancho para el teclado
    ' Es importante hacerlo antes de finalizar la aplicación,
    ' normalmente en el evento Unload o QueryUnload
    If mHook <> 0 Then
        UnhookWindowsHookEx mHook
    End If
End Sub

Public Sub comença_captura()
    ' iniciar el gancho para el teclado
    HookKeyB App.hInstance
End Sub

Private Sub acava_captura()
    ' quitar el gancho del teclado
    UnHookKeyB
End Sub

