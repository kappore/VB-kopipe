'====================================================================='
'                             KERNEL32                                '
'====================================================================='
Declare Sub Sleep Lib "kernel32.dll" (ByVal ms As Long)

'====================================================================='
'                              USER32                                 '
'====================================================================='
Declare Function GetActiveWindow Lib "USER32" () As Long
Declare Function GetForegroundWindow Lib "USER32" () As Long
Declare Function GetClassNameA Lib "USER32" (ByVal hwnd As Long, _
        ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowTextA Lib "USER32" (ByVal hwnd As Long, _
        ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function EnumChildWindows Lib "USER32" (ByVal hwndParent As Long, _
        ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function FindWindowA Lib "USER32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
Declare Function FindWindowExA Lib "USER32" (ByVal hwndParent As Long, _
        ByVal hwndChildAfter As Long, ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
Declare Function FindWindowExW Lib "USER32" (ByVal hwndParent As Long, _
        ByVal hwndChildAfter As Long, ByRef lpClassName As String, _
        ByRef lpWindowName As String) As Long
Declare Function SendMessageA Lib "USER32" (ByVal hwnd As Long, _
        ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function PostMessageA Lib "USER32" (ByVal hwnd As Long, _
        ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As String) As Long



'====================================================================='
'                      get_child_window_handle                        '
'====================================================================='
Function mycallback_EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

    Dim res1 As Long
    Dim res2 As Long
    Dim classname As String * 512
    Dim windowtext As String * 512
    Dim base As Range
    res1 = GetClassNameA(hwnd, classname, 512)
    res2 = GetWindowTextA(hwnd, windowtext, 512)

    Set base = ActiveCell
    base.Offset(0, 0).Value = hwnd
    base.Offset(0, 1).Value = classname
    base.Offset(0, 2).Value = windowtext
    base.Offset(1, 0).Activate
        
    mycallback_EnumChildProc = 1

End Function

Sub get_child_window_handle()

    Dim hwnd As Long
    Dim classname As String * 512
    Dim base As Range
    Set base = ActiveCell
    
    Sleep 2000
    hwnd = GetForegroundWindow()
    
    base.Offset(0, 0).Value = hwnd
    base.Offset(1, 0).Activate
    
    EnumChildWindows hwnd, AddressOf mycallback_EnumChildProc, 0
    
End Sub

