Option Strict Off
Option Explicit On
Module cDrag_Drop_Module

    Public Delegate Function SubClassProcDelegate(ByVal hwnd As Integer, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    '    Public Declare Function SetWindowLong Lib "USER32.DLL" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal attr As Integer, ByVal lVal As SubClassProcDelegate) As Integer
    Public Declare Function SetWindowLong Lib "USER32.DLL" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal attr As Integer, ByVal lVal As SubClassProcDelegate) As Integer

    'Public Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal Hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Integer, ByVal nIndex As Integer) As Integer
    Public Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Integer)
    Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Integer, ByVal UINT As Integer, ByVal lpStr As String, ByVal ch As Integer) As Integer
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Integer, ByVal Hwnd As Integer, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Const GWL_WNDPROC As Short = (-4)
    Public Const WM_DROPFILES As Integer = &H233
    Public PrevWndFunc As Integer ' SubClassProcDelegate ' Integer

    Public obj As CDrag_Drop

    '	Public Function WndProc(ByVal Hwnd As Integer, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    '		Dim iLoop, n, FileInfo As Integer
    '		Dim Buffer As New VB6.FixedLengthString(256)
    '        'Dim tmp As String
    '		Dim length As Integer
    '		'buffer = Space(256)
    '		If msg = WM_DROPFILES Then
    '			obj.ClearFileNames()
    '			FileInfo = wParam
    '			n = DragQueryFile(FileInfo, -1, vbNullString, 0)
    '			For iLoop = 0 To n - 1
    '				length = DragQueryFile(FileInfo, iLoop, Buffer.Value, 256)
    '				Buffer.Value = Trim(Buffer.Value)
    '				obj.AddInFileNames(Buffer.Value)
    '			Next 

    '			obj.NowRaiseEvent()

    '			DragFinish(FileInfo) 'wParam
    '			WndProc = 0
    '		Else
    '			WndProc = CallWindowProc(PrevWndFunc, Hwnd, msg, wParam, lParam)
    '		End If


    '	End Function
End Module