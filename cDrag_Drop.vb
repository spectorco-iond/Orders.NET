Option Strict Off
Option Explicit On
Friend Class CDrag_Drop

    '	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '	'       THE DRAG CLASS
    '	'       ~~~~~~~~~~~~~~

    '	'You can use this class in any of your projects, in anyform
    '	'you like, modify it according to your needs, but give credit
    '	'where credit is due.
    '	'                               -Author:    Muhammad Abubakar
    '	'                                       <joehacker@yahoo.com>
    '	'                                       http://go.to/abubakar

    '	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



    Public Event FilesDroped()
    Private m_DragHwnd As Integer
    Private m_FileCount As Short
    Private FileNames() As String
    Private Working As Boolean
    Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal Hwnd As Integer, ByVal fAccept As Integer)

    '	Friend Sub AddInFileNames(ByRef Buffer As String)
    '		ReDim Preserve FileNames(m_FileCount)
    '		FileNames(m_FileCount) = Buffer
    '		m_FileCount = m_FileCount + 1
    '		'Debug.Print "file recieved : " & Buffer

    '	End Sub

    '	Friend Sub NowRaiseEvent()
    '		RaiseEvent FilesDroped()
    '	End Sub

    '	Friend Sub ClearFileNames()
    '		ReDim FileNames(0)
    '		m_FileCount = 0

    '	End Sub
    '	Public Function StartDrag() As Integer
    '		'This will start monitoring for the message of WM_DROPFILES
    '		'If already working then we wont subclass again
    '		If Working = False Then
    '			If m_DragHwnd > 0 Then
    '				DragAcceptFiles(m_DragHwnd, True)
    '				'Set obj = Me

    '                'UPGRADE_WARNING: Add a delegate for AddressOf WndProc Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'

    '                PrevWndFunc = SetWindowLong(m_DragHwnd, GWL_WNDPROC, AddressOf WndProc) ' SubClassProc)
    '                'SetWindowLong(m_DragHwnd, GWL_WNDPROC, AddressOf WndProc)
    '				StartDrag = 1 'Successfully started
    '				Working = True

    '			Else
    '				StartDrag = 0 'Unsuccessful, handle not given

    '			End If
    '		Else
    '			StartDrag = 2
    '		End If

    '	End Function

    '    'Public Function DelegateWindowLongSub(ByRef m_DragHwnd As Integer, ByRef GWL_WNDPROC As Long)

    '    '    PrevWndFunc = SetWindowLong(m_DragHwnd, GWL_WNDPROC, AddressOf WndProc)

    '    'End Function

    Public Property DragHwnd() As Integer
        Get
            DragHwnd = m_DragHwnd

        End Get
        Set(ByVal Value As Integer)

            If Not Working Then m_DragHwnd = Value

        End Set
    End Property

    '	Public ReadOnly Property FileCount() As Short
    '		Get
    '			FileCount = m_FileCount

    '		End Get
    '	End Property
    '	Public Function StopDrag() As Integer
    '		'Stop subclassing and monitoring of WM_DROPFILES message.

    '		If Working = True Then
    '			SetWindowLong(m_DragHwnd, GWL_WNDPROC, PrevWndFunc)
    '			DragAcceptFiles(m_DragHwnd, False)
    '			Working = False
    '			StopDrag = 1 'successfully stoped subclassing

    '		Else
    '			StopDrag = 0 'It was already not subclassed so no need to unsubclass

    '		End If
    '	End Function
    Public Function FileName(ByRef Index As Short) As String
        If Index >= 0 And Index <= m_FileCount Then
            FileName = FileNames(Index)
        Else
            FileName = ""
        End If

    End Function

    '	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '	Private Sub Class_Initialize_Renamed()
    '		m_DragHwnd = 0
    '		m_FileCount = 0
    '		'obj is declared in BAS- <CDrag_Drop_Module> of type CDrag_Drop
    '		obj = Me

    '	End Sub
    '	Public Sub New()
    '		MyBase.New()
    '		Class_Initialize_Renamed()
    '	End Sub

    '	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '	Private Sub Class_Terminate_Renamed()
    '		If Working = True Then StopDrag()

    '	End Sub
    '	Protected Overrides Sub Finalize()
    '		Class_Terminate_Renamed()
    '		MyBase.Finalize()
    '	End Sub
End Class