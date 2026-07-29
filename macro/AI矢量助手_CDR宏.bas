Attribute VB_Name = "AIVectorHelper"
'================================================================
'  AIʸ������ ���� CorelDRAW �����������
'  ���� CorelDRAW X4 ~ 2024+ (32/64λ����Ӧ)
'
'  �ṩ 4 ���ɰ󶨵���������ť�ĺ�:
'    AI����_��    �򿪲�����, ���Զ������� CDR �����Ҳ�(����ģʽ,
'                   ʼ����ʾ�� CDR ֮��, ���϶�, �Ƽ��ճ�ʹ��)
'    AI����_ͣ��    �����"Ƕ��"CDR �������Ҳ�, ���Ʋ��봰(ʵ�鹦��)
'    AI����_����    ȡ��ͣ��, �ָ�Ϊ���϶����������
'    AI����_�ر�    �رղ�����
'
'  �� ��װ(ֻ��һ��):
'   1. CDR �˵�: ���� > �� > ��༭��(�� Visual Basic �༭��), ��ݼ� Alt+F11
'   2. ��๤����ѡ�� GlobalMacros �� �˵�[�ļ� > �����ļ�] �� ѡ�� .bas
'   3. ��������Ĭ��λ��, �޸��·� PLUGIN_PATH Ϊʵ��·��
'   4. Ctrl+S ���� GlobalMacros, �رձ༭��
'  �� �ӵ�������:
'   5. ���� > ѡ�� > �Զ��� > ����(�ɰ�: ���� > �Զ��� > �����б�),
'      ���Ͻ��������ѡ[��], �ҵ� AIVectorHelper.AI����_��,
'      ��ס�ϵ����⹤�����ϼ���; ��[���]ѡ������Ը���ť��ͼ�ꡣ
'      Ҳ������[��ݼ�]ѡ�������һ����ݼ�(�� Ctrl+Shift+A)��
'================================================================
Option Explicit

' ������ ���������·��(���б䶯���޸�) ������
Private Const PLUGIN_PATH As String = "C:\CDR���\AIʸ������\AIʸ������.hta"
' HTA ���ڱ���(���ڲ��Ҵ���, ������ .hta �� <title> ��ȫһ��)
Private Const HTA_TITLE As String = "AI ʸ������ v2 �� CorelDRAW ���"
' ������(����)
Private Const PANEL_W As Long = 480

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'---------------- Win32 API ����(32/64λ˫����) ----------------
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nW As Long, ByVal nH As Long, ByVal bRepaint As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function PostMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
    #If Win64 Then
        Private Declare PtrSafe Function SetWindowLongApi Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNew As LongPtr) As LongPtr
        Private Declare PtrSafe Function GetWindowLongApi Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    #Else
        Private Declare PtrSafe Function SetWindowLongApi Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNew As LongPtr) As LongPtr
        Private Declare PtrSafe Function GetWindowLongApi Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    #End If
#Else
    Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
    Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
    Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nW As Long, ByVal nH As Long, ByVal bRepaint As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function PostMessageW Lib "user32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Function SetWindowLongApi Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNew As Long) As Long
    Private Declare Function GetWindowLongApi Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const GWL_HWNDPARENT As Long = -8
Private Const WS_CHILD As Long = &H40000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_VISIBLE As Long = &H10000000
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SW_RESTORE As Long = 9
Private Const WM_CLOSE As Long = &H10

'================= ������ 1: ��(��������, �Ƽ�) =================
Public Sub AI����_��()
    On Error GoTo EH
    Dim hHta
    hHta = EnsureHtaRunning()
    If hHta = 0 Then Exit Sub
    ' ��Ϊ CDR ��"��������": ʼ�ո��� CDR �Ϸ�, �����ڵ������л�
    Dim hCdr
    hCdr = GetCdrHwnd()
    If hCdr <> 0 Then SetWindowLongApi hHta, GWL_HWNDPARENT, hCdr
    SnapToRight hHta, hCdr
    ShowWindow hHta, SW_RESTORE
    Exit Sub
EH:
    MsgBox "��ʧ��: " & Err.Description, vbCritical, "AIʸ������"
End Sub

'================= ������ 2: ͣ��(Ƕ��CDR����, ʵ��) =================
Public Sub AI����_ͣ��()
    On Error GoTo EH
    Dim hHta
    hHta = EnsureHtaRunning()
    If hHta = 0 Then Exit Sub
    Dim hCdr
    hCdr = GetCdrHwnd()
    If hCdr = 0 Then
        MsgBox "δ�ܻ�ȡ CorelDRAW �����ڡ�", vbExclamation, "AIʸ������"
        Exit Sub
    End If
    ' ȥ��������, ��Ϊ�Ӵ�����ʽ��Ƕ��
    Dim style
    style = GetWindowLongApi(hHta, GWL_STYLE)
    style = (style And (Not WS_POPUP) And (Not WS_CAPTION) And (Not WS_THICKFRAME)) Or WS_CHILD Or WS_VISIBLE
    SetWindowLongApi hHta, GWL_STYLE, style
    SetParent hHta, hCdr
    Dim rc As RECT
    GetClientRect hCdr, rc
    ' ռ�ݿͻ����Ҳ�, Ԥ�������������͵ײ�״̬���߶�
    MoveWindow hHta, rc.Right - PANEL_W, 110, PANEL_W, rc.Bottom - 110 - 30, 1
    SetWindowPos hHta, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_SHOWWINDOW Or &H2 Or &H1
    Exit Sub
EH:
    MsgBox "ͣ��ʧ��: " & Err.Description, vbCritical, "AIʸ������"
End Sub

'================= ������ 3: ����(ȡ��ͣ��) =================
Public Sub AI����_����()
    On Error GoTo EH
    Dim hHta
    hHta = FindHta()
    If hHta = 0 Then Exit Sub
    Dim style
    style = GetWindowLongApi(hHta, GWL_STYLE)
    style = (style And (Not WS_CHILD)) Or WS_POPUP Or WS_CAPTION Or WS_THICKFRAME Or WS_VISIBLE
    SetParent hHta, 0
    SetWindowLongApi hHta, GWL_STYLE, style
    Dim hCdr
    hCdr = GetCdrHwnd()
    If hCdr <> 0 Then SetWindowLongApi hHta, GWL_HWNDPARENT, hCdr
    SnapToRight hHta, hCdr
    Exit Sub
EH:
    MsgBox "����ʧ��: " & Err.Description, vbCritical, "AIʸ������"
End Sub

'================= ������ 4: �ر���� =================
Public Sub AI����_�ر�()
    Dim hHta
    hHta = FindHta()
    If hHta <> 0 Then PostMessageW hHta, WM_CLOSE, 0, 0
End Sub

'============ 进程内导入(供插件面板通过 RunMacro 调用) ============
' 绿色版/精简版 CDR 的进程外 COM Import 接口损坏, 由本函数在 CDR 进程内执行导入
Public Function AI_ImportFile(ByVal sPath As String) As Long
    On Error GoTo EH
    Dim d As Document
    Set d = ActiveDocument
    If d Is Nothing Then Set d = CreateDocument
    d.ActiveLayer.Import sPath
    AI_ImportFile = 1
    Exit Function
EH:
    AI_ImportFile = 0
End Function

'================================ �ڲ����� ================================

Private Function FindHta()
    FindHta = FindWindowW(0, StrPtr(HTA_TITLE))
End Function

Private Function EnsureHtaRunning()
    Dim h
    h = FindHta()
    If h <> 0 Then EnsureHtaRunning = h: Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(PLUGIN_PATH) Then
        MsgBox "δ�ҵ����������:" & vbCrLf & PLUGIN_PATH & vbCrLf & vbCrLf & _
               "��򿪺�༭��(Alt+F11), �޸� AIVectorHelper ģ���е� PLUGIN_PATH ������", _
               vbExclamation, "AIʸ������"
        EnsureHtaRunning = 0
        Exit Function
    End If

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run "mshta.exe """ & PLUGIN_PATH & """", 1, False

    ' �ȴ����ڳ���(��� 12 ��)
    Dim i As Long
    For i = 1 To 120
        Sleep 100
        h = FindHta()
        If h <> 0 Then Exit For
    Next i
    If h = 0 Then MsgBox "������δ������(���ܱ���ȫ���������� mshta.exe)��", vbExclamation, "AIʸ������"
    EnsureHtaRunning = h
End Function

Private Function GetCdrHwnd()
    ' ����ȡ CDR ����ģ�͵������ھ��(���ְ汾֧��), ʧ����ȡǰ̨����
    Dim h
    h = 0
    On Error Resume Next
    Dim o As Object
    Set o = Application
    h = o.AppWindow.Handle
    If h = 0 Then h = o.MainWindow.Handle
    On Error GoTo 0
    If h = 0 Then h = GetForegroundWindow()
    GetCdrHwnd = h
End Function

Private Sub SnapToRight(ByVal hHta, ByVal hCdr)
    Dim rc As RECT
    If hCdr <> 0 Then
        GetWindowRect hCdr, rc
    Else
        rc.Left = 100: rc.Top = 60: rc.Right = 1700: rc.Bottom = 1000
    End If
    Dim w As Long, h As Long
    w = PANEL_W
    h = rc.Bottom - rc.Top - 120
    If h < 500 Then h = 500
    SetWindowPos hHta, 0, rc.Right - w - 12, rc.Top + 90, w, h, SWP_NOZORDER Or SWP_SHOWWINDOW
End Sub
