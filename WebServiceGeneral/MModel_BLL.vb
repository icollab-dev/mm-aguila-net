Imports System.Runtime.InteropServices
Imports System.Web.Services

Public Class MModel_BLL
    Inherits WebService

    Public Const SleepTime As Integer = 10
    Public Const WaitTimeOut As Integer = 10
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function FindWindowEx(ByVal parentHandle As IntPtr, ByVal childAfter As IntPtr, ByVal lclassName As String, ByVal windowTitle As String) As IntPtr

    End Function
    <DllImport("user32.dll")>
    Public Shared Function GetDlgItem(ByVal hDlg As IntPtr, ByVal nIDDlgItem As Integer) As IntPtr

    End Function
    <DllImport("user32.dll")>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As UInteger, ByVal lParam As Integer) As Integer

    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr,
    <Out> ByVal lParam As StringBuilder) As IntPtr

    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As Integer, ByVal lParam As StringBuilder) As IntPtr

    End Function

    Public Shared Sub Doclick(ByVal PROPOBJ As IntPtr)
        SendMessage(PROPOBJ, 513UI, 0UI, 0)
        SendMessage(PROPOBJ, 514UI, 0UI, 0)
    End Sub

    'telerik
    Public Shared Function GetWindowTextRaw(ByVal hwnd As IntPtr) As String
        Dim lParam As StringBuilder = New StringBuilder(SendMessage(hwnd, 14UI, 0UI, 0) + 1)
        SendMessage(hwnd, 13UI, CType(lParam.Capacity, IntPtr), lParam)
        Return lParam.ToString()
    End Function


    Public Shared Function SetWindowTextRaw(ByVal hwnd As IntPtr, ByVal text As String)
        Return SendMessage(hwnd, 12UI, 0, New StringBuilder(text))
    End Function
End Class

