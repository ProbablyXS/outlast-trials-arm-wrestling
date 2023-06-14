Imports System.Runtime.InteropServices

Module Var

    Public Const DWMWA_EXTENDED_FRAME_BOUNDS = 9&      ' methode precise

    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String,
    ByVal lpWindowName As String) As IntPtr

    <DllImport("user32.dll", EntryPoint:="GetWindowRect")>
    Public Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("dwmapi.dll")>      ' methode precise
    Public Function DwmGetWindowAttribute(ByVal hwnd As IntPtr, ByVal dwAttribute As Integer, <Out> ByRef pvAttribute As RECT, ByVal cbAttribute As Integer) As Integer
    End Function

    Public hWnd As IntPtr
    Public OL, RR, OT, OB,
        window_h, window_w As Integer
    Public bitmap As Bitmap
    Public g As Graphics

    Public Structure RECT

        Public left, top, right, bottom As Integer

    End Structure

End Module
