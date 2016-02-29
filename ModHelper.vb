Imports System.IO
Imports System.Runtime.InteropServices

Module ModHelpers
    Private Declare Auto Function FindWindow Lib "user32.dll" ( _
 ByVal lpClassName As String, _
 ByVal lpWindowName As String _
 ) As Long

    <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Function FindWindowByClass( _
     ByVal lpClassName As String, _
     ByVal zero As IntPtr) As IntPtr
    End Function

    Friend Function modFindWindow(ByVal classname As String) As Long
        Dim hwnd As Long
        Try
            hwnd = FindWindow(classname, vbNullString)
        Catch ex As Exception
            Debug.Print(ex.Message)
            Return Nothing
        End Try
        Return hwnd
    End Function
    Friend Sub LogOutPut(ByVal strLog As String, ByVal strAsFile As String)
        Try

            strLog = strLog & "  " & Environ("USERNAME") & "  " & DateTime.Today.ToLongDateString & " " & Now.ToLongTimeString

            If File.Exists(strAsFile) = True Then

                Dim sw As StreamWriter = File.AppendText(strAsFile)
                sw.WriteLine(strLog)
                sw.Flush()
                sw.Close()

            Else

                Dim sw As StreamWriter = File.CreateText(strAsFile)
                sw.WriteLine(strLog)
                sw.Flush()
                sw.Close()

            End If

        Catch ex As Exception
        End Try

    End Sub
End Module
