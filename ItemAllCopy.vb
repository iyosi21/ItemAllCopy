Imports System
Imports System.Windows.Forms

Module Program
    Sub Main()
        'コピーするファイルのパスをStringCollectionに追加する
        Dim Clip_Text As String = Clipboard.GetText()
        Dim path As String = System.Windows.Forms.Application.StartupPath
        Dim i As Integer
        Dim sFilePath As String = System.Windows.Forms.Application.ExecutablePath
        Dim files As New System.Collections.Specialized.StringCollection()
        Dim files_name As String() = System.IO.Directory.GetFiles(path, "*", System.IO.SearchOption.AllDirectories)

        For i = 0 To UBound(files_name)
            '実行ファイル自身はコピーしたくないので、その判定。
            If Not files_name(i) = sFilePath Then
                files.Add(files_name(i))
            End If
        Next i
        'クリップボードにコピーする
        Clipboard.SetFileDropList(files)


    End Sub
End Module
