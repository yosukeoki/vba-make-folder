Attribute VB_Name = "Module1"
Option Explicit

Sub フォルダ生成()
    ' 変数定義
    Dim wb1 As Workbook
    Dim ws11 As Worksheet
    Dim Dir1 As String, Dir2 As String, Foldername1 As String, Foldername2 As String
    Dim i As Long
    ' 処理内容
    ' ブック指定
    Set wb1 = ThisWorkbook
    ' MsgBox wb1.Name
    ' シート指定
    Set ws11 = wb1.Worksheets(1)
    ' MsgBox ws11.Name
    Foldername1 = ws11.Cells(5, 3).Value
    Dir1 = ThisWorkbook.Path & "\"
    ' MsgBox Foldername2
    ' カウントしながら処理
    MsgBox "フォルダ確認"
    Dir2 = Dir1 & Foldername1
    Foldername2 = Dir(Dir2, vbDirectory)
    MsgBox Dir2
    If Not Foldername2 <> "" Then
        MsgBox "フォルダなし"
        MkDir (Dir2)
        MsgBox "フォルダ作成した" & vbCrLf & "終了"
    Else
        MsgBox "フォルダあり"
        ' 有れば末尾(1)のフォルダ作成、(5)まである
        Call confirm_and_create_folder(Dir1, Foldername1)
    End If
    
    MsgBox "Fin"
End Sub

' 有れば末尾(1)の複製フォルダ作成、(5)まで
Function confirm_and_create_folder(ByVal parentPath As String, ByVal baseName As String) As String
    Dim Dir2 As String, Foldername2 As String
    Dim i As Long
    i = 1
    For i = 1 To 5
        Dir2 = parentPath & baseName & "(" & i & ")"
        Foldername2 = Dir(Dir2, vbDirectory)
        MsgBox "末尾(" & i & ")フォルダ確認" & vbCrLf & Dir2
        If Not Foldername2 <> "" Then
            MsgBox "フォルダなし"
            MkDir (Dir2)
            MsgBox "フォルダ作成した" & vbCrLf & "終了"
            Exit For
        Else
            MsgBox "フォルダあり"
        End If
    Next
End Function
