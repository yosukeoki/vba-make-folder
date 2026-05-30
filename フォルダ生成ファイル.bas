Attribute VB_Name = "Module1"
Option Explicit

Sub フォルダ生成3()
    ' 変数定義
    Dim wb1 As Workbook
    Dim ws11 As Worksheet
    Dim dir1 As String, Foldername1 As String
    ' Dim dir2 As String, Foldername2 As String
    Dim i As Long
    ' 処理内容
    ' 起点のDirectoryを取得
    dir1 = ThisWorkbook.Path & "\"
    ' ブック指定
    Set wb1 = ThisWorkbook
    ' MsgBox wb1.Name
    ' シート指定
    Set ws11 = wb1.Worksheets(1)
    ' MsgBox ws11.Name
    ' フォルダ名を取得して有無確認の準備
    ' 1回目開始
    For i = 1 To 10
        ' ファイル名を取得
        Foldername1 = ws11.Cells(i + 4, 4).Value
        ' 空欄を見つけたら処理終了(歯抜けのところも処理終了の対象)
        If Foldername1 = "" Then
            Exit For
        End If
        ' check and create
        Call check_and_create(dir1, Foldername1)
        ' 入力引数：Dir1(=funcdir1), Foldername1(=funcfn1)
        ' 出力戻値：無
        ' 中の変数：Foldername1,Dir2,
    Next
    
    MsgBox "Fin"
End Sub

Sub check_and_create(ByVal funcdir1 As String, ByVal funcfn1 As String)
    Dim Dir2 As String, Foldername2 As String
    MsgBox "フォルダ確認"
    Dir2 = funcdir1 & funcfn1
    Foldername2 = Dir(Dir2, vbDirectory)
    MsgBox Dir2
    If Not Foldername2 <> "" Then
        MsgBox "フォルダなし"
        MkDir (Dir2)
        MsgBox "フォルダ作成した" & vbCrLf & "終了"
    Else
        MsgBox "フォルダあり"
        ' 有れば末尾(1)のフォルダ作成、(3)まである
        Call confirm_and_create_folder(funcdir1, funcfn1)
    End If
End Sub


Sub フォルダ生成2()
    ' 変数定義
    Dim wb1 As Workbook
    Dim ws11 As Worksheet
    Dim dir1 As String, Dir2 As String, Foldername1 As String, Foldername2 As String
    Dim i As Long
    ' 処理内容
    ' 起点のDirectoryを取得
    dir1 = ThisWorkbook.Path & "\"
    ' ブック指定
    Set wb1 = ThisWorkbook
    ' MsgBox wb1.Name
    ' シート指定
    Set ws11 = wb1.Worksheets(1)
    ' MsgBox ws11.Name
    ' フォルダ名を取得して有無確認の準備
    ' 1回目開始
    For i = 1 To 10
        Foldername1 = ws11.Cells(i + 4, 4).Value
        ' 空欄を見つけたら処理終了(歯抜けのところも処理終了の対象)
        If Foldername1 = "" Then
            Exit For
        End If
        MsgBox "フォルダ確認"
        Dir2 = dir1 & Foldername1
        Foldername2 = Dir(Dir2, vbDirectory)
        MsgBox Dir2
        If Not Foldername2 <> "" Then
            MsgBox "フォルダなし"
            MkDir (Dir2)
            MsgBox "フォルダ作成した" & vbCrLf & "終了"
        Else
            MsgBox "フォルダあり"
            ' 有れば末尾(1)のフォルダ作成、(3)まである
            Call confirm_and_create_folder(dir1, Foldername1)
        End If
    Next
    
    MsgBox "Fin"
End Sub

Sub フォルダ生成1()
    ' 変数定義
    Dim wb1 As Workbook
    Dim ws11 As Worksheet
    Dim dir1 As String, Dir2 As String, Foldername1 As String, Foldername2 As String
    Dim i As Long
    ' 処理内容
    ' ブック指定
    Set wb1 = ThisWorkbook
    ' MsgBox wb1.Name
    ' シート指定
    Set ws11 = wb1.Worksheets(1)
    ' MsgBox ws11.Name
    Foldername1 = ws11.Cells(5, 4).Value
    dir1 = ThisWorkbook.Path & "\"
    ' MsgBox Foldername2
    ' カウントしながら処理
    MsgBox "フォルダ確認"
    Dir2 = dir1 & Foldername1
    Foldername2 = Dir(Dir2, vbDirectory)
    MsgBox Dir2
    If Not Foldername2 <> "" Then
        MsgBox "フォルダなし"
        MkDir (Dir2)
        MsgBox "フォルダ作成した" & vbCrLf & "終了"
    Else
        MsgBox "フォルダあり"
        ' 有れば末尾(1)のフォルダ作成、(5)まである
        Call confirm_and_create_folder(dir1, Foldername1)
    End If
    
    MsgBox "Fin"
End Sub

' 有れば末尾(1)の複製フォルダ作成、(3)まで
Function confirm_and_create_folder(ByVal parentPath As String, ByVal baseName As String) As String
    Dim Dir2 As String, Foldername2 As String
    Dim i As Long
    ' i = 1
    For i = 1 To 3
        Dir2 = parentPath & baseName & "(" & i & ")"
        Foldername2 = Dir(Dir2, vbDirectory)
        MsgBox "末尾(" & i & ")フォルダ確認" & vbCrLf & Dir2
        If Not Foldername2 <> "" Then
            MsgBox "フォルダなし"
            MkDir (Dir2)
            MsgBox "フォルダ作成した" & vbCrLf & "終了"
            Exit For
        Else
            If i <> 3 Then
                MsgBox "フォルダあり"
            Else
                MsgBox "(3)までフォルダあり、フォルダを削除して下さい"
                Exit For
            End If
        End If
    Next
End Function
