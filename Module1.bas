Attribute VB_Name = "Module1"
Sub lnkchg()
    Dim folderPath As String
    Dim oldServer As String
    Dim newServer As String
    Dim fso As Object
    Dim folder As Object
    Dim dialog As FileDialog
    Dim ws As Worksheet
    Dim fileCount As Long ' 変換されたファイル数をカウントする変数
    
    ' サーバー名を取得するシートを指定
    Set ws = ThisWorkbook.Sheets("Sheet1") ' シート名を適切に変更

    ' 変更前のサーバー名またはIPアドレスをシートから取得
    oldServer = ws.Range("B1").Value ' 変更前のサーバー名が入力されているセルを指定

    ' 変更後のサーバー名またはIPアドレスをシートから取得
    newServer = ws.Range("B2").Value ' 変更後のサーバー名が入力されているセルを指定

    ' 変更前後のサーバー名が未入力なら処理を中止
    If Trim(oldServer) = "" Or Trim(newServer) = "" Then
        MsgBox "変更前または変更後のサーバー名が入力されていません。処理を中止します。", vbExclamation
        Exit Sub
    End If

    ' フォルダ選択ダイアログを表示
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With dialog
        .Title = "フォルダを選択してください"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) ' 選択されたフォルダパスを取得
        Else
            MsgBox "フォルダが選択されていません。処理を中止します。"
            Exit Sub
        End If
    End With

    ' FileSystemObjectを作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダを取得
    Set folder = fso.GetFolder(folderPath)
    
    ' フォルダ内のショートカットを再帰的に変更（カウント付き）
    fileCount = 0 ' カウント初期化
    ProcessFolder folder, oldServer, newServer, fileCount, fso
    
    ' 変換されたファイル数を表示
    MsgBox fileCount & " ファイルのリンク先を変更しました。"
End Sub

Sub ProcessFolder(ByVal folder As Object, ByVal oldServer As String, ByVal newServer As String, ByRef fileCount As Long, ByRef fso As Object)
    Dim wshShell As Object
    Dim shortcut As Object
    Dim file As Object
    Dim subFolder As Object
    Dim backupFilePath As String

    ' シェルオブジェクトを作成
    Set wshShell = CreateObject("WScript.Shell")
    
    ' フォルダ内のすべてのファイルを処理
    For Each file In folder.Files
        ' .lnkファイルをチェック
        If LCase(Right(file.Name, 4)) = ".lnk" Then
            ' ショートカットをロード
            Set shortcut = wshShell.CreateShortcut(file.Path)
            
            ' リンク先がoldServerで始まる場合のみ変更
            If Left(shortcut.TargetPath, Len(oldServer)) = oldServer Then
                ' 変更前のファイルをバックアップとしてコピー
                backupFilePath = fso.BuildPath(folder.Path, fso.GetBaseName(file.Name) & "_old.lnk")
                fso.CopyFile file.Path, backupFilePath, True
                
                ' 新しいサーバー名に置き換え
                shortcut.TargetPath = Replace(shortcut.TargetPath, oldServer, newServer)
                shortcut.Save
                
                ' 変換されたファイルをカウント
                fileCount = fileCount + 1
            End If
        End If
    Next file
    
    ' サブフォルダ内のファイルも再帰的に処理
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder, oldServer, newServer, fileCount, fso
    Next subFolder
End Sub

