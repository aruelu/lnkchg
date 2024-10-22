Attribute VB_Name = "Module1"
Sub lnkchg()
    Dim folderPath As String
    Dim oldServer As String
    Dim newServer As String
    Dim fso As Object
    Dim folder As Object
    Dim dialog As FileDialog
    Dim ws As Worksheet
    Dim fileCount As Long ' �ϊ����ꂽ�t�@�C�������J�E���g����ϐ�
    
    ' �T�[�o�[�����擾����V�[�g���w��
    Set ws = ThisWorkbook.Sheets("Sheet1") ' �V�[�g����K�؂ɕύX

    ' �ύX�O�̃T�[�o�[���܂���IP�A�h���X���V�[�g����擾
    oldServer = ws.Range("B1").Value ' �ύX�O�̃T�[�o�[�������͂���Ă���Z�����w��

    ' �ύX��̃T�[�o�[���܂���IP�A�h���X���V�[�g����擾
    newServer = ws.Range("B2").Value ' �ύX��̃T�[�o�[�������͂���Ă���Z�����w��

    ' �ύX�O��̃T�[�o�[���������͂Ȃ珈���𒆎~
    If Trim(oldServer) = "" Or Trim(newServer) = "" Then
        MsgBox "�ύX�O�܂��͕ύX��̃T�[�o�[�������͂���Ă��܂���B�����𒆎~���܂��B", vbExclamation
        Exit Sub
    End If

    ' �t�H���_�I���_�C�A���O��\��
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With dialog
        .Title = "�t�H���_��I�����Ă�������"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) ' �I�����ꂽ�t�H���_�p�X���擾
        Else
            MsgBox "�t�H���_���I������Ă��܂���B�����𒆎~���܂��B"
            Exit Sub
        End If
    End With

    ' FileSystemObject���쐬
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �t�H���_���擾
    Set folder = fso.GetFolder(folderPath)
    
    ' �t�H���_���̃V���[�g�J�b�g���ċA�I�ɕύX�i�J�E���g�t���j
    fileCount = 0 ' �J�E���g������
    ProcessFolder folder, oldServer, newServer, fileCount, fso
    
    ' �ϊ����ꂽ�t�@�C������\��
    MsgBox fileCount & " �t�@�C���̃����N���ύX���܂����B"
End Sub

Sub ProcessFolder(ByVal folder As Object, ByVal oldServer As String, ByVal newServer As String, ByRef fileCount As Long, ByRef fso As Object)
    Dim wshShell As Object
    Dim shortcut As Object
    Dim file As Object
    Dim subFolder As Object
    Dim backupFilePath As String

    ' �V�F���I�u�W�F�N�g���쐬
    Set wshShell = CreateObject("WScript.Shell")
    
    ' �t�H���_���̂��ׂẴt�@�C��������
    For Each file In folder.Files
        ' .lnk�t�@�C�����`�F�b�N
        If LCase(Right(file.Name, 4)) = ".lnk" Then
            ' �V���[�g�J�b�g�����[�h
            Set shortcut = wshShell.CreateShortcut(file.Path)
            
            ' �����N�悪oldServer�Ŏn�܂�ꍇ�̂ݕύX
            If Left(shortcut.TargetPath, Len(oldServer)) = oldServer Then
                ' �ύX�O�̃t�@�C�����o�b�N�A�b�v�Ƃ��ăR�s�[
                backupFilePath = fso.BuildPath(folder.Path, fso.GetBaseName(file.Name) & "_old.lnk")
                fso.CopyFile file.Path, backupFilePath, True
                
                ' �V�����T�[�o�[���ɒu������
                shortcut.TargetPath = Replace(shortcut.TargetPath, oldServer, newServer)
                shortcut.Save
                
                ' �ϊ����ꂽ�t�@�C�����J�E���g
                fileCount = fileCount + 1
            End If
        End If
    Next file
    
    ' �T�u�t�H���_���̃t�@�C�����ċA�I�ɏ���
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder, oldServer, newServer, fileCount, fso
    Next subFolder
End Sub

