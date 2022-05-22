Attribute VB_Name = "JugyoSien"
Option Explicit
Private Declare PtrSafe Function PathIsDirectoryEmpty Lib "SHLWAPI.DLL" _
    Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Boolean

Sub Warituke()
    Dim SaveDir As String
    Dim fileName As String
    Dim fileNo As Integer
    Dim fileDir As String
    '�ۑ��ꏊ
    fileDir = ActivePresentation.Path
    SaveDir = ActivePresentation.Path & "\tmp"
    
    
    ' tmp�t�H���_�̒��g����ɂ���
    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir SaveDir
        '�g�p�\�ȃt�@�C���ԍ����擾
        fileNo = FreeFile
        Open SaveDir & "\LogFile.txt" For Output As #fileNo
        Print #fileNo, "���̃t�@�C���͍폜���Ȃ��ł��������B"
        Close #fileNo
    Else
         'PathIsDirectoryEmpty�̓t�H���_����̂Ƃ�1��Ԃ�
        '�itrue��0�ȊO�̒l�S�Ă��w�����ߒ��Ӂj
        If PathIsDirectoryEmpty(SaveDir) <> 1 Then
            Kill SaveDir & "\*"
            '�g�p�\�ȃt�@�C���ԍ����擾
            fileNo = FreeFile
            Open SaveDir & "\LogFile.txt" For Output As #fileNo
            Print #fileNo, "���̃t�@�C���͍폜���Ȃ��ł��������B"
            Close #fileNo
        End If
    End If
  
    '�摜�Ƃ��ďo��
    With ActivePresentation
        fileName = Left(.Name, InStrRev(.Name, ".") - 1)
    End With
    
    '�I���X���C�h�ꖇ�ꖇ�ɂ��ď���
    With ActiveWindow.Selection
        If .Type >= ppSelectionSlides Then
            Debug.Print "�I���X���C�h��: " & .SlideRange.Count
            '�I�𒆂̃X���C�h���w��̏ꏊ�ɃX���C�h�ԍ������ĕۑ�
            Dim i As Long
            For i = 1 To .SlideRange.Count
                With .SlideRange(i)
                    .Export SaveDir & "\" & fileName & Format(.SlideIndex, "0000") & ".jpg", "JPG"
                End With
            Next i
        End If
    End With

    '�V�����p���|���쐬
    Dim tmpPrs As Presentation
    Set tmpPrs = Presentations.Add
    tmpPrs.Slides.Add _
        Index:=1, _
        Layout:=ppLayoutBlank
    ActivePresentation.PageSetup.SlideSize = ppSlideSizeA4Paper
    Call addPhotoTilingSlide(SaveDir)
    
    'PDF�ŕۑ�
    With ActivePresentation
        '�w��̏ꏊ�Ƀt�@�C����ۑ�
        .Export fileDir & "\_�z�z�p_" & fileName & ".pdf", "PDF"
    End With
    tmpPrs.SaveAs (fileDir & "\_�z�z�p_" & fileName)
    MsgBox ("A4�T�C�Y�̊��t��PDF���쐬���܂����BPDF�t�@�C���������t�H���_�ɍ쐬����Ă��܂��B")
    '�t�@�C�����J��
    'CreateObject("Shell.Application").ShellExecute fileDir & "\" & fileName & ".pdf"

End Sub

Function addPhotoTilingSlide(SaveDir)
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ �ϐ���` ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    Dim i As Integer
    ' �t�@�C������
    Dim szPath As String
    Dim objFileSystem As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim strExtension As String
    ' �X���C�h�̃T�C�Y
    Dim iSlideWidth As Integer
    Dim iSlideHeight As Integer
    ' �\��t����摜�̃T�C�Y
    Dim iImageWidth As Integer
    Dim iImageHeight As Integer
    ' �摜�I�u�W�F�N�g
    Dim stImageShape As Shape
    ' �摜�f�[�^�̉��ɕ��ׂ鐔
    Dim iImageColumnCount As Integer
    ' �摜�f�[�^�z�u���̌��Ԏw��
    Dim iMarginSlideEdge As Integer
    Dim iMarginImage As Integer
    Dim iMarginTotal As Integer

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ �p�����[�^�w�� ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]'
    '//[[ �摜�f�[�^�̉��ɕ��ׂ鐔�̎w�� ]]
    iImageColumnCount = 2
    '//[[ �摜�f�[�^�z�u�̌��� ]]
    iMarginSlideEdge = 20
    iMarginImage = 1

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ ���݂̃X���C�h�̃T�C�Y���|�C���g�Ŏ擾 ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    iSlideWidth = ActivePresentation.PageSetup.SlideWidth
    iSlideHeight = ActivePresentation.PageSetup.SlideHeight
    '//[[ �}�[�W���̉��Z ]]
    iMarginTotal = iMarginSlideEdge * 2 + iMarginImage * (iImageColumnCount - 1)

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ �t�H���_�I�� ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    szPath = SaveDir
    Debug.Print szPath

    ' �t�H���_�I������Ă��Ȃ���ΏI��
    If szPath = "" Then
        Exit Function
    End If

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFileSystem.GetFolder(szPath)

    i = 0
    For Each objFile In objFolder.Files
        strExtension = objFileSystem.GetExtensionName(objFile.Name)
        If strExtension <> "jpg" Then
            GoTo LoopLast
        End If
        
        ' �t�@�C�����̕\��
        Debug.Print objFile.Path
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ �摜�̑}��
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        Set stImageShape = ActiveWindow.Selection.SlideRange.Shapes.AddPicture( _
            fileName:=objFile.Path, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=0, _
            Top:=0)

        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ �摜�̏c����̌Œ�
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        stImageShape.LockAspectRatio = msoTrue

        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ 1���ڂ̉摜����A�摜�T�C�Y�v�Z�i�t�H���_���摜�͂��ׂē����T�C�Y�Ƃ���j
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        If i = 0 Then
            iImageWidth = Fix((iSlideWidth - iMarginTotal) / iImageColumnCount)
            stImageShape.Width = iImageWidth
            iImageHeight = stImageShape.Height
        End If

        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ �摜�T�C�Y�E�ʒu�̎w��    ]]
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        stImageShape.Width = iImageWidth
        stImageShape.Height = iImageHeight
        stImageShape.Left = iMarginSlideEdge + Int(i Mod iImageColumnCount) * (iImageWidth + iMarginImage)
        stImageShape.Top = iMarginSlideEdge + Int(i / iImageColumnCount) * (iImageHeight + iMarginImage)

        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ �摜�������ꍇ�A�X���C�h�̒ǉ� ]]
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        If ((i + 1) Mod iImageColumnCount) = 0 And (stImageShape.Top + stImageShape.Height + iMarginImage + iImageHeight) > iSlideHeight Then
            ActivePresentation.Slides.Add( _
                Index:=ActivePresentation.Slides.Count + 1, _
                Layout:=ppLayoutBlank).Select
            i = 0
            ' �X���C�h�ǉ�����DoEvents�i����I��Windows�ցi���[�U�[�ցj�����߂����߁j
            DoEvents
        Else
            ' 10���1�xDoEvents�i����I��Windows�ցi���[�U�[�ցj�����߂����߁j
            i = i + 1
            If i Mod 10 = 0 Then
                DoEvents
            End If
        End If
LoopLast:
    Next

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ �I������]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFileSystem = Nothing
    Set stImageShape = Nothing

End Function

'//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
'//[[ Function   : �t�H���_�I���_�C�A���O                                       ]]
'//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
Private Function SelectFolderInBrowser(Optional vRootFolder As Variant) As String
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ �ϐ���`                   ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    Dim objFolder As Object
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ �t�H���_�I���_�C�A���O     ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    Set objFolder = CreateObject("Shell.Application").BrowseForFolder( _
                                 0, _
                                 "�摜�t�H���_�I��", _
                                 &H211, _
                                 vRootFolder)

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ �I�񂾃p�X���擾     ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    If Not (objFolder Is Nothing) Then
        SelectFolderInBrowser = objFolder.Items.Item.Path
    Else
        SelectFolderInBrowser = ""
    End If
    Set objFolder = Nothing
End Function

Sub changeFont_BIZUDP()
  Const FNT_NAME = "BIZ UDP�S�V�b�N"

  Dim shp As Shape

  With ActivePresentation.SlideMaster.Shapes
    For Each shp In .Placeholders

      If shp.PlaceholderFormat.Type <> ppPlaceholderDate Then
        With shp.TextFrame.TextRange.Font
          .Name = FNT_NAME
          .NameFarEast = FNT_NAME
        End With
      End If

    Next shp
  End With

End Sub
