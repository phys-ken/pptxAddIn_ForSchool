Attribute VB_Name = "JugyoSien"
Option Explicit
Private Declare PtrSafe Function PathIsDirectoryEmpty Lib "SHLWAPI.DLL" _
    Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Boolean

Sub Warituke()
    Dim SaveDir As String
    Dim fileName As String
    Dim fileNo As Integer
    Dim fileDir As String
    '保存場所
    fileDir = ActivePresentation.Path
    SaveDir = ActivePresentation.Path & "\tmp"
    
    
    ' tmpフォルダの中身を空にする
    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir SaveDir
        '使用可能なファイル番号を取得
        fileNo = FreeFile
        Open SaveDir & "\LogFile.txt" For Output As #fileNo
        Print #fileNo, "このファイルは削除しないでください。"
        Close #fileNo
    Else
         'PathIsDirectoryEmptyはフォルダが空のとき1を返す
        '（trueは0以外の値全てを指すため注意）
        If PathIsDirectoryEmpty(SaveDir) <> 1 Then
            Kill SaveDir & "\*"
            '使用可能なファイル番号を取得
            fileNo = FreeFile
            Open SaveDir & "\LogFile.txt" For Output As #fileNo
            Print #fileNo, "このファイルは削除しないでください。"
            Close #fileNo
        End If
    End If
  
    '画像として出力
    With ActivePresentation
        fileName = Left(.Name, InStrRev(.Name, ".") - 1)
    End With
    
    '選択スライド一枚一枚について処理
    With ActiveWindow.Selection
        If .Type >= ppSelectionSlides Then
            Debug.Print "選択スライド数: " & .SlideRange.Count
            '選択中のスライドを指定の場所にスライド番号をつけて保存
            Dim i As Long
            For i = 1 To .SlideRange.Count
                With .SlideRange(i)
                    .Export SaveDir & "\" & fileName & Format(.SlideIndex, "0000") & ".jpg", "JPG"
                End With
            Next i
        End If
    End With

    '新しいパワポを作成
    Dim tmpPrs As Presentation
    Set tmpPrs = Presentations.Add
    tmpPrs.Slides.Add _
        Index:=1, _
        Layout:=ppLayoutBlank
    ActivePresentation.PageSetup.SlideSize = ppSlideSizeA4Paper
    Call addPhotoTilingSlide(SaveDir)
    
    'PDFで保存
    With ActivePresentation
        '指定の場所にファイルを保存
        .Export fileDir & "\_配布用_" & fileName & ".pdf", "PDF"
    End With
    tmpPrs.SaveAs (fileDir & "\_配布用_" & fileName)
    MsgBox ("A4サイズの割付済PDFを作成しました。PDFファイルも同じフォルダに作成されています。")
    'ファイルを開く
    'CreateObject("Shell.Application").ShellExecute fileDir & "\" & fileName & ".pdf"

End Sub

Function addPhotoTilingSlide(SaveDir)
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ 変数定義 ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    Dim i As Integer
    ' ファイル操作
    Dim szPath As String
    Dim objFileSystem As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim strExtension As String
    ' スライドのサイズ
    Dim iSlideWidth As Integer
    Dim iSlideHeight As Integer
    ' 貼り付ける画像のサイズ
    Dim iImageWidth As Integer
    Dim iImageHeight As Integer
    ' 画像オブジェクト
    Dim stImageShape As Shape
    ' 画像データの横に並べる数
    Dim iImageColumnCount As Integer
    ' 画像データ配置時の隙間指定
    Dim iMarginSlideEdge As Integer
    Dim iMarginImage As Integer
    Dim iMarginTotal As Integer

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ パラメータ指定 ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]'
    '//[[ 画像データの横に並べる数の指定 ]]
    iImageColumnCount = 2
    '//[[ 画像データ配置の隙間 ]]
    iMarginSlideEdge = 20
    iMarginImage = 1

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ 現在のスライドのサイズをポイントで取得 ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    iSlideWidth = ActivePresentation.PageSetup.SlideWidth
    iSlideHeight = ActivePresentation.PageSetup.SlideHeight
    '//[[ マージンの演算 ]]
    iMarginTotal = iMarginSlideEdge * 2 + iMarginImage * (iImageColumnCount - 1)

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ フォルダ選択 ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    szPath = SaveDir
    Debug.Print szPath

    ' フォルダ選択されていなければ終了
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
        
        ' ファイル名の表示
        Debug.Print objFile.Path
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ 画像の挿入
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        Set stImageShape = ActiveWindow.Selection.SlideRange.Shapes.AddPicture( _
            fileName:=objFile.Path, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=0, _
            Top:=0)

        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ 画像の縦横比の固定
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        stImageShape.LockAspectRatio = msoTrue

        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ 1枚目の画像から、画像サイズ計算（フォルダ内画像はすべて同じサイズとする）
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        If i = 0 Then
            iImageWidth = Fix((iSlideWidth - iMarginTotal) / iImageColumnCount)
            stImageShape.Width = iImageWidth
            iImageHeight = stImageShape.Height
        End If

        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ 画像サイズ・位置の指定    ]]
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        stImageShape.Width = iImageWidth
        stImageShape.Height = iImageHeight
        stImageShape.Left = iMarginSlideEdge + Int(i Mod iImageColumnCount) * (iImageWidth + iMarginImage)
        stImageShape.Top = iMarginSlideEdge + Int(i / iImageColumnCount) * (iImageHeight + iMarginImage)

        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        '//[[ 画像が多い場合、スライドの追加 ]]
        '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
        If ((i + 1) Mod iImageColumnCount) = 0 And (stImageShape.Top + stImageShape.Height + iMarginImage + iImageHeight) > iSlideHeight Then
            ActivePresentation.Slides.Add( _
                Index:=ActivePresentation.Slides.Count + 1, _
                Layout:=ppLayoutBlank).Select
            i = 0
            ' スライド追加時にDoEvents（定期的にWindowsへ（ユーザーへ）制御を戻すため）
            DoEvents
        Else
            ' 10回に1度DoEvents（定期的にWindowsへ（ユーザーへ）制御を戻すため）
            i = i + 1
            If i Mod 10 = 0 Then
                DoEvents
            End If
        End If
LoopLast:
    Next

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ 終了処理]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFileSystem = Nothing
    Set stImageShape = Nothing

End Function

'//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
'//[[ Function   : フォルダ選択ダイアログ                                       ]]
'//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
Private Function SelectFolderInBrowser(Optional vRootFolder As Variant) As String
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ 変数定義                   ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    Dim objFolder As Object
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ フォルダ選択ダイアログ     ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    Set objFolder = CreateObject("Shell.Application").BrowseForFolder( _
                                 0, _
                                 "画像フォルダ選択", _
                                 &H211, _
                                 vRootFolder)

    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    '//[[ 選んだパスを取得     ]]
    '//[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]
    If Not (objFolder Is Nothing) Then
        SelectFolderInBrowser = objFolder.Items.Item.Path
    Else
        SelectFolderInBrowser = ""
    End If
    Set objFolder = Nothing
End Function

Sub changeFont_BIZUDP()
  Const FNT_NAME = "BIZ UDPゴシック"

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
