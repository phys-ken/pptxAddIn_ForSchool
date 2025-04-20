'===================================================================
' Module : waritukeMaxModule
'          A3 縦スライドへ “幅優先・左上詰め” で高密度タイリング
'===================================================================
Option Explicit

'--- Windows API 宣言(32bit/64bit両対応) --------------------------
#If VBA7 Then
    Private Declare PtrSafe Function PathIsDirectoryEmpty Lib "Shlwapi.dll" _
        Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
#Else
    Private Declare Function PathIsDirectoryEmpty Lib "Shlwapi.dll" _
        Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
#End If

'--- 定数 ----------------------------------------------------------
Private Const COL_COUNT  As Long  = 2       ' 列数固定
Private Const MARGIN_TB  As Single = 28     ' 上下マージン(pt)
Private Const MARGIN_LR  As Single = 28     ' 左右マージン(pt)
Private Const GAP_MIN    As Single = 8.5    ' 最小ギャップ(pt)
Private Const MIN_RATIO  As Single = 0.05   ' 拡大率下限
' PowerPoint 定数（バージョン依存対策）
Private Const ppSaveAsPDF          As Long = 32
Private Const ppFixedFormatTypePDF As Long = 2

'===================================================================
' メイン
'===================================================================
Public Sub waritukeMax()

    Dim fileDir As String, saveDir As String, fileName As String
    Dim fNum As Integer
    
    '---------------------------------------------------------------
    ' ① tmp フォルダ準備
    '---------------------------------------------------------------
    fileDir = ActivePresentation.Path
    saveDir = fileDir & "\tmp"
    
    If Dir(saveDir, vbDirectory) = "" Then
        MkDir saveDir
    ElseIf PathIsDirectoryEmpty(saveDir) <> 1 Then
        Kill saveDir & "\*"
    End If
    
    fNum = FreeFile
    Open saveDir & "\LogFile.txt" For Output As #fNum
        Print #fNum, "このファイルは削除しないでください。"
    Close #fNum
    
    '---------------------------------------------------------------
    ' ② スライド取得 → JPG 化
    '---------------------------------------------------------------
    Dim sr As SlideRange
    With ActiveWindow.Selection
        If .Type < ppSelectionSlides Then _
            MsgBox "スライドが選択されていません。": Exit Sub
        Set sr = .SlideRange
    End With
    
    Dim arr() As Slide, i As Long, j As Long
    ReDim arr(1 To sr.Count)
    For i = 1 To sr.Count: Set arr(i) = sr(i): Next i
    
    ' インデックス昇順へ並べ替え
    Dim tmp As Slide
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j).SlideIndex < arr(i).SlideIndex Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i
    
    fileName = Left(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1)
    For i = 1 To UBound(arr)
        arr(i).Export saveDir & "\" & fileName & Format(i, "0000") & ".jpg", "JPG"
    Next i
    
    '---------------------------------------------------------------
    ' ③ 代表画像サイズ取得
    '---------------------------------------------------------------
    Dim repW As Single, repH As Single
    With ActivePresentation.Slides(1).Shapes _
         .AddPicture(saveDir & "\" & fileName & "0001.jpg", msoFalse, msoTrue, -1000, -1000)
        repW = .Width: repH = .Height: .Delete
    End With
    
    '---------------------------------------------------------------
    ' ④ 新プレゼン作成（A3 縦） & レイアウト決定
    '---------------------------------------------------------------
    Dim newPrs As Presentation: Set newPrs = Presentations.Add
    With newPrs.PageSetup
        .SlideWidth = 842  ' 297 mm
        .SlideHeight = 1190 ' 420 mm
    End With
    newPrs.Slides.Add 1, ppLayoutBlank
    
    Dim sldW As Single: sldW = newPrs.PageSetup.SlideWidth
    Dim sldH As Single: sldH = newPrs.PageSetup.SlideHeight
    
    Dim usableW As Single: usableW = (sldW - 2 * MARGIN_LR) - (COL_COUNT - 1) * GAP_MIN
    Dim ratioW As Single:  ratioW = usableW / (COL_COUNT * repW)
    If ratioW < MIN_RATIO Then ratioW = MIN_RATIO
    
    Dim newW As Single: newW = repW * ratioW
    Dim newH As Single: newH = repH * ratioW
    
    Dim usableH As Single: usableH = sldH - 2 * MARGIN_TB
    Dim rowCnt As Long: rowCnt = Int((usableH + GAP_MIN) / (newH + GAP_MIN))
    If rowCnt < 1 Then rowCnt = 1
    
    Dim chunkSize As Long: chunkSize = rowCnt * COL_COUNT
    
    '---------------------------------------------------------------
    ' ⑤ 画像タイリング
    '---------------------------------------------------------------
    addPhotoTilingSlidesPacked newPrs, saveDir, fileName, rowCnt, newW, newH, chunkSize
    
    '---------------------------------------------------------------
    ' ⑥ PDF & PPTX 保存
    '---------------------------------------------------------------
    Dim outPdf As String, outPptx As String
    outPdf  = fileDir & "\_配布用_" & fileName & ".pdf"
    outPptx = fileDir & "\_配布用_" & fileName & ".pptx"
    
    ' 既存ファイルがあれば削除（上書き失敗回避）
    If Dir(outPdf)  <> "" Then Kill outPdf
    If Dir(outPptx) <> "" Then Kill outPptx
    
    ' --- PDF 出力 ---
    On Error Resume Next
    newPrs.SaveAs outPdf, ppSaveAsPDF              ' 通常はこちらで成功
    If Err.Number <> 0 Then
        Err.Clear
        newPrs.ExportAsFixedFormat outPdf, ppFixedFormatTypePDF
    End If
    On Error GoTo 0
    
    ' --- PPTX 保存 ---
    newPrs.SaveAs outPptx
    
    ' ページ数計算（切り上げ）
    Dim imgTotal As Long: imgTotal = sr.Count
    Dim pageTotal As Long: pageTotal = (imgTotal + chunkSize - 1) \ chunkSize
    
    MsgBox "割付完了 : 1ページ " & chunkSize & " 枚 × " & _
           pageTotal & " ページで出力しました。", vbInformation
End Sub

'===================================================================
' 左上詰めタイル配置
'===================================================================
Private Sub addPhotoTilingSlidesPacked(prs As Presentation, _
                                       saveDir As String, _
                                       baseName As String, _
                                       rowCnt As Long, _
                                       w As Single, h As Single, _
                                       chunkSize As Long)
    
    '--- 画像リスト収集 --------------------------------------------
    Dim fso As Object, fld As Object, fil As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(saveDir)
    
    Dim lst As New Collection
    For Each fil In fld.Files
        If InStr(fil.Name, baseName) > 0 And Right(fil.Name, 4) = ".jpg" Then lst.Add fil.Path
    Next
    If lst.Count = 0 Then MsgBox "画像が見つかりません。": Exit Sub
    
    '--- 各ページ配置 ----------------------------------------------
    Dim startIdx As Long: startIdx = 1
    Dim pageIdx  As Long: pageIdx = 1
    
    Do While startIdx <= lst.Count
        
        Dim endIdx As Long: endIdx = startIdx + chunkSize - 1
        If endIdx > lst.Count Then endIdx = lst.Count
        
        Dim sld As Slide
        If pageIdx = 1 Then
            Set sld = prs.Slides(1)
        Else
            Set sld = prs.Slides.Add(pageIdx, ppLayoutBlank)
        End If
        pageIdx = pageIdx + 1
        
        Dim r As Long, c As Long, k As Long: k = startIdx
        For r = 0 To rowCnt - 1
            For c = 0 To COL_COUNT - 1
                If k > endIdx Then Exit For
                
                Dim shp As Shape
                Set shp = sld.Shapes.AddPicture(lst.Item(k), msoFalse, msoTrue, 0, 0)
                shp.Width = w: shp.Height = h
                shp.Left = MARGIN_LR + c * (w + GAP_MIN)
                shp.Top  = MARGIN_TB + r * (h + GAP_MIN)
                
                With shp.Line
                    .Visible = msoTrue: .Weight = 1: .ForeColor.RGB = RGB(0, 0, 0)
                End With
                
                k = k + 1
            Next c
        Next r
        
        startIdx = endIdx + 1
    Loop
End Sub
