'===============================================================
' 選択スライドのタイトルテキストを「★【解答】＋元タイトル」に置き換え、
' フォント色を明るい黄色にするマクロ
'===============================================================
Public Sub convertAns()
    Dim sel As Selection
    Set sel = ActiveWindow.Selection
    
    ' スライドが選択されているかチェック
    If sel.Type < ppSelectionSlides Then
        MsgBox "スライドを選択してください。", vbExclamation
        Exit Sub
    End If
    
    Dim sr As SlideRange
    Set sr = sel.SlideRange
    
    Dim sld As Slide, shp As Shape
    For Each sld In sr
        For Each shp In sld.Shapes
            ' プレースホルダーかつタイトル系なら処理
            If shp.Type = msoPlaceholder Then
                With shp.PlaceholderFormat
                    If .Type = ppPlaceholderTitle _
                    Or  .Type = ppPlaceholderCenterTitle Then
                        
                        ' 元テキストを取得
                        With shp.TextFrame
                            If .HasText Then
                                Dim orig As String
                                orig = .TextRange.Text
                                
                                ' テキストを置換＆フォント色を変える
                                With .TextRange
                                    .Text = "★【解答】" & orig
                                    .Font.Color.RGB = RGB(255, 255, 0)  ' 明るい黄色
                                End With
                            End If
                        End With
                        
                    End If
                End With
            End If
        Next shp
    Next sld
    
    MsgBox "タイトルを解答モードに変換しました。", vbInformation
End Sub
