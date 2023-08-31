'モジュール内で共通して使える変数宣言
Dim macroFlag As Boolean, filename As String, rng As Range, sheetname As String, JANcoderng As String, num As Integer, JANcode As String, filepath As String, filepath_all
Function ConvertToBarcodeTwo(Inte As Integer)
'13桁JANコード用の事業者コードの変換

    If Inte = 0 Then
        ConvertToBarcodeTwo = "@"
    ElseIf Inte = 1 Then
        ConvertToBarcodeTwo = "A"
    ElseIf Inte = 2 Then
        ConvertToBarcodeTwo = "B"
    ElseIf Inte = 3 Then
        ConvertToBarcodeTwo = "C"
    ElseIf Inte = 4 Then
        ConvertToBarcodeTwo = "D"
    ElseIf Inte = 5 Then
        ConvertToBarcodeTwo = "E"
    ElseIf Inte = 6 Then
        ConvertToBarcodeTwo = "F"
    ElseIf Inte = 7 Then
        ConvertToBarcodeTwo = "G"
    ElseIf Inte = 8 Then
        ConvertToBarcodeTwo = "H"
    Else
        ConvertToBarcodeTwo = "I"
    End If
End Function
Function ConvertToBarcodeShopCode(Inte As Integer)
'JANコードの商品コード用の変換

    If Inte = 0 Then
        ConvertToBarcodeShopCode = "P"
    ElseIf Inte = 1 Then
        ConvertToBarcodeShopCode = "Q"
    ElseIf Inte = 2 Then
        ConvertToBarcodeShopCode = "R"
    ElseIf Inte = 3 Then
        ConvertToBarcodeShopCode = "S"
    ElseIf Inte = 4 Then
        ConvertToBarcodeShopCode = "T"
    ElseIf Inte = 5 Then
        ConvertToBarcodeShopCode = "U"
    ElseIf Inte = 6 Then
        ConvertToBarcodeShopCode = "V"
    ElseIf Inte = 7 Then
        ConvertToBarcodeShopCode = "W"
    ElseIf Inte = 8 Then
        ConvertToBarcodeShopCode = "X"
    Else
        ConvertToBarcodeShopCode = "Y"
    End If
End Function

Sub FileSelectOpen()
'JANコードが含まれているファイルをダイアログから選択させ、ファイル名を変数filenameに代入する
    
    'ダイアログを表示し、ファイルパスを取得する
    filepath_all = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xl*;*.csv;*.xml;*.txt;*.prn;*.dif;*.slk", MultiSelect:=True)

    'ファイルが選択されなかった場合の例外処理
    If IsArray(filepath_all) = False Then
        MsgBox ("ファイルが選択されなかったため、マクロは停止されます。")
        End
    End If
End Sub
Sub CellSelect()
'JANコードの一番上のセルを選択させ、シート名を変数sheetnameにセルアドレスを変数JANcoderngに代入する

'エラー回避用
On Error GoTo ErrHandl

    '特定の関数内でのみ使える変数宣言
    Dim n As Integer, notJANFlag As Boolean, Flag As Boolean
    
    'ループを抜け出す用のフラグの初期化
    Flag = False
    
    'JANコードのあるセルが選択されるまでループを続ける
    While Flag = False
        'JANコードか判断するフラグの初期化
        notJANFlag = False
        'セル選択用のInputBox表示
        Set rng = Application.InputBox( _
            Prompt:="JANコードの一番上のセルを選択してOKを押してください。", _
            Type:=8)
        'セルの値が全て数字かチェックする
        For n = 1 To Len(rng)
            'もし数字ではなかった場合、JANコード判断用フラグを立てる
            If IsNumeric(Mid(rng, n, 1)) = False Then
                notJANFlag = True
            End If
            'セルの値が13桁もしくは8桁かチェックする
            '13桁でも8桁でもない場合、JANコード判断用フラグを立てる
            If Len(rng) <> 13 Then
                If Len(rng) <> 8 Then
                    notJANFlag = True
                End If
            End If
        Next
        '選択されたセルがJANコードだった場合、ループを抜けるフラグが立つ
        If notJANFlag = False Then
            Flag = True
        '選択されたセルがJANコードでなかった場合、JANコード出ない旨を通知しループを続ける
        Else
            MsgBox ("JANコードではないセルが選択されました。" & vbCrLf & "正しいセルを選択しなおしてください。")
        End If
    Wend
    '変数filenameにブック名を代入する
    filename = rng.Parent.Parent.Name
    '変数sheetnameにシート名を代入する
    sheetname = rng.Parent.Name
    '変数JANcoderngにセルの位置を代入する
    JANcoderng = rng.Address(False, False)
Exit Sub

'キャンセル時の処理
ErrHandl:
    MsgBox ("マクロを強制終了します")
    If macroFlag = True Then
        filename = Dir(filepath)
        Application.DisplayAlerts = False
        Workbooks(filename).Close
        Application.DisplayAlerts = True
    End If
    End
End Sub
Sub ActiveSelectedCell()
'指定のブックのシートをアクティブにする

    If filename = "" Then
        End
    Else
        '指定のワークブックをアクティブにする
        Workbooks(filename).Activate
    End If
    '指定のシートをアクティブにする
    Worksheets(sheetname).Activate
End Sub
Sub ConvertToBarcode()
'指定のセルのJANコードをバーコードフォント用に変換する

    '特定の関数内でのみ使える変数宣言
    Dim L1 As String, L2 As String, L3 As String, L4 As String, L5 As String, L6 As String, L7 As String, L8 As String, L9 As String, L10 As String, L11 As String, L12 As String, CD As Integer
    
    'バーコードを作成する列の幅を18にする
    ActiveWorkbook.ActiveSheet.Columns(rng.Column + 1).ColumnWidth = 18
    
    'JANコードそれぞれの桁をバーコード用に変換する
    For n = rng.Row To Cells(rng.Row, rng.Column).End(xlDown).Row
        '変数JANcodeにJANコードを代入する
        JANcode = Cells(n, rng.Column).Value
        If Len(JANcode) = 13 Then
            '以下JANコードが13桁だった場合の処理
            CD = Left(JANcode, 1)
            If CD = 0 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = Mid(JANcode, 3, 1)
                L3 = Mid(JANcode, 4, 1)
                L4 = Mid(JANcode, 5, 1)
                L5 = Mid(JANcode, 6, 1)
                L6 = Mid(JANcode, 7, 1)
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 1 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = Mid(JANcode, 3, 1)
                L3 = ConvertToBarcodeTwo(Mid(JANcode, 4, 1))
                L4 = Mid(JANcode, 5, 1)
                L5 = ConvertToBarcodeTwo(Mid(JANcode, 6, 1))
                L6 = ConvertToBarcodeTwo(Mid(JANcode, 7, 1))
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 2 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = Mid(JANcode, 3, 1)
                L3 = ConvertToBarcodeTwo(Mid(JANcode, 4, 1))
                L4 = ConvertToBarcodeTwo(Mid(JANcode, 5, 1))
                L5 = Mid(JANcode, 6, 1)
                L6 = ConvertToBarcodeTwo(Mid(JANcode, 7, 1))
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 3 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = Mid(JANcode, 3, 1)
                L3 = ConvertToBarcodeTwo(Mid(JANcode, 4, 1))
                L4 = ConvertToBarcodeTwo(Mid(JANcode, 5, 1))
                L5 = ConvertToBarcodeTwo(Mid(JANcode, 6, 1))
                L6 = Mid(JANcode, 7, 1)
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 4 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = ConvertToBarcodeTwo(Mid(JANcode, 3, 1))
                L3 = Mid(JANcode, 4, 1)
                L4 = Mid(JANcode, 5, 1)
                L5 = ConvertToBarcodeTwo(Mid(JANcode, 6, 1))
                L6 = ConvertToBarcodeTwo(Mid(JANcode, 7, 1))
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 5 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = ConvertToBarcodeTwo(Mid(JANcode, 3, 1))
                L3 = ConvertToBarcodeTwo(Mid(JANcode, 4, 1))
                L4 = Mid(JANcode, 5, 1)
                L5 = Mid(JANcode, 6, 1)
                L6 = ConvertToBarcodeTwo(Mid(JANcode, 7, 1))
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 6 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = ConvertToBarcodeTwo(Mid(JANcode, 3, 1))
                L3 = ConvertToBarcodeTwo(Mid(JANcode, 4, 1))
                L4 = ConvertToBarcodeTwo(Mid(JANcode, 5, 1))
                L5 = Mid(JANcode, 6, 1)
                L6 = Mid(JANcode, 7, 1)
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 7 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = ConvertToBarcodeTwo(Mid(JANcode, 3, 1))
                L3 = Mid(JANcode, 4, 1)
                L4 = ConvertToBarcodeTwo(Mid(JANcode, 5, 1))
                L5 = Mid(JANcode, 6, 1)
                L6 = ConvertToBarcodeTwo(Mid(JANcode, 7, 1))
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 8 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = ConvertToBarcodeTwo(Mid(JANcode, 3, 1))
                L3 = Mid(JANcode, 4, 1)
                L4 = ConvertToBarcodeTwo(Mid(JANcode, 5, 1))
                L5 = ConvertToBarcodeTwo(Mid(JANcode, 6, 1))
                L6 = Mid(JANcode, 7, 1)
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            If CD = 9 Then
                L1 = Mid(JANcode, 2, 1)
                L2 = ConvertToBarcodeTwo(Mid(JANcode, 3, 1))
                L3 = ConvertToBarcodeTwo(Mid(JANcode, 4, 1))
                L4 = Mid(JANcode, 5, 1)
                L5 = ConvertToBarcodeTwo(Mid(JANcode, 6, 1))
                L6 = Mid(JANcode, 7, 1)
                L7 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
                L8 = ConvertToBarcodeShopCode(Mid(JANcode, 9, 1))
                L9 = ConvertToBarcodeShopCode(Mid(JANcode, 10, 1))
                L10 = ConvertToBarcodeShopCode(Mid(JANcode, 11, 1))
                L11 = ConvertToBarcodeShopCode(Mid(JANcode, 12, 1))
                L12 = ConvertToBarcodeShopCode(Mid(JANcode, 13, 1))
            End If
            '変数Barcodeに変換した文字を全てつなげて文字列にして代入する
            Barcode = "(" & L1 & L2 & L3 & L4 & L5 & L6 & "|" & L7 & L8 & L9 & L10 & L11 & L12 & L13 & ")"
            '指定のセルの１個右にバーコード用文字列を上書きする
            ActiveWorkbook.ActiveSheet.Cells(n, rng.Column + 1).Value = Barcode
            'バーコード用セルのフォントをバーコード用フォントにする
            ActiveWorkbook.ActiveSheet.Cells(n, rng.Column + 1).Font.Name = "JAN TT"
            'バーコード用セルの文字を大きくする
            ActiveWorkbook.ActiveSheet.Cells(n, rng.Column + 1).Font.Size = 60
            'バーコードを作成する行の高さを50にする
            ActiveWorkbook.ActiveSheet.Rows(n).RowHeight = 50
        ElseIf Len(JANcode) = 8 Then
            '以下JANコードが8桁だった場合の処理
            L1 = Mid(JANcode, 1, 1)
            L2 = Mid(JANcode, 2, 1)
            L3 = Mid(JANcode, 3, 1)
            L4 = Mid(JANcode, 4, 1)
            L5 = ConvertToBarcodeShopCode(Mid(JANcode, 5, 1))
            L6 = ConvertToBarcodeShopCode(Mid(JANcode, 6, 1))
            L7 = ConvertToBarcodeShopCode(Mid(JANcode, 7, 1))
            L8 = ConvertToBarcodeShopCode(Mid(JANcode, 8, 1))
            '変数Barcodeに変換した文字を全てつなげて文字列にして代入する
            Barcode = "(" & L1 & L2 & L3 & L4 & "|" & L5 & L6 & L7 & L8 & ")"
            '指定のセルの１個右にバーコード用文字列を上書きする
            ActiveWorkbook.ActiveSheet.Cells(n, rng.Column + 1).Value = Barcode
            'バーコード用セルのフォントをバーコード用フォントにする
            ActiveWorkbook.ActiveSheet.Cells(n, rng.Column + 1).Font.Name = "JAN TT"
            'バーコード用セルの文字を大きくする
            ActiveWorkbook.ActiveSheet.Cells(n, rng.Column + 1).Font.Size = 60
            'バーコードを作成する行の高さを50にする
            ActiveWorkbook.ActiveSheet.Rows(n).RowHeight = 50
        End If
    Next
End Sub
Sub SaveAndClose()
'ファイルを保存して閉じる
    'ファイルを保存する際の確認ポップアップを消す
    Application.DisplayAlerts = False
    'ファイルを保存する
    Workbooks(filename).Save
    'ファイルを閉じる
    Workbooks(filename).Close
    'ファイルを保存する際のポップアップを戻す
    Application.DisplayAlerts = True
End Sub
Sub JANバーコード作成_13桁8桁対応_ファイル指定()
'全体実行マクロ
    macroFlag = True
    'ファイル選択ダイアログでファイルを選択、開く
    FileSelectOpen
    'セル選択用のInputBoxを表示させ、JANコードの一番上のセルを取得する
    For i = 1 To UBound(filepath_all)
        filepath = filepath_all(i)
        Workbooks.Open filepath
        CellSelect
        '選択されたブック、シートをアクティブにする
        ActiveSelectedCell
        '指定のセルのJANコードをバーコードフォント用に変換し、１個右のセルにバーコードを表示するする
        ConvertToBarcode
        'ファイルを保存して閉じる
        SaveAndClose
    Next
End Sub
Sub JANバーコード作成_13桁8桁対応_セル選択のみ()
'全体実行マクロ
    'セル選択用のInputBoxを表示させ、JANコードの一番上のセルを取得する
    CellSelect
    '選択されたブック、シートをアクティブにする
    ActiveSelectedCell
    '指定のセルのJANコードをバーコードフォント用に変換し、１個右のセルにバーコードを表示するする
    ConvertToBarcode
End Sub
