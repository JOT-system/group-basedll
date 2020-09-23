Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports BATDLL

Module CB00008TBLupdate3

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.CB00008TBLupdate3 /@1 /@2 /@3         　　　　　　　　　　　　　　　　　 ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：テーブル記号名称　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：入力先(ディレクトリ+ファイル名)                                             ■
    '■          ※省略時、 c:\APPL\FILES\RECEIVE\SRVENEX\TABLE\テーブル名.dat"とする        ■
    '■　　・@3：プロファイルＩＤ一覧ファイルパス                                            ■
    '■　注意　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　入力ファイルにヘッダ行は必須、主キー無しテーブルはサポート外　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Dim WW_Now As Date = Date.Now
    Const CompCode As String = "02"
    Const CompCode0 As String = "020"
    Const breakCnt As Integer = 1

    Sub Main()
        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_TBLNAME As String = ""
        Dim WW_InPARA_FilePath As String = ""
        Dim WW_InPARA_ProfIDInfo As String = ""

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat          'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New BATDLL.CS0053FILEdir_bat      'アップロードFile格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     'テーブル記号名称
                    WW_InPARA_TBLNAME = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(テーブル名　　　)：" & WW_InPARA_TBLNAME)
                Case 2     '入力先(ディレクトリ+ファイル名)
                    WW_InPARA_FilePath = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(入力先　　　　　)：" & WW_InPARA_FilePath)
                Case 3     'PROFID一覧ファイルパス
                    WW_InPARA_ProfIDInfo = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(プロファイルＩＤ一覧　　)：" & WW_InPARA_ProfIDInfo)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"                 'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                                'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                             'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00008TBLupdate3処理開始"            '
        CS0054LOGWrite_bat.NIWEA = "I"                                      '
        CS0054LOGWrite_bat.TEXT = "CB00008TBLupdate3.exe /" & WW_InPARA_TBLNAME & " /" & WW_InPARA_FilePath & " /" & WW_InPARA_ProfIDInfo & " "
        CS0054LOGWrite_bat.MESSAGENO = "00000"                              'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                             'ログ入力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        Dim WW_SRVname As String = ""
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)                'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0051APSRVname_bat"          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "APサーバー名称取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "APサーバー名称取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0051APSRVname_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ ログ格納ディレクトリ取得(InParm無し)
        Dim WW_LOGdir As String = ""
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If CS0052LOGdir_bat.ERR = "00000" Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                    'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ログ格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "ログ格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ アップロードFile格納ディレクトリ取得(InParm無し)
        Dim WW_FILEdir As String = ""
        CS0053FILEdir_bat.CS0053FILEdir_bat()
        If CS0053FILEdir_bat.ERR = "00000" Then
            WW_FILEdir = Trim(CS0053FILEdir_bat.FILEdirStr)                 'アップロードFile格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "File格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "File格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0053FILEdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If


        '■■■　コマンドライン第二引数(入力先)より対象ディレクトリの決定　■■■
        Dim WW_Folder As String = ""
        Dim WW_UPfiles As String() = {}

        'ディレクトリ指定無しの場合、デフォルト(c:\APPL\APPLFILES\SRVENEX\)設定
        If WW_InPARA_FilePath = "" Then
            WW_Folder = WW_FILEdir & "\RECEIVE\SRVENEX\"
        Else
            '末尾に\を付加する
            If WW_InPARA_FilePath.LastIndexOf("\") <> WW_InPARA_FilePath.Length - 1 Then
                WW_Folder = WW_InPARA_FilePath & "\"
            Else
                WW_Folder = WW_InPARA_FilePath
            End If

            'コマンドライン第二引数(出力先)のチェック  …　自SRVディレクトリのみ可(\\xxxx形式は×)
            If InStr(WW_Folder, ":") = 0 Or Mid(WW_Folder, 2, 1) <> ":" Then
                CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
                CS0054LOGWrite_bat.NIWEA = "E"                                  '
                CS0054LOGWrite_bat.TEXT = "引数2フォーマットエラー：" & WW_InPARA_FilePath
                CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                Environment.Exit(100)
            End If
        End If

        '■■■　コマンドライン第一引数(テーブル)のチェック＆対象ファイル取得　■■■
        '○ パラメータチェック(テーブル名)
        If System.IO.Directory.Exists(WW_Folder) Then
            If WW_InPARA_TBLNAME = "" Then
                WW_UPfiles = System.IO.Directory.GetFiles(WW_Folder, "*.dat", System.IO.SearchOption.AllDirectories)
            Else
                WW_UPfiles = System.IO.Directory.GetFiles(WW_Folder, WW_InPARA_TBLNAME & ".dat", System.IO.SearchOption.AllDirectories)
            End If
        End If

        'PROFID取得
        '入力ファイル検索
        Dim ProfTxt As String = WW_InPARA_ProfIDInfo
        Dim PrfSr As New System.IO.StreamReader(ProfTxt, System.Text.Encoding.GetEncoding("utf-8"))
        Dim RdLine As String = ""
        'Dim PrfList As New ArrayList()
        Dim PrfList As New List(Of String())

        Try
            '■File情報をすべて読み込む
            While (Not PrfSr.EndOfStream)

                RdLine = PrfSr.ReadLine()
                Dim splLine As String() = {}
                splLine = RdLine.Split(",")

                PrfList.Add(splLine)

            End While

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "PROFID取得失敗"                   '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
        End Try

        '■■■　テーブル更新処理　■■■
        For Each WW_file As String In WW_UPfiles

            '送信されたフォルダー（端末ID）が自サーバーだったら対象
            If WW_file.IndexOf(WW_SRVname & "\") < 0 Then
                Continue For
            End If

            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "テーブル更新ファイル"             '
            CS0054LOGWrite_bat.NIWEA = "W"                                  '
            CS0054LOGWrite_bat.TEXT = "処理ファイル（" & WW_file & "）"
            CS0054LOGWrite_bat.MESSAGENO = "00000"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

            'ファイル名からテーブル名を取り出す
            WW_InPARA_TBLNAME = System.IO.Path.GetFileName(WW_file).Replace(".dat", "")

            'テーブル名変更
            Select Case WW_InPARA_TBLNAME
                Case "L0001_TOKEI"
                Case "L0003_SUMMARYN"
                Case "L0004_SUMMARYK"
                Case "L0005_SUMMARYY"
                Case "MA002_SHARYOA"
                Case "MA003_SHARYOB"
                Case "MA004_SHARYOC"
                Case "MA006_SHABANORG"
                Case "MB002_STAFFORG"
                Case "MC001_FIXVALUE"
                Case "MC002_TORIHIKISAKI"
                Case "MC007_TODKORG"
                Case "MC004_PRODUCT"
                    WW_InPARA_TBLNAME = "MD001_PRODUCT"
                Case "MC005_PRODORG"
                    WW_InPARA_TBLNAME = "MD002_PRODORG"
                Case "S0004_USER"
                Case "S0007_UPROFVARI"
                    WW_InPARA_TBLNAME = "S0023_PROFMVARI"
                Case "S0010_UPROFVIEW"
                    WW_InPARA_TBLNAME = "S0025_PROFMVIEW"
                Case "S0011_UPROFXLS"
                    WW_InPARA_TBLNAME = "S0026_PROFMXLS"
                Case "T0003_NIORDER"
                Case "T0004_HORDER"
                Case "T0005_NIPPO"
                Case "T0007_KINTAI"
                Case "TA001_SHARYOSTAT"
                Case Else
                    Continue For
            End Select

            Dim fileName As String = WW_file

            '項目削除、名称変更
            Select Case WW_InPARA_TBLNAME
                    'Case "MC004_PRODUCT"
                Case "MD001_PRODUCT"
                    fileName = WW_file.Replace("MC004_PRODUCT", "MD001_PRODUCT")

                    'Case "MC005_PRODORG"
                Case "MD002_PRODORG"
                    fileName = WW_file.Replace("MC005_PRODORG", "MD002_PRODORG")

                    'Case "S0007_UPROFVARI"
                Case "S0023_PROFMVARI"
                    fileName = WW_file.Replace("S0007_UPROFVARI", "S0023_PROFMVARI")

                    'Case "S0010_UPROFVIEW"
                Case "S0025_PROFMVIEW"
                    fileName = WW_file.Replace("S0010_UPROFVIEW", "S0025_PROFMVIEW")

                    'Case "S0011_UPROFXLS"
                Case "S0026_PROFMXLS"
                    fileName = WW_file.Replace("S0011_UPROFXLS", "S0026_PROFMXLS")

            End Select

            '名称変更
            fileName = fileName.Replace(WW_InPARA_TBLNAME, WW_InPARA_TBLNAME & "Changing")

            '入力ファイル検索
            Dim sr As New System.IO.StreamReader(WW_file, System.Text.Encoding.GetEncoding("utf-8"))
            Dim sw As New System.IO.StreamWriter(fileName, False, System.Text.Encoding.GetEncoding("unicode"))

            Dim WW_InFile_Field As List(Of String)
            Dim WW_InFile_FieldHead As List(Of String)
            Dim WW_InFile_FieldValue As List(Of String)
            Dim WW_InFile_Index As List(Of String)
            Dim WW_Linecnt As Integer = 0
            WW_InFile_Field = New List(Of String)
            WW_InFile_FieldHead = New List(Of String)
            WW_InFile_FieldValue = New List(Of String)
            WW_InFile_Index = New List(Of String)

            'Dim AryList As ArrayList = New ArrayList
            Dim AryList As List(Of String()) = New List(Of String())
            Dim WW_Buff As String = ""
            Dim AppFlg As Boolean = False
            Dim TgtFlg As Boolean = True
            Dim headFlg As Boolean = False

            Try
                '■File情報をすべて読み込む
                While (Not sr.EndOfStream)
                    WW_InFile_FieldValue = New List(Of String)

                    AppFlg = False

                    '10000件を超える度に出力する
                    If WW_Linecnt >= breakCnt AndAlso (WW_Linecnt Mod breakCnt) = 0 Then

                        'ファイル作成
                        If AryList.Count > 0 Then

                            'TABLEフォルダーに抽出データファイルを出力（テーブル名.dat)
                            Dim WriteStr As String = ""

                            Try

                                If Not headFlg Then

                                    'DATヘッダーデータ出力
                                    For i As Integer = 0 To WW_InFile_Field.Count - 1
                                        WriteStr = WriteStr & WW_InFile_Field.Item(i).ToString
                                        If (WW_InFile_Field.Count - 1) = i Then
                                            WriteStr = WriteStr & ControlChars.NewLine
                                        Else
                                            WriteStr = WriteStr & ControlChars.Tab
                                        End If
                                    Next
                                    'DAT Line出力
                                    sw.Write(WriteStr)

                                    headFlg = True

                                End If

                                'DAT明細データ出力
                                For j As Integer = 0 To AryList.Count - 1
                                    WriteStr = ""
                                    For k As Integer = 0 To AryList(j).Count - 1
                                        WriteStr = WriteStr & AryList(j)(k).ToString
                                        If (AryList(j).Count - 1) = k Then
                                            WriteStr = WriteStr & ControlChars.NewLine
                                        Else
                                            WriteStr = WriteStr & ControlChars.Tab
                                        End If
                                    Next
                                    'DAT Line出力
                                    sw.Write(WriteStr)
                                Next


                            Catch ex As System.SystemException
                                '閉じる
                                sw.Close()
                                sw.Dispose()

                                CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
                                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                                CS0054LOGWrite_bat.INFPOSI = WW_InPARA_TBLNAME & " FILE OUTPUT ERR"    '
                                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                                CS0054LOGWrite_bat.TEXT = ex.ToString
                                CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
                                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                                Environment.Exit(100)

                            End Try

                        End If

                        AryList = New List(Of String())

                    End If

                    '○フィールドデータ切り出し
                    WW_Buff = sr.ReadLine()
                    Do
                        If WW_Linecnt = 0 Then
                            'ヘッダー行(フィールド名）取得＆チェック
                            WW_InFile_Field.Add(Mid(WW_Buff, 1, InStr(WW_Buff, ControlChars.Tab) - 1))
                            WW_InFile_FieldHead.Add(Mid(WW_Buff, 1, InStr(WW_Buff, ControlChars.Tab) - 1))
                            WW_Buff = Mid(WW_Buff, InStr(WW_Buff, ControlChars.Tab) + 1, 8000)
                            If InStr(WW_Buff, ControlChars.Tab) = 0 And WW_Buff <> "" Then
                                WW_InFile_Field.Add(WW_Buff)
                                WW_InFile_FieldHead.Add(WW_Buff)
                            End If
                        Else
                            'データ行取得
                            WW_InFile_FieldValue.Add(Mid(WW_Buff, 1, InStr(WW_Buff, vbTab) - 1))
                            WW_Buff = Mid(WW_Buff, InStr(WW_Buff, ControlChars.Tab) + 1, 8000)
                            If InStr(WW_Buff, ControlChars.Tab) = 0 And WW_Buff <> "" Then
                                WW_InFile_FieldValue.Add(WW_Buff)
                            End If
                        End If
                    Loop Until InStr(WW_Buff, ControlChars.Tab) = 0

                    '変換処理
                    Dim WW_CONV As Integer = -1
                    Select Case WW_InPARA_TBLNAME

                        Case "L0001_TOKEI"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("NACPRODUCTCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If

                                WW_InFile_Field.Insert(41, "NACPRODUCTCODE")                               '品名コード
                                WW_InFile_Field.Insert(155, "PAYWWORKTIME")                                '所定内時間
                                WW_InFile_Field.Insert(161, "PAYSDAIWORKTIME")                             '日曜代休出勤
                                WW_InFile_Field.Insert(162, "PAYSDAINIGHTTIME")                            '日曜代休深夜
                                WW_InFile_Field.Insert(165, "PAYHDAIWORKTIME")                             '代休出勤
                                WW_InFile_Field.Insert(166, "PAYHDAINIGHTTIME")                            '代休深夜
                                WW_InFile_Field.Insert(169, "PAYNENMATUNISSU")                             '年末出勤日数
                                WW_InFile_Field.Insert(189, "PAYHAYADETIME")                               '早出時間
                                WW_InFile_Field.Insert(190, "PAYHAISOTIME")                                '配送時間
                                WW_InFile_Field.Insert(191, "PAYSHACHUHAKNISSU")                           '車中泊日数
                                WW_InFile_Field.Insert(192, "PAYMODELDISTANCE")                            'モデル距離
                                WW_InFile_Field.Insert(193, "PAYJIKYUSHATIME")                             '時給者時間
                                WW_InFile_Field.Insert(194, "PAYJYOMUTIME")                                '乗務時間
                                WW_InFile_Field.Insert(195, "PAYHWORKNISSU")                               '休日出勤日数
                                WW_InFile_Field.Insert(196, "PAYKAITENCNT")                                '回転数
                                WW_InFile_Field.Insert(197, "PAYSENJYOCNT")                                '洗浄回数
                                WW_InFile_Field.Insert(198, "PAYUNLOADADDCNT1")                            '危険物荷卸回数1
                                WW_InFile_Field.Insert(199, "PAYUNLOADADDCNT2")                            '危険物荷卸回数2
                                WW_InFile_Field.Insert(200, "PAYUNLOADADDCNT3")                            '危険物荷卸回数3
                                WW_InFile_Field.Insert(201, "PAYUNLOADADDCNT4")                            '危険物荷卸回数4
                                WW_InFile_Field.Insert(202, "PAYSHORTDISTANCE1")                           '短距離手当1
                                WW_InFile_Field.Insert(203, "PAYSHORTDISTANCE2")                           '短距離手当2
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(41, "")
                                WW_InFile_FieldValue.Insert(155, "0")                                   '所定内時間
                                WW_InFile_FieldValue.Insert(161, "0")                                   '日曜代休出勤
                                WW_InFile_FieldValue.Insert(162, "0")                                   '日曜代休深夜
                                WW_InFile_FieldValue.Insert(165, "0")                                   '代休出勤
                                WW_InFile_FieldValue.Insert(166, "0")                                   '代休深夜
                                WW_InFile_FieldValue.Insert(169, "0")                                   '年末出勤日数
                                WW_InFile_FieldValue.Insert(189, "0")                                   '早出時間
                                WW_InFile_FieldValue.Insert(190, "0")                                   '配送時間
                                WW_InFile_FieldValue.Insert(191, "0")                                   '車中泊日数
                                WW_InFile_FieldValue.Insert(192, "0")                                   'モデル距離
                                WW_InFile_FieldValue.Insert(193, "0")                                   '時給者時間
                                WW_InFile_FieldValue.Insert(194, "0")                                   '乗務時間
                                WW_InFile_FieldValue.Insert(195, "0")                                   '休日出勤日数
                                WW_InFile_FieldValue.Insert(196, "0")                                   '回転数
                                WW_InFile_FieldValue.Insert(197, "0")                                   '洗浄回数
                                WW_InFile_FieldValue.Insert(198, "0")                                   '危険物荷卸回数1
                                WW_InFile_FieldValue.Insert(199, "0")                                   '危険物荷卸回数2
                                WW_InFile_FieldValue.Insert(200, "0")                                   '危険物荷卸回数3
                                WW_InFile_FieldValue.Insert(201, "0")                                   '危険物荷卸回数4
                                WW_InFile_FieldValue.Insert(202, "0")                                   '短距離手当1
                                WW_InFile_FieldValue.Insert(203, "0")                                   '短距離手当2

                                Dim OilTypeIdx As Integer = 0
                                Dim Product1Idx As Integer = 0
                                Dim Product2Idx As Integer = 0
                                Dim NacProductIdx As Integer = 0
                                OilTypeIdx = WW_InFile_Field.IndexOf("NACOILTYPE")
                                Product1Idx = WW_InFile_Field.IndexOf("NACPRODUCT1")
                                Product2Idx = WW_InFile_Field.IndexOf("NACPRODUCT2")
                                NacProductIdx = WW_InFile_Field.IndexOf("NACPRODUCTCODE")
                                If WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> vbNullChar AndAlso
                                        WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(NacProductIdx) = CompCode & WW_InFile_FieldValue(OilTypeIdx).Trim.ToString & WW_InFile_FieldValue(Product1Idx).Trim.ToString & WW_InFile_FieldValue(Product2Idx).Trim.ToString      '品名コード
                                End If

                                Dim ActShabanIdx As Integer = 0
                                ActShabanIdx = WW_InFile_Field.IndexOf("ACTSHABAN")
                                If WW_InFile_FieldValue(ActShabanIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ActShabanIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ActShabanIdx) = CompCode0 & WW_InFile_FieldValue(ActShabanIdx).Trim.ToString
                                End If

                                Dim Shaban1Idx As Integer = 0
                                Shaban1Idx = WW_InFile_Field.IndexOf("NACTSHABAN1")
                                If WW_InFile_FieldValue(Shaban1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban1Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban1Idx) = CompCode0 & WW_InFile_FieldValue(Shaban1Idx).Trim.ToString
                                End If

                                Dim Shaban2Idx As Integer = 0
                                Shaban2Idx = WW_InFile_Field.IndexOf("NACTSHABAN2")
                                If WW_InFile_FieldValue(Shaban2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban2Idx) = CompCode0 & WW_InFile_FieldValue(Shaban2Idx).Trim.ToString
                                End If

                                Dim Shaban3Idx As Integer = 0
                                Shaban3Idx = WW_InFile_Field.IndexOf("NACTSHABAN3")
                                If WW_InFile_FieldValue(Shaban3Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban3Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban3Idx) = CompCode0 & WW_InFile_FieldValue(Shaban3Idx).Trim.ToString
                                End If
                            End If

                        Case "L0003_SUMMARYN"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                            Else
                                '明細

                                Dim Shaban1Idx As Integer = 0
                                Shaban1Idx = WW_InFile_Field.IndexOf("NACTSHABAN1")

                                '既に変換済みの場合、スルーする
                                If WW_InFile_FieldValue(Shaban1Idx).Length > 5 And Mid(WW_InFile_FieldValue(Shaban1Idx), 1, 3) = CompCode0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If

                                If WW_InFile_FieldValue(Shaban1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban1Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban1Idx) = CompCode0 & WW_InFile_FieldValue(Shaban1Idx).Trim.ToString
                                End If

                                Dim Shaban2Idx As Integer = 0
                                Shaban2Idx = WW_InFile_Field.IndexOf("NACTSHABAN2")
                                If WW_InFile_FieldValue(Shaban2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban2Idx) = CompCode0 & WW_InFile_FieldValue(Shaban2Idx).Trim.ToString
                                End If

                                Dim Shaban3Idx As Integer = 0
                                Shaban3Idx = WW_InFile_Field.IndexOf("NACTSHABAN3")
                                If WW_InFile_FieldValue(Shaban3Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban3Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban3Idx) = CompCode0 & WW_InFile_FieldValue(Shaban3Idx).Trim.ToString
                                End If

                                Dim KeyShaban1Idx As Integer = 0
                                KeyShaban1Idx = WW_InFile_Field.IndexOf("KEYTSHABAN1")
                                If WW_InFile_FieldValue(KeyShaban1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(KeyShaban1Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(KeyShaban1Idx) = Mid(WW_InFile_FieldValue(KeyShaban1Idx), 1, 1) & CompCode0 & Mid(WW_InFile_FieldValue(KeyShaban1Idx).Trim, 2)
                                End If

                                Dim KeyShaban2Idx As Integer = 0
                                KeyShaban2Idx = WW_InFile_Field.IndexOf("KEYTSHABAN2")
                                If WW_InFile_FieldValue(KeyShaban2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(KeyShaban2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(KeyShaban2Idx) = Mid(WW_InFile_FieldValue(KeyShaban2Idx), 1, 1) & CompCode0 & Mid(WW_InFile_FieldValue(KeyShaban2Idx).Trim, 2)
                                End If

                                Dim KeyShaban3Idx As Integer = 0
                                KeyShaban3Idx = WW_InFile_Field.IndexOf("KEYTSHABAN3")
                                If WW_InFile_FieldValue(KeyShaban3Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(KeyShaban3Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(KeyShaban3Idx) = Mid(WW_InFile_FieldValue(KeyShaban3Idx), 1, 1) & CompCode0 & Mid(WW_InFile_FieldValue(KeyShaban3Idx).Trim, 2)
                                End If
                            End If

                        Case "L0004_SUMMARYK"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("PAYHAYADETIME")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(248, "PAYHAYADETIME")                               '早出時間
                                WW_InFile_Field.Insert(249, "PAYHAISOTIME")                                '配送時間
                                WW_InFile_Field.Insert(250, "PAYNENMATUNISSU")                             '年末出勤日数
                                WW_InFile_Field.Insert(251, "PAYSHACHUHAKNISSU")                           '車中泊日数
                                WW_InFile_Field.Insert(252, "PAYMODELDISTANCE")                            'モデル距離
                                WW_InFile_Field.Insert(253, "PAYJIKYUSHATIME")                             '時給者時間
                                WW_InFile_Field.Insert(254, "PAYHDAIWORKTIME")                             '代休出勤
                                WW_InFile_Field.Insert(255, "PAYHDAINIGHTTIME")                            '代休深夜
                                WW_InFile_Field.Insert(256, "PAYSDAIWORKTIME")                             '日曜代休出勤
                                WW_InFile_Field.Insert(257, "PAYSDAINIGHTTIME")                            '日曜代休深夜
                                WW_InFile_Field.Insert(258, "PAYWWORKTIME")                                '所定内時間
                                WW_InFile_Field.Insert(259, "PAYJYOMUTIME")                                '乗務時間
                                WW_InFile_Field.Insert(260, "PAYHWORKNISSU")                               '休日出勤日数
                                WW_InFile_Field.Insert(261, "PAYKAITENCNT")                                '回転数
                                WW_InFile_Field.Insert(262, "PAYSENJYOCNT")                                '洗浄回数
                                WW_InFile_Field.Insert(263, "PAYUNLOADADDCNT1")                            '危険物荷卸回数1
                                WW_InFile_Field.Insert(264, "PAYUNLOADADDCNT2")                            '危険物荷卸回数2
                                WW_InFile_Field.Insert(265, "PAYUNLOADADDCNT3")                            '危険物荷卸回数3
                                WW_InFile_Field.Insert(266, "PAYUNLOADADDCNT4")                            '危険物荷卸回数4
                                WW_InFile_Field.Insert(267, "PAYSHORTDISTANCE1")                           '短距離手当1
                                WW_InFile_Field.Insert(268, "PAYSHORTDISTANCE2")                           '短距離手当2
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(248, "0")                                   '早出時間
                                WW_InFile_FieldValue.Insert(249, "0")                                   '配送時間
                                WW_InFile_FieldValue.Insert(250, "0")                                   '年末出勤日数
                                WW_InFile_FieldValue.Insert(251, "0")                                   '車中泊日数
                                WW_InFile_FieldValue.Insert(252, "0")                                   'モデル距離
                                WW_InFile_FieldValue.Insert(253, "0")                                   '時給者時間
                                WW_InFile_FieldValue.Insert(254, "0")                                   '代休出勤
                                WW_InFile_FieldValue.Insert(255, "0")                                   '代休深夜
                                WW_InFile_FieldValue.Insert(256, "0")                                   '日曜代休出勤
                                WW_InFile_FieldValue.Insert(257, "0")                                   '日曜代休深夜
                                WW_InFile_FieldValue.Insert(258, "0")                                   '所定内時間
                                WW_InFile_FieldValue.Insert(259, "0")                                   '乗務時間
                                WW_InFile_FieldValue.Insert(260, "0")                                   '休日出勤日数
                                WW_InFile_FieldValue.Insert(261, "0")                                   '回転数
                                WW_InFile_FieldValue.Insert(262, "0")                                   '洗浄回数
                                WW_InFile_FieldValue.Insert(263, "0")                                   '危険物荷卸回数1
                                WW_InFile_FieldValue.Insert(264, "0")                                   '危険物荷卸回数2
                                WW_InFile_FieldValue.Insert(265, "0")                                   '危険物荷卸回数3
                                WW_InFile_FieldValue.Insert(266, "0")                                   '危険物荷卸回数4
                                WW_InFile_FieldValue.Insert(267, "0")                                   '短距離手当1
                                WW_InFile_FieldValue.Insert(268, "0")                                   '短距離手当2
                            End If

                        Case "L0005_SUMMARYY"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                            Else
                                '明細
                                Dim Shaban1Idx As Integer = 0
                                Shaban1Idx = WW_InFile_Field.IndexOf("NACTSHABAN1")

                                '既に変換済みの場合、スルーする
                                If WW_InFile_FieldValue(Shaban1Idx).Length > 5 And Mid(WW_InFile_FieldValue(Shaban1Idx), 1, 3) = CompCode0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If

                                If WW_InFile_FieldValue(Shaban1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban1Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban1Idx) = CompCode0 & WW_InFile_FieldValue(Shaban1Idx).Trim.ToString
                                End If

                                Dim Shaban2Idx As Integer = 0
                                Shaban2Idx = WW_InFile_Field.IndexOf("NACTSHABAN2")
                                If WW_InFile_FieldValue(Shaban2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban2Idx) = CompCode0 & WW_InFile_FieldValue(Shaban2Idx).Trim.ToString
                                End If

                                Dim Shaban3Idx As Integer = 0
                                Shaban3Idx = WW_InFile_Field.IndexOf("NACTSHABAN3")
                                If WW_InFile_FieldValue(Shaban3Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Shaban3Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(Shaban3Idx) = CompCode0 & WW_InFile_FieldValue(Shaban3Idx).Trim.ToString
                                End If

                                Dim KeyShaban1Idx As Integer = 0
                                KeyShaban1Idx = WW_InFile_Field.IndexOf("KEYTSHABAN1")
                                If WW_InFile_FieldValue(KeyShaban1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(KeyShaban1Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(KeyShaban1Idx) = Mid(WW_InFile_FieldValue(KeyShaban1Idx), 1, 1) & CompCode0 & Mid(WW_InFile_FieldValue(KeyShaban1Idx).Trim, 2)
                                End If

                                Dim KeyShaban2Idx As Integer = 0
                                KeyShaban2Idx = WW_InFile_Field.IndexOf("KEYTSHABAN2")
                                If WW_InFile_FieldValue(KeyShaban2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(KeyShaban2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(KeyShaban2Idx) = Mid(WW_InFile_FieldValue(KeyShaban2Idx), 1, 1) & CompCode0 & Mid(WW_InFile_FieldValue(KeyShaban2Idx).Trim, 2)
                                End If

                                Dim KeyShaban3Idx As Integer = 0
                                KeyShaban3Idx = WW_InFile_Field.IndexOf("KEYTSHABAN3")
                                If WW_InFile_FieldValue(KeyShaban3Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(KeyShaban3Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(KeyShaban3Idx) = Mid(WW_InFile_FieldValue(KeyShaban3Idx), 1, 1) & CompCode0 & Mid(WW_InFile_FieldValue(KeyShaban3Idx).Trim, 2)
                                End If
                            End If

                        Case "MA002_SHARYOA"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                            Else
                                '明細
                                Dim ShabanIdx As Integer = 0
                                ShabanIdx = WW_InFile_Field.IndexOf("TSHABAN")

                                '既に変換済みの場合、スルーする
                                If Trim(WW_InFile_FieldValue(ShabanIdx)).Length > 5 And Mid(WW_InFile_FieldValue(ShabanIdx), 1, 3) = CompCode0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If

                                If WW_InFile_FieldValue(ShabanIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanIdx) = CompCode0 & WW_InFile_FieldValue(ShabanIdx).Trim.ToString
                                End If
                            End If

                        Case "MA003_SHARYOB"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                            Else
                                '明細
                                Dim ShabanIdx As Integer = 0
                                ShabanIdx = WW_InFile_Field.IndexOf("TSHABAN")

                                '既に変換済みの場合、スルーする
                                If Trim(WW_InFile_FieldValue(ShabanIdx)).Length > 5 And Mid(WW_InFile_FieldValue(ShabanIdx), 1, 3) = CompCode0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If

                                If WW_InFile_FieldValue(ShabanIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanIdx) = CompCode0 & WW_InFile_FieldValue(ShabanIdx).Trim.ToString
                                End If
                            End If

                        Case "MA004_SHARYOC"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                            Else
                                '明細
                                Dim ShabanIdx As Integer = 0
                                ShabanIdx = WW_InFile_Field.IndexOf("TSHABAN")

                                '既に変換済みの場合、スルーする
                                If Trim(WW_InFile_FieldValue(ShabanIdx)).Length > 5 And Mid(WW_InFile_FieldValue(ShabanIdx), 1, 3) = CompCode0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If

                                If WW_InFile_FieldValue(ShabanIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanIdx) = CompCode0 & WW_InFile_FieldValue(ShabanIdx).Trim.ToString
                                End If
                            End If

                        Case "MA006_SHABANORG"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("MANGOWNCONT")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If

                                WW_InFile_Field.Insert(25, "MANGOWNCONT")                              '契約区分
                                WW_InFile_Field.Insert(26, "JSRSHABAN")                                'JSR車番コード
                            Else
                                '明細
                                Dim ShabanFIdx As Integer = 0
                                ShabanFIdx = WW_InFile_Field.IndexOf("TSHABANF")
                                If WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanFIdx) = CompCode0 & WW_InFile_FieldValue(ShabanFIdx).Trim.ToString
                                End If

                                Dim ShabanBIdx As Integer = 0
                                ShabanBIdx = WW_InFile_Field.IndexOf("TSHABANB")
                                If WW_InFile_FieldValue(ShabanBIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanBIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanBIdx) = CompCode0 & WW_InFile_FieldValue(ShabanBIdx).Trim.ToString
                                End If

                                Dim ShabanB2Idx As Integer = 0
                                ShabanB2Idx = WW_InFile_Field.IndexOf("TSHABANB2")
                                If WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanB2Idx) = CompCode0 & WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString
                                End If

                                WW_InFile_FieldValue.Insert(25, "")                                   '契約区分
                                WW_InFile_FieldValue.Insert(26, "")                                   'JSR車番コード
                            End If

                        Case "MB002_STAFFORG"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("JSRSTAFFCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(4, "JSRSTAFFCODE")                             'JSR従業員コード
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(4, "")                                   'JSR従業員コード
                            End If

                        Case "MC001_FIXVALUE"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("SYSTEMKEYFLG")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(12, "SYSTEMKEYFLG")                             'システムキーフラグ
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(12, "1")                                   'システムキーフラグ
                            End If

                        Case "MC002_TORIHIKISAKI"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("CAMPCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(0, "CAMPCODE")                                 '会社コード
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(0, CompCode)                              '会社コード
                            End If

                        Case "MC007_TODKORG"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("JSRTODOKECODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(8, "JSRTODOKECODE")                            'JSR届先コード
                                WW_InFile_Field.Insert(9, "SHUKABASHO")                               '出荷場所
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(8, "")                                    'JSR届先コード
                                WW_InFile_FieldValue.Insert(9, " ")                                   '出荷場所
                            End If

                        'Case "MC004_PRODUCT"
                        Case "MD001_PRODUCT"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("CAMPCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(0, "CAMPCODE")                                 '会社コード
                                WW_InFile_Field.Insert(1, "PRODUCTCODE")                              '品名コード
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(0, CompCode)                              '会社コード
                                WW_InFile_FieldValue.Insert(1, "")                                    '品名コード

                                Dim OilTypeIdx As Integer = 0
                                Dim Product1Idx As Integer = 0
                                Dim Product2Idx As Integer = 0
                                Dim ProductCodeIdx As Integer = 0
                                OilTypeIdx = WW_InFile_Field.IndexOf("OILTYPE")
                                Product1Idx = WW_InFile_Field.IndexOf("PRODUCT1")
                                Product2Idx = WW_InFile_Field.IndexOf("PRODUCT2")
                                ProductCodeIdx = WW_InFile_Field.IndexOf("PRODUCTCODE")
                                If WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCodeIdx) = CompCode & WW_InFile_FieldValue(OilTypeIdx).Trim.ToString & WW_InFile_FieldValue(Product1Idx).Trim.ToString & WW_InFile_FieldValue(Product2Idx).Trim.ToString      '品名コード
                                End If
                            End If

                        'Case "MC005_PRODORG"
                        Case "MD002_PRODORG"
                            Dim OilTypeIdx As Integer = 0
                            Dim Product1Idx As Integer = 0
                            Dim Product2Idx As Integer = 0
                            OilTypeIdx = WW_InFile_FieldHead.IndexOf("OILTYPE")
                            Product1Idx = WW_InFile_FieldHead.IndexOf("PRODUCT1")
                            Product2Idx = WW_InFile_FieldHead.IndexOf("PRODUCT2")

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("PRODUCTCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.RemoveAt(Product2Idx)                           '品名２
                                WW_InFile_Field.RemoveAt(Product1Idx)                           '品名１
                                WW_InFile_Field.RemoveAt(OilTypeIdx)                            '油種

                                WW_InFile_Field.Insert(2, "PRODUCTCODE")                              '品名コード
                                WW_InFile_Field.Insert(19, "JSRPRODUCT")                              'JSR品名コード
                                WW_InFile_Field.Insert(20, "UNLOADADDTANKA")                          '荷卸時加算単価
                                WW_InFile_Field.Insert(21, "LOADINGTANKA")                            '積込単価
                            Else
                                '明細

                                If WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue.Insert(2, CompCode & WW_InFile_FieldValue(OilTypeIdx).Trim.ToString & WW_InFile_FieldValue(Product1Idx).Trim.ToString & WW_InFile_FieldValue(Product2Idx).Trim.ToString)      '品名コード
                                Else
                                    WW_InFile_FieldValue.Insert(2, "")
                                End If

                                WW_InFile_FieldValue.RemoveAt(Product2Idx)                      '品名２
                                WW_InFile_FieldValue.RemoveAt(Product1Idx)                      '品名１
                                WW_InFile_FieldValue.RemoveAt(OilTypeIdx)                       '油種

                                WW_InFile_FieldValue.Insert(19, "")                                    'JSR品名コード
                                WW_InFile_FieldValue.Insert(20, "0")                                   '荷卸時加算単価
                                WW_InFile_FieldValue.Insert(21, "0")                                   '積込単価

                            End If

                        Case "S0004_USER"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("CAMPROLE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(8, "CAMPROLE")                                 '会社権限
                                WW_InFile_Field.Insert(9, "MAPROLE")                                  '更新権限
                                WW_InFile_Field.Insert(10, "ORGROLE")                                 '部署権限
                                WW_InFile_Field.Insert(11, "VIEWPROFID")                              '画面プロファイルID
                                WW_InFile_Field.Insert(12, "RPRTPROFID")                              '帳票プロファイルID
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(8, "")                                    '会社権限
                                WW_InFile_FieldValue.Insert(9, "")                                    '更新権限
                                WW_InFile_FieldValue.Insert(10, "")                                   '部署権限
                                WW_InFile_FieldValue.Insert(11, "")                                   '画面プロファイルID
                                WW_InFile_FieldValue.Insert(12, "")                                   '帳票プロファイルID
                            End If

                        'Case "S0007_UPROFVARI"
                        Case "S0023_PROFMVARI"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("PROFID")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                Dim UserIdIdx As Integer = 0
                                Dim TitolKbnIdx As Integer = 0
                                Dim TitolIdx As Integer = 0
                                UserIdIdx = WW_InFile_Field.IndexOf("USERID")
                                TitolKbnIdx = WW_InFile_Field.IndexOf("TITOLKBN")
                                TitolIdx = WW_InFile_Field.IndexOf("TITOL")

                                If WW_InFile_Field(UserIdIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(UserIdIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(UserIdIdx) = "PROFID"
                                End If

                                If WW_InFile_Field(TitolKbnIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(TitolKbnIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(TitolKbnIdx) = "TITLEKBN"
                                End If

                                If WW_InFile_Field(TitolIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(TitolIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(TitolIdx) = "TITLENAMES"
                                End If
                            Else
                                '明細

                                'ユーザー、プロフID変換
                                Dim ProfIdIdx As Integer = 0
                                ProfIdIdx = WW_InFile_Field.IndexOf("PROFID")

                                If WW_InFile_Field(ProfIdIdx).ToString <> "" Then
                                    For i As Integer = 0 To PrfList.Count - 1
                                        If DirectCast(PrfList(i), String())(0) = WW_InFile_FieldValue(ProfIdIdx).Trim.ToString Then

                                            WW_InFile_FieldValue(ProfIdIdx) = DirectCast(PrfList(i), String())(1)
                                            'PROFID変換対象
                                            AppFlg = True
                                            Exit For

                                        End If
                                    Next
                                End If
                            End If

                        'Case "S0010_UPROFVIEW"
                        Case "S0025_PROFMVIEW"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("CAMPCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(0, "CAMPCODE")                                 '会社コード
                                WW_InFile_Field.Insert(14, "PREFIX")                                   '接頭句
                                WW_InFile_Field.Insert(15, "SUFFIX")                                   '接尾句
                                WW_InFile_Field.Insert(19, "SORTKBN")                                  '昇降区分
                                WW_InFile_Field.Insert(21, "WIDTH")                                    '横幅
                                WW_InFile_Field.Insert(22, "OBJECTTYPE")                               'オブジェクトタイプ
                                WW_InFile_Field.Insert(23, "FORMATTYPE")                               'フォーマットタイプ
                                WW_InFile_Field.Insert(24, "FORMATVALUE")                              'フォーマット書式
                                WW_InFile_Field.Insert(25, "FIXCOL")                                   '固定列
                                WW_InFile_Field.Insert(26, "REQUIRED")                                 '入力必須
                                WW_InFile_Field.Insert(27, "COLORSET")                                 '色設定
                                WW_InFile_Field.Insert(28, "ADDEVENT1")                                '追加イベント１
                                WW_InFile_Field.Insert(29, "ADDFUNC1")                                 '追加ファンクション１
                                WW_InFile_Field.Insert(30, "ADDEVENT2")                                '追加イベント２
                                WW_InFile_Field.Insert(31, "ADDFUNC2")                                 '追加ファンクション２
                                WW_InFile_Field.Insert(32, "ADDEVENT3")                                '追加イベント３
                                WW_InFile_Field.Insert(33, "ADDFUNC3")                                 '追加ファンクション３
                                WW_InFile_Field.Insert(34, "ADDEVENT4")                                '追加イベント４
                                WW_InFile_Field.Insert(35, "ADDFUNC4")                                 '追加ファンクション４
                                WW_InFile_Field.Insert(36, "ADDEVENT5")                                '追加イベント５
                                WW_InFile_Field.Insert(37, "ADDFUNC5")                                 '追加ファンクション５

                                Dim UserIdIdx As Integer = 0
                                Dim TitolKbnIdx As Integer = 0
                                Dim TabIdx As Integer = 0
                                Dim SeqIdx As Integer = 0
                                Dim PojitionIdx As Integer = 0
                                Dim NamesIdx As Integer = 0
                                Dim NamelIdx As Integer = 0
                                Dim SortIdx As Integer = 0
                                Dim HdkbnIdx As Integer = 0
                                UserIdIdx = WW_InFile_Field.IndexOf("USERID")
                                TitolKbnIdx = WW_InFile_Field.IndexOf("TITOLKBN")
                                TabIdx = WW_InFile_Field.IndexOf("TAB")
                                SeqIdx = WW_InFile_Field.IndexOf("SEQ")
                                PojitionIdx = WW_InFile_Field.IndexOf("POJITION")
                                NamesIdx = WW_InFile_Field.IndexOf("NAMES")
                                NamelIdx = WW_InFile_Field.IndexOf("NAMEL")
                                SortIdx = WW_InFile_Field.IndexOf("SORT")
                                HdkbnIdx = WW_InFile_Field.IndexOf("HDKBN")

                                If WW_InFile_Field(UserIdIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(UserIdIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(UserIdIdx) = "PROFID"
                                End If

                                If WW_InFile_Field(TitolKbnIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(TitolKbnIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(TitolKbnIdx) = "TITLEKBN"
                                End If

                                If WW_InFile_Field(TabIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(TabIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(TabIdx) = "TABID"
                                End If

                                If WW_InFile_Field(SeqIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(SeqIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(SeqIdx) = "POSICOL"
                                End If

                                If WW_InFile_Field(PojitionIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(PojitionIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(PojitionIdx) = "POSIROW"
                                End If

                                If WW_InFile_Field(NamesIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(NamesIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(NamesIdx) = "FIELDNAMES"
                                End If

                                If WW_InFile_Field(NamelIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(NamelIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(NamelIdx) = "FIELDNAMEL"
                                End If

                                If WW_InFile_Field(SortIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(SortIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(SortIdx) = "SORTORDER"
                                End If

                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(0, CompCode)                              '会社コード
                                WW_InFile_FieldValue.Insert(14, "")                                    '接頭句
                                WW_InFile_FieldValue.Insert(15, "")                                    '接尾句
                                WW_InFile_FieldValue.Insert(19, "")                                    '昇降区分
                                WW_InFile_FieldValue.Insert(21, "0")                                   '横幅
                                WW_InFile_FieldValue.Insert(22, "")                                    'オブジェクトタイプ
                                WW_InFile_FieldValue.Insert(23, "")                                    'フォーマットタイプ
                                WW_InFile_FieldValue.Insert(24, "")                                    'フォーマット書式
                                WW_InFile_FieldValue.Insert(25, "")                                    '固定列
                                WW_InFile_FieldValue.Insert(26, "")                                    '入力必須
                                WW_InFile_FieldValue.Insert(27, "")                                    '色設定
                                WW_InFile_FieldValue.Insert(28, "")                                    '追加イベント１
                                WW_InFile_FieldValue.Insert(29, "")                                    '追加ファンクション１
                                WW_InFile_FieldValue.Insert(30, "")                                    '追加イベント２
                                WW_InFile_FieldValue.Insert(31, "")                                    '追加ファンクション２
                                WW_InFile_FieldValue.Insert(32, "")                                    '追加イベント３
                                WW_InFile_FieldValue.Insert(33, "")                                    '追加ファンクション３
                                WW_InFile_FieldValue.Insert(34, "")                                    '追加イベント４
                                WW_InFile_FieldValue.Insert(35, "")                                    '追加ファンクション４
                                WW_InFile_FieldValue.Insert(36, "")                                    '追加イベント５
                                WW_InFile_FieldValue.Insert(37, " ")                                   '追加ファンクション５

                                'ユーザー、プロフID変換
                                Dim ProfIdIdx As Integer = 0
                                ProfIdIdx = WW_InFile_Field.IndexOf("PROFID")

                                If WW_InFile_Field(ProfIdIdx).ToString <> vbNullChar Then
                                    For i As Integer = 0 To PrfList.Count - 1
                                        If DirectCast(PrfList(i), String())(0) = WW_InFile_FieldValue(ProfIdIdx).Trim.ToString Then

                                            WW_InFile_FieldValue(ProfIdIdx) = DirectCast(PrfList(i), String())(1)
                                            'PROFID変換対象
                                            AppFlg = True
                                            Exit For

                                        End If
                                    Next
                                End If

                                Dim PosicolIdx As Integer = 0
                                Dim PosicolStr As String = "0"
                                Dim PosirowIdx As Integer = 0
                                Dim PosirowStr As String = "0"
                                Dim HdkbnIdx As Integer = 0
                                Dim EffectStr As String = "N"
                                Dim EffectIdx As Integer = 0
                                Dim LengthIdx As Integer = 0
                                Dim widthIdx As Integer = 0
                                PosicolIdx = WW_InFile_Field.IndexOf("POSICOL")
                                PosirowIdx = WW_InFile_Field.IndexOf("POSIROW")
                                HdkbnIdx = WW_InFile_Field.IndexOf("HDKBN")
                                EffectIdx = WW_InFile_Field.IndexOf("EFFECT")
                                LengthIdx = WW_InFile_Field.IndexOf("LENGTH")
                                widthIdx = WW_InFile_Field.IndexOf("WIDTH")
                                PosicolStr = WW_InFile_FieldValue(PosicolIdx).ToString                  'POJITION
                                PosirowStr = WW_InFile_FieldValue(PosirowIdx).ToString                  'SEQ

                                If WW_InFile_FieldValue(HdkbnIdx).ToString = "H" Then
                                    'Hの場合、POSIROW：POJITION　POSICOL：SEQ
                                    If PosicolStr = "" Then
                                        PosicolStr = "0"
                                    End If
                                    WW_InFile_FieldValue(PosirowIdx) = PosirowStr
                                    WW_InFile_FieldValue(PosicolIdx) = PosicolStr
                                    If Trim(WW_InFile_FieldValue(LengthIdx)) = "" Then
                                        WW_InFile_FieldValue(LengthIdx) = 0
                                    Else
                                        WW_InFile_FieldValue(LengthIdx) = CInt(WW_InFile_FieldValue(LengthIdx))
                                    End If
                                    WW_InFile_FieldValue(widthIdx) = WW_InFile_FieldValue(LengthIdx)
                                Else
                                    'Dの場合、POSICOL：SEQ　POSIROW：POJITION
                                    '    WW_InFile_FieldValue(PosirowIdx) = PosirowStr
                                    '    WW_InFile_FieldValue(PosicolIdx) = PosicolStr

                                    Select Case WW_InFile_FieldValue(PosirowIdx).ToString
                                        Case "L"
                                            PosicolStr = "1"
                                            EffectStr = "Y"
                                        Case "M"
                                            PosicolStr = "2"
                                            EffectStr = "Y"
                                        Case "R"
                                            PosicolStr = "3"
                                            EffectStr = "Y"
                                        Case ""
                                            PosicolStr = "0"
                                            EffectStr = "N"
                                        Case Else
                                            PosicolStr = WW_InFile_FieldValue(PosirowIdx).ToString
                                    End Select

                                    If Trim(WW_InFile_FieldValue(LengthIdx)) = "" Then
                                        WW_InFile_FieldValue(LengthIdx) = 0
                                    Else
                                        WW_InFile_FieldValue(LengthIdx) = CInt(WW_InFile_FieldValue(LengthIdx))
                                    End If

                                    WW_InFile_FieldValue(PosirowIdx) = WW_InFile_FieldValue(PosicolIdx).ToString
                                    WW_InFile_FieldValue(PosicolIdx) = PosicolStr
                                    WW_InFile_FieldValue(EffectIdx) = EffectStr
                                End If
                            End If

                        'Case "S0011_UPROFXLS"
                        Case "S0026_PROFMXLS"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("CAMPCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(0, "CAMPCODE")                                 '会社コード
                                WW_InFile_Field.Insert(17, "FORMATTYPE")                               'フォーマットタイプ

                                Dim UserIdIdx As Integer = 0
                                Dim TitolKbnIdx As Integer = 0
                                Dim FieldNameIdx As Integer = 0
                                Dim PosiXIdx As Integer = 0
                                Dim PosiYIdx As Integer = 0
                                Dim StructIdx As Integer = 0
                                Dim SortIdx As Integer = 0
                                UserIdIdx = WW_InFile_Field.IndexOf("USERID")
                                TitolKbnIdx = WW_InFile_Field.IndexOf("TITOLKBN")
                                FieldNameIdx = WW_InFile_Field.IndexOf("FIELDNAME")
                                PosiXIdx = WW_InFile_Field.IndexOf("POSIX")
                                PosiYIdx = WW_InFile_Field.IndexOf("POSIY")
                                StructIdx = WW_InFile_Field.IndexOf("STRUCT")
                                SortIdx = WW_InFile_Field.IndexOf("SORT")

                                If WW_InFile_Field(UserIdIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(UserIdIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(UserIdIdx) = "PROFID"
                                End If

                                If WW_InFile_Field(TitolKbnIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(TitolKbnIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(TitolKbnIdx) = "TITLEKBN"
                                End If

                                If WW_InFile_Field(FieldNameIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(FieldNameIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(FieldNameIdx) = "FIELDNAMES"
                                End If

                                If WW_InFile_Field(PosiXIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(PosiXIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(PosiXIdx) = "POSICOL"
                                End If

                                If WW_InFile_Field(PosiYIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(PosiYIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(PosiYIdx) = "POSIROW"
                                End If

                                If WW_InFile_Field(StructIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(StructIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(StructIdx) = "STRUCTCODE"
                                End If

                                If WW_InFile_Field(SortIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_Field(SortIdx).Trim.ToString <> "" Then
                                    WW_InFile_Field(SortIdx) = "SORTORDER"
                                End If

                            Else
                                '明細

                                WW_InFile_FieldValue.Insert(0, CompCode)                              '会社コード
                                WW_InFile_FieldValue.Insert(17, "")                                   'フォーマットタイプ

                                'ユーザー、プロフID変換
                                Dim ProfIdIdx As Integer = 0
                                ProfIdIdx = WW_InFile_Field.IndexOf("PROFID")

                                If WW_InFile_Field(ProfIdIdx).ToString <> "" Then
                                    For i As Integer = 0 To PrfList.Count - 1
                                        If DirectCast(PrfList(i), String())(0) = WW_InFile_FieldValue(ProfIdIdx).Trim.ToString Then

                                            WW_InFile_FieldValue(ProfIdIdx) = DirectCast(PrfList(i), String())(1)
                                            'PROFID変換対象
                                            AppFlg = True
                                            Exit For

                                        End If
                                    Next
                                End If

                            End If

                        Case "T0003_NIORDER"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("PRODUCTCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(34, "PRODUCTCODE")                              '品名コード
                            Else
                                WW_InFile_FieldValue.Insert(34, "")
                                '明細
                                Dim OilTypeIdx As Integer = 0
                                Dim Product1Idx As Integer = 0
                                Dim Product2Idx As Integer = 0
                                Dim ProductCodeIdx As Integer = 0
                                OilTypeIdx = WW_InFile_Field.IndexOf("OILTYPE")
                                Product1Idx = WW_InFile_Field.IndexOf("PRODUCT1")
                                Product2Idx = WW_InFile_Field.IndexOf("PRODUCT2")
                                ProductCodeIdx = WW_InFile_Field.IndexOf("PRODUCTCODE")
                                If WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCodeIdx) = CompCode & WW_InFile_FieldValue(OilTypeIdx).Trim.ToString & WW_InFile_FieldValue(Product1Idx).Trim.ToString & WW_InFile_FieldValue(Product2Idx).Trim.ToString      '品名コード
                                End If

                                Dim ShabanFIdx As Integer = 0
                                ShabanFIdx = WW_InFile_Field.IndexOf("TSHABANF")
                                If WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanFIdx) = CompCode0 & WW_InFile_FieldValue(ShabanFIdx).Trim.ToString
                                End If

                                Dim ShabanBIdx As Integer = 0
                                ShabanBIdx = WW_InFile_Field.IndexOf("TSHABANB")
                                If WW_InFile_FieldValue(ShabanBIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanBIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanBIdx) = CompCode0 & WW_InFile_FieldValue(ShabanBIdx).Trim.ToString
                                End If

                                Dim ShabanB2Idx As Integer = 0
                                ShabanB2Idx = WW_InFile_Field.IndexOf("TSHABANB2")
                                If WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanB2Idx) = CompCode0 & WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString
                                End If
                            End If

                        Case "T0004_HORDER"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("PRODUCTCODE")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(34, "PRODUCTCODE")                              '品名コード
                                WW_InFile_Field.Insert(63, "JXORDERID")                                'JXオーダー識別ID
                            Else
                                WW_InFile_FieldValue.Insert(34, "")
                                WW_InFile_FieldValue.Insert(63, "")
                                '明細
                                Dim OilTypeIdx As Integer = 0
                                Dim Product1Idx As Integer = 0
                                Dim Product2Idx As Integer = 0
                                Dim ProductCodeIdx As Integer = 0
                                OilTypeIdx = WW_InFile_Field.IndexOf("OILTYPE")
                                Product1Idx = WW_InFile_Field.IndexOf("PRODUCT1")
                                Product2Idx = WW_InFile_Field.IndexOf("PRODUCT2")
                                ProductCodeIdx = WW_InFile_Field.IndexOf("PRODUCTCODE")
                                If WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilTypeIdx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product1Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCodeIdx) = CompCode & WW_InFile_FieldValue(OilTypeIdx).Trim.ToString & WW_InFile_FieldValue(Product1Idx).Trim.ToString & WW_InFile_FieldValue(Product2Idx).Trim.ToString      '品名コード
                                End If

                                Dim ShabanFIdx As Integer = 0
                                ShabanFIdx = WW_InFile_Field.IndexOf("TSHABANF")
                                If WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanFIdx) = CompCode0 & WW_InFile_FieldValue(ShabanFIdx).Trim.ToString
                                End If

                                Dim ShabanBIdx As Integer = 0
                                ShabanBIdx = WW_InFile_Field.IndexOf("TSHABANB")
                                If WW_InFile_FieldValue(ShabanBIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanBIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanBIdx) = CompCode0 & WW_InFile_FieldValue(ShabanBIdx).Trim.ToString
                                End If

                                Dim ShabanB2Idx As Integer = 0
                                ShabanB2Idx = WW_InFile_Field.IndexOf("TSHABANB2")
                                If WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanB2Idx) = CompCode0 & WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString
                                End If
                            End If

                        Case "T0005_NIPPO"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("PRODUCTCODE1")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(45, "PRODUCTCODE1")                             '品名コード１
                                WW_InFile_Field.Insert(51, "PRODUCTCODE2")                             '品名コード２
                                WW_InFile_Field.Insert(57, "PRODUCTCODE3")                             '品名コード３
                                WW_InFile_Field.Insert(63, "PRODUCTCODE4")                             '品名コード４
                                WW_InFile_Field.Insert(69, "PRODUCTCODE5")                             '品名コード５
                                WW_InFile_Field.Insert(75, "PRODUCTCODE6")                             '品名コード６
                                WW_InFile_Field.Insert(81, "PRODUCTCODE7")                             '品名コード７
                                WW_InFile_Field.Insert(87, "PRODUCTCODE8")                             '品名コード８
                                WW_InFile_Field.Insert(139, "L1HAISOGROUP")                            '配送グループ
                            Else
                                WW_InFile_FieldValue.Insert(45, "")
                                WW_InFile_FieldValue.Insert(51, "")
                                WW_InFile_FieldValue.Insert(57, "")
                                WW_InFile_FieldValue.Insert(63, "")
                                WW_InFile_FieldValue.Insert(69, "")
                                WW_InFile_FieldValue.Insert(75, "")
                                WW_InFile_FieldValue.Insert(81, "")
                                WW_InFile_FieldValue.Insert(87, "")
                                WW_InFile_FieldValue.Insert(139, "")
                                '明細
                                Dim OilType1Idx As Integer = 0
                                Dim Product11Idx As Integer = 0
                                Dim Product21Idx As Integer = 0
                                Dim ProductCode1Idx As Integer = 0
                                OilType1Idx = WW_InFile_Field.IndexOf("OILTYPE1")
                                Product11Idx = WW_InFile_Field.IndexOf("PRODUCT11")
                                Product21Idx = WW_InFile_Field.IndexOf("PRODUCT21")
                                ProductCode1Idx = WW_InFile_Field.IndexOf("PRODUCTCODE1")
                                If WW_InFile_FieldValue(OilType1Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product11Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product21Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilType1Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product11Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product21Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCode1Idx) = CompCode & WW_InFile_FieldValue(OilType1Idx).Trim.ToString & WW_InFile_FieldValue(Product11Idx).Trim.ToString & WW_InFile_FieldValue(Product21Idx).Trim.ToString      '品名コード１
                                End If

                                Dim OilType2Idx As Integer = 0
                                Dim Product12Idx As Integer = 0
                                Dim Product22Idx As Integer = 0
                                Dim ProductCode2Idx As Integer = 0
                                OilType2Idx = WW_InFile_Field.IndexOf("OILTYPE2")
                                Product12Idx = WW_InFile_Field.IndexOf("PRODUCT12")
                                Product22Idx = WW_InFile_Field.IndexOf("PRODUCT22")
                                ProductCode2Idx = WW_InFile_Field.IndexOf("PRODUCTCODE2")
                                If WW_InFile_FieldValue(OilType2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product12Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product22Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilType2Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product12Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product22Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCode2Idx) = CompCode & WW_InFile_FieldValue(OilType2Idx).Trim.ToString & WW_InFile_FieldValue(Product12Idx).Trim.ToString & WW_InFile_FieldValue(Product22Idx).Trim.ToString      '品名コード２
                                End If

                                Dim OilType3Idx As Integer = 0
                                Dim Product13Idx As Integer = 0
                                Dim Product23Idx As Integer = 0
                                Dim ProductCode3Idx As Integer = 0
                                OilType3Idx = WW_InFile_Field.IndexOf("OILTYPE3")
                                Product13Idx = WW_InFile_Field.IndexOf("PRODUCT13")
                                Product23Idx = WW_InFile_Field.IndexOf("PRODUCT23")
                                ProductCode3Idx = WW_InFile_Field.IndexOf("PRODUCTCODE3")
                                If WW_InFile_FieldValue(OilType3Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product13Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product23Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilType3Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product13Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product23Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCode3Idx) = CompCode & WW_InFile_FieldValue(OilType3Idx).Trim.ToString & WW_InFile_FieldValue(Product13Idx).Trim.ToString & WW_InFile_FieldValue(Product23Idx).Trim.ToString      '品名コード３
                                End If

                                Dim OilType4Idx As Integer = 0
                                Dim Product14Idx As Integer = 0
                                Dim Product24Idx As Integer = 0
                                Dim ProductCode4Idx As Integer = 0
                                OilType4Idx = WW_InFile_Field.IndexOf("OILTYPE4")
                                Product14Idx = WW_InFile_Field.IndexOf("PRODUCT14")
                                Product24Idx = WW_InFile_Field.IndexOf("PRODUCT24")
                                ProductCode4Idx = WW_InFile_Field.IndexOf("PRODUCTCODE4")
                                If WW_InFile_FieldValue(OilType4Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product14Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product24Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilType4Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product14Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product24Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCode4Idx) = CompCode & WW_InFile_FieldValue(OilType4Idx).Trim.ToString & WW_InFile_FieldValue(Product14Idx).Trim.ToString & WW_InFile_FieldValue(Product24Idx).Trim.ToString      '品名コード４
                                End If

                                Dim OilType5Idx As Integer = 0
                                Dim Product15Idx As Integer = 0
                                Dim Product25Idx As Integer = 0
                                Dim ProductCode5Idx As Integer = 0
                                OilType5Idx = WW_InFile_Field.IndexOf("OILTYPE5")
                                Product15Idx = WW_InFile_Field.IndexOf("PRODUCT15")
                                Product25Idx = WW_InFile_Field.IndexOf("PRODUCT25")
                                ProductCode5Idx = WW_InFile_Field.IndexOf("PRODUCTCODE5")
                                If WW_InFile_FieldValue(OilType5Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product15Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product25Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilType5Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product15Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product25Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCode5Idx) = CompCode & WW_InFile_FieldValue(OilType5Idx).Trim.ToString & WW_InFile_FieldValue(Product15Idx).Trim.ToString & WW_InFile_FieldValue(Product25Idx).Trim.ToString      '品名コード５
                                End If

                                Dim OilType6Idx As Integer = 0
                                Dim Product16Idx As Integer = 0
                                Dim Product26Idx As Integer = 0
                                Dim ProductCode6Idx As Integer = 0
                                OilType6Idx = WW_InFile_Field.IndexOf("OILTYPE6")
                                Product16Idx = WW_InFile_Field.IndexOf("PRODUCT16")
                                Product26Idx = WW_InFile_Field.IndexOf("PRODUCT26")
                                ProductCode6Idx = WW_InFile_Field.IndexOf("PRODUCTCODE6")
                                If WW_InFile_FieldValue(OilType6Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product16Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product26Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilType6Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product16Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product26Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCode6Idx) = CompCode & WW_InFile_FieldValue(OilType6Idx).Trim.ToString & WW_InFile_FieldValue(Product16Idx).Trim.ToString & WW_InFile_FieldValue(Product26Idx).Trim.ToString      '品名コード６
                                End If

                                Dim OilType7Idx As Integer = 0
                                Dim Product17Idx As Integer = 0
                                Dim Product27Idx As Integer = 0
                                Dim ProductCode7Idx As Integer = 0
                                OilType7Idx = WW_InFile_Field.IndexOf("OILTYPE7")
                                Product17Idx = WW_InFile_Field.IndexOf("PRODUCT17")
                                Product27Idx = WW_InFile_Field.IndexOf("PRODUCT27")
                                ProductCode7Idx = WW_InFile_Field.IndexOf("PRODUCTCODE7")
                                If WW_InFile_FieldValue(OilType7Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product17Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product27Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilType7Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product17Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product27Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCode7Idx) = CompCode & WW_InFile_FieldValue(OilType7Idx).Trim.ToString & WW_InFile_FieldValue(Product17Idx).Trim.ToString & WW_InFile_FieldValue(Product27Idx).Trim.ToString      '品名コード７
                                End If

                                Dim OilType8Idx As Integer = 0
                                Dim Product18Idx As Integer = 0
                                Dim Product28Idx As Integer = 0
                                Dim ProductCode8Idx As Integer = 0
                                OilType8Idx = WW_InFile_Field.IndexOf("OILTYPE8")
                                Product18Idx = WW_InFile_Field.IndexOf("PRODUCT18")
                                Product28Idx = WW_InFile_Field.IndexOf("PRODUCT28")
                                ProductCode8Idx = WW_InFile_Field.IndexOf("PRODUCTCODE8")
                                If WW_InFile_FieldValue(OilType8Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product18Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(Product28Idx).Trim.ToString <> vbNullChar AndAlso
                                    WW_InFile_FieldValue(OilType8Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product18Idx).Trim.ToString <> "" AndAlso WW_InFile_FieldValue(Product28Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ProductCode8Idx) = CompCode & WW_InFile_FieldValue(OilType8Idx).Trim.ToString & WW_InFile_FieldValue(Product18Idx).Trim.ToString & WW_InFile_FieldValue(Product28Idx).Trim.ToString      '品名コード８
                                End If

                                Dim ShabanFIdx As Integer = 0
                                ShabanFIdx = WW_InFile_Field.IndexOf("TSHABANF")
                                If WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanFIdx) = CompCode0 & WW_InFile_FieldValue(ShabanFIdx).Trim.ToString
                                End If

                                Dim ShabanBIdx As Integer = 0
                                ShabanBIdx = WW_InFile_Field.IndexOf("TSHABANB")
                                If WW_InFile_FieldValue(ShabanBIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanBIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanBIdx) = CompCode0 & WW_InFile_FieldValue(ShabanBIdx).Trim.ToString
                                End If

                                Dim ShabanB2Idx As Integer = 0
                                ShabanB2Idx = WW_InFile_Field.IndexOf("TSHABANB2")
                                If WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanB2Idx) = CompCode0 & WW_InFile_FieldValue(ShabanB2Idx).Trim.ToString
                                End If
                            End If

                        Case "T0007_KINTAI"

                            If WW_Linecnt = 0 Then
                                'ヘッダ
                                '既に変換済みの場合、スルーする
                                WW_CONV = WW_InFile_Field.IndexOf("HAYADETIME")
                                If WW_CONV >= 0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If
                                WW_InFile_Field.Insert(89, "HAYADETIME")                                '早出補填時間
                                WW_InFile_Field.Insert(90, "HAYADETIMECHO")                             '早出補填時間調整
                                WW_InFile_Field.Insert(115, "HAISOTIME")                                '配送時間
                                WW_InFile_Field.Insert(116, "NENMATUNISSU")                             '年末出勤日数
                                WW_InFile_Field.Insert(117, "NENMATUNISSUCHO")                          '年末出勤日数調整
                                WW_InFile_Field.Insert(118, "SHACHUHAKKBN")                             '車中泊区分
                                WW_InFile_Field.Insert(119, "SHACHUHAKNISSU")                           '車中泊日数
                                WW_InFile_Field.Insert(120, "SHACHUHAKNISSUCHO")                        '車中泊日数調整
                                WW_InFile_Field.Insert(121, "MODELDISTANCE")                            'モデル距離
                                WW_InFile_Field.Insert(122, "MODELDISTANCECHO")                         'モデル距離調整
                                WW_InFile_Field.Insert(123, "JIKYUSHATIME")                             '時給者時間
                                WW_InFile_Field.Insert(124, "JIKYUSHATIMECHO")                          '時給者時間調整
                                WW_InFile_Field.Insert(125, "HDAIWORKTIME")                             '代休出勤
                                WW_InFile_Field.Insert(126, "HDAIWORKTIMECHO")                          '代休出勤調整
                                WW_InFile_Field.Insert(127, "HDAINIGHTTIME")                            '代休深夜
                                WW_InFile_Field.Insert(128, "HDAINIGHTTIMECHO")                         '代休深夜調整
                                WW_InFile_Field.Insert(129, "SDAIWORKTIME")                             '日曜代休出勤
                                WW_InFile_Field.Insert(130, "SDAIWORKTIMECHO")                          '日曜代休出勤調整
                                WW_InFile_Field.Insert(131, "SDAINIGHTTIME")                            '日曜代休出勤
                                WW_InFile_Field.Insert(132, "SDAINIGHTTIMECHO")                         '日曜代休出勤調整
                                WW_InFile_Field.Insert(133, "WWORKTIME")                                '所定内時間
                                WW_InFile_Field.Insert(134, "WWORKTIMECHO")                             '所定内時間調整
                                WW_InFile_Field.Insert(135, "JYOMUTIME")                                '乗務時間
                                WW_InFile_Field.Insert(136, "JYOMUTIMECHO")                             '乗務時間調整
                                WW_InFile_Field.Insert(137, "HWORKNISSU")                               '休日出勤日数
                                WW_InFile_Field.Insert(138, "HWORKNISSUCHO")                            '休日出勤日数調整
                                WW_InFile_Field.Insert(139, "KAITENCNT")                                '回転数
                                WW_InFile_Field.Insert(140, "KAITENCNTCHO")                             '回転数調整
                                WW_InFile_Field.Insert(141, "KAITENCNT1_1")                             '回転数1-1
                                WW_InFile_Field.Insert(142, "KAITENCNTCHO1_1")                          '回転数調整1-1
                                WW_InFile_Field.Insert(143, "KAITENCNT1_2")                             '回転数1-2
                                WW_InFile_Field.Insert(144, "KAITENCNTCHO1_2")                          '回転数調整1-2
                                WW_InFile_Field.Insert(145, "KAITENCNT1_3")                             '回転数1-3
                                WW_InFile_Field.Insert(146, "KAITENCNTCHO1_3")                          '回転数調整1-3
                                WW_InFile_Field.Insert(147, "KAITENCNT1_4")                             '回転数1-4
                                WW_InFile_Field.Insert(148, "KAITENCNTCHO1_4")                          '回転数調整1-4
                                WW_InFile_Field.Insert(149, "KAITENCNT2_1")                             '回転数2-1
                                WW_InFile_Field.Insert(150, "KAITENCNTCHO2_1")                          '回転数調整2-1
                                WW_InFile_Field.Insert(151, "KAITENCNT2_2")                             '回転数2-2
                                WW_InFile_Field.Insert(152, "KAITENCNTCHO2_2")                          '回転数調整2-2
                                WW_InFile_Field.Insert(153, "KAITENCNT2_3")                             '回転数2-3
                                WW_InFile_Field.Insert(154, "KAITENCNTCHO2_3")                          '回転数調整2-3
                                WW_InFile_Field.Insert(155, "KAITENCNT2_4")                             '回転数2-4
                                WW_InFile_Field.Insert(156, "KAITENCNTCHO2_4")                          '回転数調整2-4
                                WW_InFile_Field.Insert(157, "SENJYOCNT")                                '洗浄回数
                                WW_InFile_Field.Insert(158, "SENJYOCNTCHO")                             '洗浄回数調整
                                WW_InFile_Field.Insert(159, "UNLOADADDCNT1")                            '危険物荷卸回数1
                                WW_InFile_Field.Insert(160, "UNLOADADDCNT1CHO")                         '危険物荷卸回数1調整
                                WW_InFile_Field.Insert(161, "UNLOADADDCNT2")                            '危険物荷卸回数2
                                WW_InFile_Field.Insert(162, "UNLOADADDCNT2CHO")                         '危険物荷卸回数2調整
                                WW_InFile_Field.Insert(163, "UNLOADADDCNT3")                            '危険物荷卸回数3
                                WW_InFile_Field.Insert(164, "UNLOADADDCNT3CHO")                         '危険物荷卸回数3調整
                                WW_InFile_Field.Insert(165, "UNLOADADDCNT4")                            '危険物荷卸回数4
                                WW_InFile_Field.Insert(166, "UNLOADADDCNT4CHO")                         '危険物荷卸回数4調整
                                WW_InFile_Field.Insert(167, "LOADINGCNT1")                              '危険品積込回数1
                                WW_InFile_Field.Insert(168, "LOADINGCNT1CHO")                           '危険品積込回数1調整
                                WW_InFile_Field.Insert(169, "LOADINGCNT2")                              '危険品積込回数2
                                WW_InFile_Field.Insert(170, "LOADINGCNT2CHO")                           '危険品積込回数2調整
                                WW_InFile_Field.Insert(171, "SHORTDISTANCE1")                           '短距離手当1
                                WW_InFile_Field.Insert(172, "SHORTDISTANCE1CHO")                        '短距離手当1調整
                                WW_InFile_Field.Insert(173, "SHORTDISTANCE2")                           '短距離手当2
                                WW_InFile_Field.Insert(174, "SHORTDISTANCE2CHO")                        '短距離手当2調整
                            Else
                                '明細
                                WW_InFile_FieldValue.Insert(89, "0")                                   '早出補填
                                WW_InFile_FieldValue.Insert(90, "0")                                   '早出補填調整
                                WW_InFile_FieldValue.Insert(115, "0")                                   '配送時間
                                WW_InFile_FieldValue.Insert(116, "0")                                   '年末出勤日数
                                WW_InFile_FieldValue.Insert(117, "0")                                   '年末出勤日数調整
                                WW_InFile_FieldValue.Insert(118, "0")                                   '車中泊区分
                                WW_InFile_FieldValue.Insert(119, "0")                                   '車中泊日数
                                WW_InFile_FieldValue.Insert(120, "0")                                   '車中泊日数調整
                                WW_InFile_FieldValue.Insert(121, "0")                                   'モデル距離
                                WW_InFile_FieldValue.Insert(122, "0")                                   'モデル距離調整
                                WW_InFile_FieldValue.Insert(123, "0")                                   '時給者時間
                                WW_InFile_FieldValue.Insert(124, "0")                                   '時給者時間調整
                                WW_InFile_FieldValue.Insert(125, "0")                                   '代休出勤
                                WW_InFile_FieldValue.Insert(126, "0")                                   '代休出勤調整
                                WW_InFile_FieldValue.Insert(127, "0")                                   '代休深夜
                                WW_InFile_FieldValue.Insert(128, "0")                                   '代休深夜調整
                                WW_InFile_FieldValue.Insert(129, "0")                                   '日曜代休出勤
                                WW_InFile_FieldValue.Insert(130, "0")                                   '日曜代休出勤調整
                                WW_InFile_FieldValue.Insert(131, "0")                                   '日曜代休出勤
                                WW_InFile_FieldValue.Insert(132, "0")                                   '日曜代休出勤調整
                                WW_InFile_FieldValue.Insert(133, "0")                                   '所定内時間
                                WW_InFile_FieldValue.Insert(134, "0")                                   '所定内時間調整
                                WW_InFile_FieldValue.Insert(135, "0")                                   '乗務時間
                                WW_InFile_FieldValue.Insert(136, "0")                                   '乗務時間調整
                                WW_InFile_FieldValue.Insert(137, "0")                                   '休日出勤日数
                                WW_InFile_FieldValue.Insert(138, "0")                                   '休日出勤日数調整
                                WW_InFile_FieldValue.Insert(139, "0")                                   '回転数
                                WW_InFile_FieldValue.Insert(140, "0")                                   '回転数調整
                                WW_InFile_FieldValue.Insert(141, "0")                                   '回転数1-1
                                WW_InFile_FieldValue.Insert(142, "0")                                   '回転数調整1-1
                                WW_InFile_FieldValue.Insert(143, "0")                                   '回転数1-2
                                WW_InFile_FieldValue.Insert(144, "0")                                   '回転数調整1-2
                                WW_InFile_FieldValue.Insert(145, "0")                                   '回転数1-3
                                WW_InFile_FieldValue.Insert(146, "0")                                   '回転数調整1-3
                                WW_InFile_FieldValue.Insert(147, "0")                                   '回転数1-4
                                WW_InFile_FieldValue.Insert(148, "0")                                   '回転数調整1-4
                                WW_InFile_FieldValue.Insert(149, "0")                                   '回転数2-1
                                WW_InFile_FieldValue.Insert(150, "0")                                   '回転数調整2-1
                                WW_InFile_FieldValue.Insert(151, "0")                                   '回転数2-2
                                WW_InFile_FieldValue.Insert(152, "0")                                   '回転数調整2-2
                                WW_InFile_FieldValue.Insert(153, "0")                                   '回転数2-3
                                WW_InFile_FieldValue.Insert(154, "0")                                   '回転数調整2-3
                                WW_InFile_FieldValue.Insert(155, "0")                                   '回転数2-4
                                WW_InFile_FieldValue.Insert(156, "0")                                   '回転数調整2-4
                                WW_InFile_FieldValue.Insert(157, "0")                                   '洗浄回数
                                WW_InFile_FieldValue.Insert(158, "0")                                   '洗浄回数調整
                                WW_InFile_FieldValue.Insert(159, "0")                                   '危険物荷卸回数1
                                WW_InFile_FieldValue.Insert(160, "0")                                   '危険物荷卸回数1調整
                                WW_InFile_FieldValue.Insert(161, "0")                                   '危険物荷卸回数2
                                WW_InFile_FieldValue.Insert(162, "0")                                   '危険物荷卸回数2調整
                                WW_InFile_FieldValue.Insert(163, "0")                                   '危険物荷卸回数3
                                WW_InFile_FieldValue.Insert(164, "0")                                   '危険物荷卸回数3調整
                                WW_InFile_FieldValue.Insert(165, "0")                                   '危険物荷卸回数4
                                WW_InFile_FieldValue.Insert(166, "0")                                   '危険物荷卸回数4調整
                                WW_InFile_FieldValue.Insert(167, "0")                                   '危険品積込回数1
                                WW_InFile_FieldValue.Insert(168, "0")                                   '危険品積込回数1調整
                                WW_InFile_FieldValue.Insert(169, "0")                                   '危険品積込回数2
                                WW_InFile_FieldValue.Insert(170, "0")                                   '危険品積込回数2調整
                                WW_InFile_FieldValue.Insert(171, "0")                                   '短距離手当1
                                WW_InFile_FieldValue.Insert(172, "0")                                   '短距離手当1調整
                                WW_InFile_FieldValue.Insert(173, "0")                                   '短距離手当2
                                WW_InFile_FieldValue.Insert(174, "0")                                   '短距離手当2調整
                            End If

                        Case "TA001_SHARYOSTAT"
                            If WW_Linecnt = 0 Then
                                'ヘッダ
                            Else
                                '明細
                                Dim ShabanFIdx As Integer = 0
                                ShabanFIdx = WW_InFile_Field.IndexOf("TSHABAN")

                                '既に変換済みの場合、スルーする
                                If Trim(WW_InFile_FieldValue(ShabanFIdx)).Length > 5 And Mid(WW_InFile_FieldValue(ShabanFIdx), 1, 3) = CompCode0 Then
                                    '変換なし
                                    TgtFlg = False
                                    Exit While
                                End If

                                If WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> vbNullChar AndAlso WW_InFile_FieldValue(ShabanFIdx).Trim.ToString <> "" Then
                                    WW_InFile_FieldValue(ShabanFIdx) = CompCode0 & WW_InFile_FieldValue(ShabanFIdx).Trim.ToString
                                End If
                            End If

                        Case Else
                            '変換なし
                            TgtFlg = False

                    End Select

                    If Not WW_Linecnt = 0 Then

                        'リスト設定
                        Select Case WW_InPARA_TBLNAME
                            'Case "S0007_UPROFVARI", "S0010_UPROFVIEW", "S0011_UPROFXLS"
                            Case "S0023_PROFMVARI", "S0025_PROFMVIEW", "S0026_PROFMXLS"
                                If AppFlg Then
                                    AryList.Add(WW_InFile_FieldValue.ToArray)
                                End If

                            Case Else
                                AryList.Add(WW_InFile_FieldValue.ToArray)
                        End Select

                    End If

                    WW_Linecnt = WW_Linecnt + 1

                End While

                sr.Close()
                sr.Dispose()
                sr = Nothing

                '変換なしの場合、作成ファイルを削除して次のファイルへ
                If Not TgtFlg Then

                    'ファイル削除
                    Try
                        '閉じる
                        sw.Close()
                        sw.Dispose()
                        System.IO.File.Delete(fileName)
                    Catch ex As Exception
                        CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "ファイル削除失敗" & WW_file       '
                        CS0054LOGWrite_bat.NIWEA = "A"                                  '
                        CS0054LOGWrite_bat.TEXT = ex.ToString
                        CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                        Environment.Exit(200)
                    End Try

                    '次ファイルへ
                    Continue For
                End If

                'ファイル削除
                Try
                    System.IO.File.Delete(WW_file)
                Catch ex As Exception
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "ファイル削除失敗" & WW_file       '
                    CS0054LOGWrite_bat.NIWEA = "A"                                  '
                    CS0054LOGWrite_bat.TEXT = ex.ToString
                    CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                    Environment.Exit(200)
                End Try

                'ファイル作成
                If AryList.Count > 0 Then
                    'TABLEフォルダーに抽出データファイルを出力（テーブル名.dat)
                    Dim WriteStr As String = ""

                    Try
                        If Not headFlg Then
                            'DATヘッダーデータ出力
                            For i As Integer = 0 To WW_InFile_Field.Count - 1
                                WriteStr = WriteStr & WW_InFile_Field.Item(i).ToString
                                If (WW_InFile_Field.Count - 1) = i Then
                                    WriteStr = WriteStr & ControlChars.NewLine
                                Else
                                    WriteStr = WriteStr & ControlChars.Tab
                                End If
                            Next
                            'DAT Line出力
                            sw.Write(WriteStr)
                        End If

                        'DAT明細データ出力
                        For j As Integer = 0 To AryList.Count - 1
                            WriteStr = ""
                            For k As Integer = 0 To AryList(j).Count - 1
                                WriteStr = WriteStr & AryList(j)(k).ToString
                                If (AryList(j).Count - 1) = k Then
                                    WriteStr = WriteStr & ControlChars.NewLine
                                Else
                                    WriteStr = WriteStr & ControlChars.Tab
                                End If
                            Next
                            'DAT Line出力
                            sw.Write(WriteStr)
                        Next

                    Catch ex As System.SystemException
                        '閉じる
                        sw.Close()
                        sw.Dispose()

                        CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = WW_InPARA_TBLNAME & " FILE OUTPUT ERR"    '
                        CS0054LOGWrite_bat.NIWEA = "A"                                  '
                        CS0054LOGWrite_bat.TEXT = ex.ToString
                        CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                        Environment.Exit(100)

                    End Try

                End If

                '閉じる
                sw.Close()
                sw.Dispose()

                If fileName.IndexOf("Changing") <> -1 Then

                    Dim repFileName As String = fileName.Replace("Changing", "")

                    'ファイル名変更
                    Try
                        System.IO.File.Move(fileName, repFileName)
                    Catch ex As Exception
                        CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "ファイル移動失敗" & WW_file       '
                        CS0054LOGWrite_bat.NIWEA = "A"                                  '
                        CS0054LOGWrite_bat.TEXT = ex.ToString
                        CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                        Environment.Exit(200)
                    End Try

                End If

            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = WW_InPARA_TBLNAME & " UPDATE/INSERT"               '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

                Environment.Exit(100)
            End Try
        Next

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate3"             'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00008TBLupdate3処理終了"        '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00008TBLupdate3処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
        Environment.Exit(0)

    End Sub

End Module

