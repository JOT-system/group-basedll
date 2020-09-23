Imports System
Imports System.IO

Module CB00003FolderCopy
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
	'■　コマンド例.  CB00003FolderCopy /@1 /@2 /@P3 /@4　　　　　　　　　　　　　　　　　　 ■
	'■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
	'■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
	'■　　・@1：Copy元フォルダー　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
	'■　　・@2：Copy先フォルダー 　　　　　　　　                                           ■
	'■　　・@3：世代数 　　　　　　　　　　　　　                                           ■
	'■　　・@4：運用サイクル　(Y:年、M:月、D:日、以外:日)　　　　　　　　　                 ■
	'■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
	'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

	Sub Main()

		'■■■　共通宣言　■■■
		'*共通関数宣言(BATDLL)
		Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat                 'ログ格納ディレクトリ取得
		Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat             'LogOutput DirString Get

		'■■■　共通処理　■■■
		'○ ログ格納ディレクトリ取得(InParm無し)
		Dim WW_LOGdir As String = ""
		CS0052LOGdir_bat.CS0052LOGdir_bat()
		If CS0052LOGdir_bat.ERR = "00000" Then
			WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                    'ログ格納ディレクトリ格納
		Else
			Exit Sub
		End If

		'■■■　開始メッセージ　■■■
		CS0054LOGWrite_bat.INFNMSPACE = "CB00003FolderCopy"                 'NameSpace
		CS0054LOGWrite_bat.INFCLASS = "Main"                                'クラス名
		CS0054LOGWrite_bat.INFSUBCLASS = "Main"                             'SUBクラス名
		CS0054LOGWrite_bat.INFPOSI = "CB00003FolderCopy処理開始"
		CS0054LOGWrite_bat.NIWEA = "I"
		CS0054LOGWrite_bat.TEXT = "CB00003FolderCopy処理開始"
		CS0054LOGWrite_bat.MESSAGENO = "00000"
		CS0054LOGWrite_bat.CS0054LOGWrite_bat()

		'■■■　コマンドライン引数の取得　■■■
		Dim InPARA_FolderFrom As String = ""
		Dim InPARA_FolderTo As String = ""
		Dim InPARA_Gener As Integer = 7
		Dim InPARA_Cycle As String = ""
		Dim WW_CopyFolderNM As String = ""

		'コマンドライン引数を配列取得
		Dim WW_cmds_cnt As Integer = 0
		For Each cmd As String In System.Environment.GetCommandLineArgs()
			Select Case WW_cmds_cnt
				Case 1     'Copy元フォルダー
					InPARA_FolderFrom = Mid(cmd, 2, 100)
					Console.WriteLine("引数(Copy元　　　)：" & InPARA_FolderFrom)
				Case 2     'Copy先フォルダー 
					InPARA_FolderTo = Mid(cmd, 2, 100)
					Console.WriteLine("引数(Copy先　　　)：" & InPARA_FolderTo)
				Case 3     '世代数
					Try
						Integer.TryParse(Mid(cmd, 2, 100), InPARA_Gener)
						If InPARA_Gener = 0 Then
							CS0054LOGWrite_bat.INFNMSPACE = "CB00003FolderCopy"             'NameSpace
							CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
							CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
							CS0054LOGWrite_bat.INFPOSI = "引数3エラー：" & cmd
							CS0054LOGWrite_bat.NIWEA = "E"
							CS0054LOGWrite_bat.TEXT = cmd
							CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
							CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
						End If
					Catch ex As Exception
						CS0054LOGWrite_bat.INFNMSPACE = "CB00003FolderCopy"                 'NameSpace
						CS0054LOGWrite_bat.INFCLASS = "Main"                                'クラス名
						CS0054LOGWrite_bat.INFSUBCLASS = "Main"                             'SUBクラス名
						CS0054LOGWrite_bat.INFPOSI = "引数3エラー：" & cmd
						CS0054LOGWrite_bat.NIWEA = "E"
						CS0054LOGWrite_bat.TEXT = ex.ToString
						CS0054LOGWrite_bat.MESSAGENO = "00000"                              'DBエラー
						CS0054LOGWrite_bat.CS0054LOGWrite_bat()                             'ログ入力
					End Try
					Console.WriteLine("引数(世代　　　　)：" & InPARA_Gener)
				Case 4     'Copyサイクル(D:日単位、M:月単位、Y:年単位) 
					InPARA_Cycle = Mid(cmd, 1, 100)
					Console.WriteLine("引数(サイクル　　)：" & InPARA_Cycle)
			End Select

			WW_cmds_cnt = WW_cmds_cnt + 1
		Next

		'■■■　コマンドライン　チェック　■■■
		'○ パラメータチェック(Copy元)

		'　自SRVディレクトリのみ可(\\xxxx形式は×)
		If InStr(InPARA_FolderFrom, ":") = 0 Or Mid(InPARA_FolderFrom, 2, 1) <> ":" Then
			CS0054LOGWrite_bat.INFNMSPACE = "CB00003FolderCopy"                             'NameSpace
			CS0054LOGWrite_bat.INFCLASS = "Main"                                            'クラス名
			CS0054LOGWrite_bat.INFSUBCLASS = "Main"                                         'SUBクラス名
			CS0054LOGWrite_bat.INFPOSI = "引数1チェック"
			CS0054LOGWrite_bat.NIWEA = "E"
			CS0054LOGWrite_bat.TEXT = "引数1フォーマットエラー：" & InPARA_FolderFrom
			CS0054LOGWrite_bat.MESSAGENO = "00002"                                          'パラメータエラー
			CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                         'ログ出力
			Exit Sub
		End If

		'　実在チェック
		If System.IO.Directory.Exists(InPARA_FolderFrom) Then
		Else
			CS0054LOGWrite_bat.INFNMSPACE = "CB00003FolderCopy"             'NameSpace
			CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
			CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
			CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
			CS0054LOGWrite_bat.NIWEA = "E"                                  '
			CS0054LOGWrite_bat.TEXT = "引数1指定ディレクトリ無し：" & InPARA_FolderFrom
			CS0054LOGWrite_bat.MESSAGENO = "00008"                          'ディレクトリ存在しない
			CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
			Exit Sub
		End If

		'○ パラメータチェック(Copy先)

		'　自SRVディレクトリのみ可(\\xxxx形式は×)
		If InStr(InPARA_FolderTo, ":") = 0 Or Mid(InPARA_FolderTo, 2, 1) <> ":" Then
			CS0054LOGWrite_bat.INFNMSPACE = "CB00003FolderCopy"             'NameSpace
			CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
			CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
			CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
			CS0054LOGWrite_bat.NIWEA = "E"                                  '
			CS0054LOGWrite_bat.TEXT = "引数1フォーマットエラー：" & InPARA_FolderTo
			CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
			CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
			Exit Sub
		End If

		'　実在チェック
		If System.IO.Directory.Exists(InPARA_FolderTo) Then
		Else
			CS0054LOGWrite_bat.INFNMSPACE = "CB00003FolderCopy"             'NameSpace
			CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
			CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
			CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
			CS0054LOGWrite_bat.NIWEA = "E"                                  '
			CS0054LOGWrite_bat.TEXT = "引数1指定ディレクトリ無し：" & InPARA_FolderTo
			CS0054LOGWrite_bat.MESSAGENO = "00008"                          'ディレクトリ存在しない
			CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
			Exit Sub
		End If

		'○ パラメータデフォルト設定(Copyサイクル)
		If InPARA_Cycle = Nothing Then
			InPARA_Cycle = "D"
		End If

		'■■■　フォルダ準備　■■■
		'○Copy先フォルダー作成
		Select Case InPARA_Cycle
			Case "Y"
				'　(Copy先フォルダー_年)
				WW_CopyFolderNM = InPARA_FolderTo & "\" & Date.Now.ToString("yyyy")
				If System.IO.Directory.Exists(WW_CopyFolderNM) Then
					System.IO.Directory.Delete(WW_CopyFolderNM, True)
				End If
				System.IO.Directory.CreateDirectory(WW_CopyFolderNM)
			Case "M"
				'　(Copy先フォルダー_月)
				WW_CopyFolderNM = InPARA_FolderTo & "\" & Date.Now.ToString("yyyyMM")
				If System.IO.Directory.Exists(WW_CopyFolderNM) Then
					System.IO.Directory.Delete(WW_CopyFolderNM, True)
				End If
				System.IO.Directory.CreateDirectory(WW_CopyFolderNM)
			Case "D"
				'　(Copy先フォルダー_年月日+時間)　…　日付内で複数フォルダー可能
				WW_CopyFolderNM = InPARA_FolderTo & "\" & Date.Now.ToString("yyyyMMdd") & "_" & Date.Now.ToString("HHmmss")
				System.IO.Directory.CreateDirectory(WW_CopyFolderNM)
		End Select

		'○Copy先フォルダーお掃除(世代数による)
		For Each FolderStr As String In System.IO.Directory.GetDirectories(InPARA_FolderTo, "*")

			'Copy先フォルダー直下のフォルダー名称取得
			Dim wFoldernm As String = FolderStr
			Do
				If InStr(wFoldernm, "\") <> 0 Then
					wFoldernm = Mid(wFoldernm, InStr(wFoldernm, "\") + 1, 100)
				End If
			Loop Until InStr(wFoldernm, "\") = 0

			Select Case InPARA_Cycle
				Case "Y"
					If IsNumeric(Mid(wFoldernm, 1, 4)) Then
						If CLng(Mid(wFoldernm, 1, 4)) < CLng(Date.Now.AddYears((InPARA_Gener - 1) * -1).ToString("yyyy")) Then
							'フォルダー削除
							System.IO.Directory.Delete(FolderStr, True)
						End If
					End If
				Case "M"
					If IsNumeric(Mid(wFoldernm, 1, 6)) Then
						If CLng(Mid(wFoldernm, 1, 6)) < CLng(Date.Now.AddMonths((InPARA_Gener - 1) * -1).ToString("yyyyMM")) Then
							'フォルダー削除
							System.IO.Directory.Delete(FolderStr, True)
						End If
					End If
				Case "D"
					If IsNumeric(Mid(wFoldernm, 1, 8)) Then
						If CLng(Mid(wFoldernm, 1, 8)) < CLng(Date.Now.AddDays((InPARA_Gener - 1) * -1).ToString("yyyyMMdd")) Then
							'フォルダー削除
							System.IO.Directory.Delete(FolderStr, True)
						End If
					End If
			End Select

		Next

		'■■■　コピー処理　■■■
		'○DOSコマンドプロセス準備
		Dim WW_CMDproc As New System.Diagnostics.Process()
		'　cmd.exeのパスを取得　⇒　FileNameプロパティに指定
		WW_CMDproc.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
		'出力を読み取れるようにする
		WW_CMDproc.StartInfo.UseShellExecute = False            'シェル使用
		WW_CMDproc.StartInfo.RedirectStandardOutput = True      '結果取得
		WW_CMDproc.StartInfo.RedirectStandardInput = False
		WW_CMDproc.StartInfo.CreateNoWindow = True              'ウィンドウ非表示

		'○コマンドライン指定（"/c"は実行後閉じるために必要）
		WW_CMDproc.StartInfo.Arguments = "/c xcopy " & InPARA_FolderFrom & " " & WW_CopyFolderNM & " /e /c /h "

		'○DOSコマンド起動
		'　実行
		WW_CMDproc.Start()
		Dim WW_results As String = WW_CMDproc.StandardOutput.ReadToEnd()

		'　プロセス終了まで待機する
		WW_CMDproc.WaitForExit()
		WW_CMDproc.Close()

		'■■■　終了メッセージ　■■■
		CS0054LOGWrite_bat.INFNMSPACE = "CB00003FolderCopy"             'NameSpace
		CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
		CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
		CS0054LOGWrite_bat.INFPOSI = "CB00003FolderCopy処理終了"
		CS0054LOGWrite_bat.NIWEA = "I"
		CS0054LOGWrite_bat.TEXT = "CB00003FolderCopy処理終了"
		CS0054LOGWrite_bat.MESSAGENO = "00000"
		CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

	End Sub


End Module
