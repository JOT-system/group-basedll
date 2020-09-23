Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CB0011JobControl
    Inherits System.ServiceProcess.ServiceBase

    'UserService は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' 処理のメイン エントリ ポイントです。
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' 2 つ以上の NT サービスを同じプロセス内で実行できます。別のサービスを
        ' この処理に追加するには、以下の行を追加して
        ' 2 番目のサービス オブジェクトを作成してください。例 :
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New CB0011JobControl}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'コンポーネント デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ: 以下のプロシージャはコンポーネント デザイナーで必要です。
    ' コンポーネント デザイナーを使って変更できます。  
    ' コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.JobTimer = New System.Timers.Timer()
        CType(Me.JobTimer, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'JobTimer
        '
        Me.JobTimer.Interval = 1000.0R
        '
        'CB0011JobControl
        '
        Me.ServiceName = "CB0011JobControl"
        CType(Me.JobTimer, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Friend WithEvents JobTimer As System.Timers.Timer

End Class
