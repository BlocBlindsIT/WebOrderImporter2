Imports System.Text

Public Class ImportTimer

    Private Const TIMER_INTERVAL As Integer = 600000        ' Milliseconds
    Private WithEvents xTimer As New System.Windows.Forms.Timer

    Public Sub New()
        xTimer = New Timer
        xTimer.Interval = TIMER_INTERVAL
    End Sub

    Public Sub StartTimer()
        Timer_Tick()
        xTimer.Start()
    End Sub

    Public Sub StopTimer()
        xTimer.Stop()
    End Sub

    Private Sub Timer_Tick() Handles xTimer.Tick

        ' --- Writing out start of import attempt ---
        Main.Console_WriteMessage("")
        Main.Console_WriteMessage("Attempting import", True)


        ' --- Attempting to set Order Import lock ---
        Dim resLocked As ResultBoolean = ImportLock_Set()
        If resLocked.Result = True Then


            ' --- Attempting import ---
            For Each strWebsite In System.Enum.GetNames(GetType(Importer.enumWebsites))
                Dim lstOutput As List(Of String) = Importer.ImportOrders(strWebsite)
                For Each strOutput As String In lstOutput
                    Main.Console_WriteMessage(strOutput, strWebsite)
                Next

                Dim strFabricImportResult As String = ImportFabricSamples(strWebsite)
                Main.Console_WriteMessage(strFabricImportResult, strWebsite)
            Next


            ' --- Attempting to release lock on database ---
            Dim resLockReleased As ResultBoolean = ImportLock_Release()
            If resLockReleased.Result = False Then
                Main.Console_WriteMessage(resLockReleased.Message())
            End If

        Else
            Main.Console_WriteMessage(resLocked.Message)

        End If



        Main.Console_WriteMessage("Next import at " & DateAdd(DateInterval.Second, TIMER_INTERVAL / 1000, Now).ToString("yyyy-MM-dd HH:mm:ss"))
        Main.RefreshMainForm()

    End Sub

End Class
