Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Deployment.Application
Imports System.Text

Public Class Main

    Dim intCount As Integer = 0
    Dim class_ImportTimer As ImportTimer
    Dim class_lstConsoleMessages As New List(Of String)

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        RefreshMainForm()

    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            With System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
                lblVersion.Text = "Version" & .Major & "." & .Minor & "." & .Build & "." & .Revision
            End With
        Catch ex As Exception

        End Try

    End Sub


    Public Sub RefreshMainForm()

        ' --- Defining objects and variables ---
        Dim strSitesList As String = ConfigurationManager.AppSettings("Sites")
        Dim strSitesArray As String() = strSitesList.Split(",")
        Dim intSafetyCutout As Integer = 0
        Dim intCountOfOrdersToImport As Integer = 0
        Dim strSqlQuery As String
        Dim strConnection As String
        Dim strResult As String = ""
        Dim strFabricSamplesToImport As String = ""
        Dim cmd As SqlCommand


        ' --- Emptying listview ---
        intSafetyCutout = 0
        Do While lvwAvailableWebsites.Items.Count > 0 And intSafetyCutout < 50
            lvwAvailableWebsites.Items.RemoveAt(0)
            intSafetyCutout += 1
        Loop


        ' --- Setting query to get count of orders to import ---
        Dim sbQuery As New System.Text.StringBuilder
        sbQuery.Append("SELECT Count(order_no) FROM ORD_HEADER ")
        sbQuery.Append("WHERE OrderStatus = 'PAY_AUTH' And ProductionSystemID IS NULL ")
        sbQuery.Append("AND ((PaymentType <> 'FREE') ")
        sbQuery.Append("AND Exists (SELECT * FROM ORD_LINES WHERE ORD_LINES.Order_No = ORD_HEADER.Order_No AND ProductName <> 'MeasureService') ")                              ' Qualifier to exclude those paid orders whose only order line is a Measure Service
        sbQuery.Append("OR (PaymentType = 'FREE' AND EXISTS(SELECT * FROM ORD_LINES WHERE ORD_LINES.Order_No = ORD_HEADER.Order_No AND ProductName = 'VirtualFabric')))")       ' Qualifier to include those orders that are free, but have a virtual fabric to be imported
        sbQuery.Append(";")

        Dim sbBFSQuery As New System.Text.StringBuilder
        sbBFSQuery.Append("SELECT Count(Order_No) FROM ORD_HEADER ")
        sbBFSQuery.Append("WHERE OrderStatus = 'PAY_AUTH' ")
        sbBFSQuery.Append("AND ProductionSystemID IS NULL ")
        sbBFSQuery.Append("AND EXISTS(SELECT * FROM ORD_LINES WHERE ORD_LINES.Order_No = ORD_HEADER.Order_No) ")
        sbBFSQuery.Append(";")


        ' --- Adding list of sites to listbox ---
        For Each strSite As String In strSitesArray

            ' --- Selecting appropriate query ---
            Select Case strSite
                Case "BFSUK"
                    strSqlQuery = sbBFSQuery.ToString()
                Case Else
                    strSqlQuery = sbQuery.ToString()
            End Select


            ' --- Connecting to target site's database to get count of orders to import ---
            Try
                strConnection = ConfigurationManager.ConnectionStrings(strSite).ConnectionString

                Try
                    Using conn As SqlConnection = New SqlConnection(strConnection)
                        cmd = New SqlCommand(strSqlQuery, conn)
                        conn.Open()

                        strResult = cmd.ExecuteScalar
                    End Using

                    strFabricSamplesToImport = Importer.Get_FabricSamplesToImport_Count(strSite).ToString

                    lvwAvailableWebsites.Items.Add(New ListViewItem({strSite, strResult, strFabricSamplesToImport}))

                Catch ex As Exception
                    lvwAvailableWebsites.Items.Add(New ListViewItem({strSite, "~", "~"}))

                End Try

            Catch ex As Exception
                ' No connection string exists for this site. Do nothing
            End Try

        Next

    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        If IsNothing(class_ImportTimer) = True Then
            Me.Console_WriteMessage("Started import timer", True)
            class_ImportTimer = New ImportTimer()
            class_ImportTimer.StartTimer()
        Else
            Me.Console_WriteMessage("Restarted import timer", True)
            class_ImportTimer.StopTimer()
            class_ImportTimer = Nothing
            class_ImportTimer = New ImportTimer()
            class_ImportTimer.StartTimer()
        End If
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        class_ImportTimer.StopTimer()
        class_ImportTimer = Nothing
        Me.Console_WriteMessage("Stopped import timer", True)
    End Sub

    Private Sub btnImportOrders_Click(sender As Object, e As EventArgs) Handles btnImportOrders.Click
        If lvwAvailableWebsites.SelectedItems.Count >= 1 Then
            Dim lstOutputText As List(Of String)
            Dim strWebsite As String = lvwAvailableWebsites.SelectedItems(0).Text
            Me.Console_WriteMessage("Importing from site", lvwAvailableWebsites.SelectedItems(0).Text, True)
            lstOutputText = Importer.ImportOrders(lvwAvailableWebsites.SelectedItems(0).Text)
            For Each strOutput As String In lstOutputText
                Me.Console_WriteMessage(strOutput, strWebsite)
            Next
        Else
            Me.Console_WriteMessage("Please select a site")
        End If
        RefreshMainForm()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        RefreshMainForm()

    End Sub

    Private Sub btnImportSamples_Click(sender As Object, e As EventArgs) Handles btnImportSamples.Click

        If lvwAvailableWebsites.SelectedItems.Count >= 1 Then
            Dim strWebsite As String = lvwAvailableWebsites.SelectedItems(0).Text
            Try
                Importer.ImportFabricSamples(strWebsite)
                Me.Console_WriteMessage("Fabric samples imported", strWebsite)
            Catch ex As Exception
                Me.Console_WriteMessage(ex.Message & vbNewLine & ex.StackTrace, strWebsite)
            End Try

        Else
            Me.Console_WriteMessage("Please select a site")
        End If
        RefreshMainForm()

    End Sub


    Public Sub Console_WriteMessage(strMessage As String)
        Console_WriteMessage(strMessage, "", False)
    End Sub

    Public Sub Console_WriteMessage(strMessage As String, strWebsite As String)
        Console_WriteMessage(strMessage, strWebsite, False)
    End Sub

    Public Sub Console_WriteMessage(strMessage As String, blnIncludeTimestamp As Boolean)
        Console_WriteMessage(strMessage, "", blnIncludeTimestamp)
    End Sub

    Public Sub Console_WriteMessage(strMessage As String, strWebsite As String, blnIncludeTimestamp As Boolean)


        ' --- Defining constants, objects and variables ---
        Const MAX_MESSAGES As Integer = 100
        Dim sbConsoleContent As New StringBuilder
        Dim strLine As String = ""


        ' --- Adding new message to list ---
        If strMessage <> "" Then
            If blnIncludeTimestamp = True Then strLine += Now.ToShortTimeString()
            strLine += ":- "
            If strWebsite <> "" Then strLine += strWebsite & ": "
            strLine += strMessage
        End If
        class_lstConsoleMessages.Insert(0, strLine)


        ' --- Removing old lines from the list ---
        If class_lstConsoleMessages.Count > MAX_MESSAGES Then
            class_lstConsoleMessages.RemoveRange(MAX_MESSAGES, class_lstConsoleMessages.Count - MAX_MESSAGES)
        End If


        ' --- Writing list into console textbox ---
        For Each str As String In class_lstConsoleMessages
            sbConsoleContent.AppendLine(str)
        Next
        txtOutput.Text = sbConsoleContent.ToString()

    End Sub


End Class
