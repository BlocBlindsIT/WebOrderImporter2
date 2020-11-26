<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.btnImportOrders = New System.Windows.Forms.Button()
        Me.lblOutput = New System.Windows.Forms.Label()
        Me.lvwAvailableWebsites = New System.Windows.Forms.ListView()
        Me.Website = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.OrdersToImport = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SamplesToImport = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnImportSamples = New System.Windows.Forms.Button()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtOutput
        '
        Me.txtOutput.Location = New System.Drawing.Point(12, 210)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ReadOnly = True
        Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOutput.Size = New System.Drawing.Size(359, 239)
        Me.txtOutput.TabIndex = 0
        Me.txtOutput.Text = "Ready..."
        '
        'btnImportOrders
        '
        Me.btnImportOrders.Location = New System.Drawing.Point(12, 154)
        Me.btnImportOrders.Name = "btnImportOrders"
        Me.btnImportOrders.Size = New System.Drawing.Size(96, 23)
        Me.btnImportOrders.TabIndex = 1
        Me.btnImportOrders.Text = "Import Orders"
        Me.btnImportOrders.UseVisualStyleBackColor = True
        '
        'lblOutput
        '
        Me.lblOutput.AutoSize = True
        Me.lblOutput.Location = New System.Drawing.Point(12, 194)
        Me.lblOutput.Name = "lblOutput"
        Me.lblOutput.Size = New System.Drawing.Size(48, 13)
        Me.lblOutput.TabIndex = 2
        Me.lblOutput.Text = "Console:"
        '
        'lvwAvailableWebsites
        '
        Me.lvwAvailableWebsites.AutoArrange = False
        Me.lvwAvailableWebsites.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Website, Me.OrdersToImport, Me.SamplesToImport})
        Me.lvwAvailableWebsites.HideSelection = False
        Me.lvwAvailableWebsites.Location = New System.Drawing.Point(12, 51)
        Me.lvwAvailableWebsites.Name = "lvwAvailableWebsites"
        Me.lvwAvailableWebsites.Scrollable = False
        Me.lvwAvailableWebsites.Size = New System.Drawing.Size(359, 97)
        Me.lvwAvailableWebsites.TabIndex = 6
        Me.lvwAvailableWebsites.UseCompatibleStateImageBehavior = False
        Me.lvwAvailableWebsites.View = System.Windows.Forms.View.Details
        '
        'Website
        '
        Me.Website.Text = "Website"
        Me.Website.Width = 117
        '
        'OrdersToImport
        '
        Me.OrdersToImport.Text = "Orders To Import"
        Me.OrdersToImport.Width = 107
        '
        'SamplesToImport
        '
        Me.SamplesToImport.Text = "Sample Orders to Import"
        Me.SamplesToImport.Width = 143
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(312, 153)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(59, 24)
        Me.btnRefresh.TabIndex = 7
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'btnImportSamples
        '
        Me.btnImportSamples.Location = New System.Drawing.Point(114, 153)
        Me.btnImportSamples.Name = "btnImportSamples"
        Me.btnImportSamples.Size = New System.Drawing.Size(96, 23)
        Me.btnImportSamples.TabIndex = 8
        Me.btnImportSamples.Text = "Import Samples"
        Me.btnImportSamples.UseVisualStyleBackColor = True
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(12, 13)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(113, 13)
        Me.lblVersion.TabIndex = 9
        Me.lblVersion.Text = "Version: {0}.{1}.{2}.{3}"
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(282, 154)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(24, 23)
        Me.btnStop.TabIndex = 10
        Me.btnStop.Text = "■"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(251, 154)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(25, 23)
        Me.btnStart.TabIndex = 11
        Me.btnStart.Text = "▶"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(383, 461)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.btnImportSamples)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.lvwAvailableWebsites)
        Me.Controls.Add(Me.lblOutput)
        Me.Controls.Add(Me.btnImportOrders)
        Me.Controls.Add(Me.txtOutput)
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.Name = "Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Web Order Importer"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtOutput As TextBox
    Friend WithEvents btnImportOrders As Button
    Friend WithEvents lblOutput As Label
    Friend WithEvents lvwAvailableWebsites As ListView
    Friend WithEvents Website As ColumnHeader
    Friend WithEvents OrdersToImport As ColumnHeader
    Friend WithEvents btnRefresh As Button
    Friend WithEvents SamplesToImport As ColumnHeader
    Friend WithEvents btnImportSamples As Button
    Friend WithEvents lblVersion As Label
    Friend WithEvents btnStop As Button
    Friend WithEvents btnStart As Button
End Class
