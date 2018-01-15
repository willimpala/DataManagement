<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMain
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
        Me.MenuStripMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBarRunning = New System.Windows.Forms.ToolStripProgressBar()
        Me.TestToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProgressBarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OffToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.PictureBoxBizCoachMain = New System.Windows.Forms.PictureBox()
        Me.ListBoxCustomerInfo = New System.Windows.Forms.ListBox()
        Me.MenuStripMain.SuspendLayout()
        Me.StatusStripMain.SuspendLayout()
        CType(Me.PictureBoxBizCoachMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStripMain
        '
        Me.MenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.TestToolStripMenuItem})
        Me.MenuStripMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuStripMain.Name = "MenuStripMain"
        Me.MenuStripMain.Size = New System.Drawing.Size(1064, 24)
        Me.MenuStripMain.TabIndex = 0
        Me.MenuStripMain.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'StatusStripMain
        '
        Me.StatusStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBarRunning})
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 628)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Size = New System.Drawing.Size(1064, 22)
        Me.StatusStripMain.TabIndex = 1
        Me.StatusStripMain.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(39, 17)
        Me.ToolStripStatusLabel1.Text = "Status"
        '
        'ToolStripProgressBarRunning
        '
        Me.ToolStripProgressBarRunning.Name = "ToolStripProgressBarRunning"
        Me.ToolStripProgressBarRunning.Size = New System.Drawing.Size(200, 16)
        Me.ToolStripProgressBarRunning.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        '
        'TestToolStripMenuItem
        '
        Me.TestToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TimersToolStripMenuItem})
        Me.TestToolStripMenuItem.Name = "TestToolStripMenuItem"
        Me.TestToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.TestToolStripMenuItem.Text = "Test"
        '
        'TimersToolStripMenuItem
        '
        Me.TimersToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProgressBarToolStripMenuItem})
        Me.TimersToolStripMenuItem.Name = "TimersToolStripMenuItem"
        Me.TimersToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.TimersToolStripMenuItem.Text = "Timers"
        '
        'ProgressBarToolStripMenuItem
        '
        Me.ProgressBarToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OnToolStripMenuItem, Me.OffToolStripMenuItem})
        Me.ProgressBarToolStripMenuItem.Name = "ProgressBarToolStripMenuItem"
        Me.ProgressBarToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.ProgressBarToolStripMenuItem.Text = "ProgressBar - Status"
        '
        'OnToolStripMenuItem
        '
        Me.OnToolStripMenuItem.Name = "OnToolStripMenuItem"
        Me.OnToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.OnToolStripMenuItem.Text = "On"
        '
        'OffToolStripMenuItem
        '
        Me.OffToolStripMenuItem.Name = "OffToolStripMenuItem"
        Me.OffToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.OffToolStripMenuItem.Text = "Off"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 122)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(82, 13)
        Me.Label8.TabIndex = 61
        Me.Label8.Text = "Customer Name"
        '
        'PictureBoxBizCoachMain
        '
        Me.PictureBoxBizCoachMain.Image = Global.DataManagement.My.Resources.Resources.BizCoach_Our_Analytics_Logo_Low_Res_July_2016
        Me.PictureBoxBizCoachMain.InitialImage = Global.DataManagement.My.Resources.Resources.BizCoach_Our_Analytics_Logo_Low_Res_July_2016
        Me.PictureBoxBizCoachMain.Location = New System.Drawing.Point(12, 27)
        Me.PictureBoxBizCoachMain.Name = "PictureBoxBizCoachMain"
        Me.PictureBoxBizCoachMain.Size = New System.Drawing.Size(179, 91)
        Me.PictureBoxBizCoachMain.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBoxBizCoachMain.TabIndex = 62
        Me.PictureBoxBizCoachMain.TabStop = False
        '
        'ListBoxCustomerInfo
        '
        Me.ListBoxCustomerInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ListBoxCustomerInfo.FormattingEnabled = True
        Me.ListBoxCustomerInfo.Location = New System.Drawing.Point(12, 138)
        Me.ListBoxCustomerInfo.Name = "ListBoxCustomerInfo"
        Me.ListBoxCustomerInfo.Size = New System.Drawing.Size(222, 485)
        Me.ListBoxCustomerInfo.Sorted = True
        Me.ListBoxCustomerInfo.TabIndex = 60
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1064, 650)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.PictureBoxBizCoachMain)
        Me.Controls.Add(Me.ListBoxCustomerInfo)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.MenuStripMain)
        Me.MainMenuStrip = Me.MenuStripMain
        Me.Name = "FrmMain"
        Me.Text = "BizCoach 2.0 Database Automatic"
        Me.MenuStripMain.ResumeLayout(False)
        Me.MenuStripMain.PerformLayout()
        Me.StatusStripMain.ResumeLayout(False)
        Me.StatusStripMain.PerformLayout()
        CType(Me.PictureBoxBizCoachMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStripMain As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents StatusStripMain As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents TestToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TimersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProgressBarToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OnToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OffToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripProgressBarRunning As ToolStripProgressBar
    Friend WithEvents Label8 As Label
    Friend WithEvents PictureBoxBizCoachMain As PictureBox
    Friend WithEvents ListBoxCustomerInfo As ListBox
End Class
