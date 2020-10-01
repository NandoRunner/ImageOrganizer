Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports FAndradeTI.Util
Imports FAndradeTI.Util.FileSystem
Imports FAndradeTI.Util.WinForms
Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        WinReg.SubKey = "SOFTWARE\\" + Application.CompanyName + "\\" + Application.ProductName

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnOrganize As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderBrowserDialog2 As FolderBrowserDialog
    Friend WithEvents Panel1 As Panel
    Friend WithEvents txtPattern As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents btnDestino As Button
    Friend WithEvents txtTarget As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btnOrigem As Button
    Friend WithEvents txtSource As TextBox
    Friend WithEvents lblEvaluate As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnRemove As Button
    Friend WithEvents btnAdd As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents txtExtension As TextBox
    Friend WithEvents lbExtensions As ListBox
    Friend WithEvents btnEvaluatePattern As Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnOrganize = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.FolderBrowserDialog2 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtExtension = New System.Windows.Forms.TextBox()
        Me.lbExtensions = New System.Windows.Forms.ListBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblEvaluate = New System.Windows.Forms.Label()
        Me.btnEvaluatePattern = New System.Windows.Forms.Button()
        Me.txtPattern = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnDestino = New System.Windows.Forms.Button()
        Me.txtTarget = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnOrigem = New System.Windows.Forms.Button()
        Me.txtSource = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnOrganize
        '
        Me.btnOrganize.BackColor = System.Drawing.Color.Silver
        Me.btnOrganize.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOrganize.Image = CType(resources.GetObject("btnOrganize.Image"), System.Drawing.Image)
        Me.btnOrganize.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOrganize.Location = New System.Drawing.Point(417, 262)
        Me.btnOrganize.Name = "btnOrganize"
        Me.btnOrganize.Size = New System.Drawing.Size(286, 69)
        Me.btnOrganize.TabIndex = 0
        Me.btnOrganize.Text = "&Organize"
        Me.btnOrganize.UseVisualStyleBackColor = False
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Lavender
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.lblEvaluate)
        Me.Panel1.Controls.Add(Me.btnEvaluatePattern)
        Me.Panel1.Controls.Add(Me.txtPattern)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.btnDestino)
        Me.Panel1.Controls.Add(Me.btnOrganize)
        Me.Panel1.Controls.Add(Me.txtTarget)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnOrigem)
        Me.Panel1.Controls.Add(Me.txtSource)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(760, 537)
        Me.Panel1.TabIndex = 8
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.GroupBox1.Controls.Add(Me.btnRemove)
        Me.GroupBox1.Controls.Add(Me.btnAdd)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtExtension)
        Me.GroupBox1.Controls.Add(Me.lbExtensions)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(16, 202)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(373, 256)
        Me.GroupBox1.TabIndex = 18
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Image File Extensions"
        '
        'btnRemove
        '
        Me.btnRemove.BackColor = System.Drawing.Color.Silver
        Me.btnRemove.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemove.Location = New System.Drawing.Point(13, 211)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(129, 30)
        Me.btnRemove.TabIndex = 21
        Me.btnRemove.Text = "&Remove"
        Me.btnRemove.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.Silver
        Me.btnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.Location = New System.Drawing.Point(13, 126)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(129, 30)
        Me.btnAdd.TabIndex = 20
        Me.btnAdd.Text = "&Add >>"
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(9, 45)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(143, 24)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "New Extension:"
        '
        'txtExtension
        '
        Me.txtExtension.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExtension.Location = New System.Drawing.Point(13, 81)
        Me.txtExtension.Name = "txtExtension"
        Me.txtExtension.Size = New System.Drawing.Size(129, 26)
        Me.txtExtension.TabIndex = 15
        '
        'lbExtensions
        '
        Me.lbExtensions.FormattingEnabled = True
        Me.lbExtensions.ItemHeight = 24
        Me.lbExtensions.Location = New System.Drawing.Point(177, 45)
        Me.lbExtensions.Name = "lbExtensions"
        Me.lbExtensions.Size = New System.Drawing.Size(114, 196)
        Me.lbExtensions.TabIndex = 0
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(16, 479)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(725, 39)
        Me.ProgressBar1.TabIndex = 17
        '
        'lblEvaluate
        '
        Me.lblEvaluate.AutoSize = True
        Me.lblEvaluate.Location = New System.Drawing.Point(520, 148)
        Me.lblEvaluate.Name = "lblEvaluate"
        Me.lblEvaluate.Size = New System.Drawing.Size(0, 13)
        Me.lblEvaluate.TabIndex = 16
        '
        'btnEvaluatePattern
        '
        Me.btnEvaluatePattern.BackColor = System.Drawing.Color.Silver
        Me.btnEvaluatePattern.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEvaluatePattern.Location = New System.Drawing.Point(395, 139)
        Me.btnEvaluatePattern.Name = "btnEvaluatePattern"
        Me.btnEvaluatePattern.Size = New System.Drawing.Size(97, 28)
        Me.btnEvaluatePattern.TabIndex = 15
        Me.btnEvaluatePattern.Text = "&Evaluate"
        Me.btnEvaluatePattern.UseVisualStyleBackColor = False
        '
        'txtPattern
        '
        Me.txtPattern.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPattern.Location = New System.Drawing.Point(164, 140)
        Me.txtPattern.Name = "txtPattern"
        Me.txtPattern.Size = New System.Drawing.Size(225, 26)
        Me.txtPattern.TabIndex = 14
        Me.txtPattern.Text = "yyyy (MM) MMMM"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(25, 140)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(133, 24)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Folder Pattern:"
        '
        'btnDestino
        '
        Me.btnDestino.BackColor = System.Drawing.Color.DarkGray
        Me.btnDestino.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDestino.Image = CType(resources.GetObject("btnDestino.Image"), System.Drawing.Image)
        Me.btnDestino.Location = New System.Drawing.Point(709, 81)
        Me.btnDestino.Name = "btnDestino"
        Me.btnDestino.Size = New System.Drawing.Size(32, 26)
        Me.btnDestino.TabIndex = 12
        Me.btnDestino.Text = "..."
        Me.btnDestino.UseVisualStyleBackColor = False
        '
        'txtTarget
        '
        Me.txtTarget.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTarget.Location = New System.Drawing.Point(109, 81)
        Me.txtTarget.Name = "txtTarget"
        Me.txtTarget.Size = New System.Drawing.Size(594, 26)
        Me.txtTarget.TabIndex = 11
        Me.txtTarget.Text = "D:\Imagens\_Celular"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(25, 81)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 24)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Target:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(25, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 24)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Source:"
        '
        'btnOrigem
        '
        Me.btnOrigem.BackColor = System.Drawing.Color.DarkGray
        Me.btnOrigem.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOrigem.Image = CType(resources.GetObject("btnOrigem.Image"), System.Drawing.Image)
        Me.btnOrigem.Location = New System.Drawing.Point(709, 24)
        Me.btnOrigem.Name = "btnOrigem"
        Me.btnOrigem.Size = New System.Drawing.Size(32, 26)
        Me.btnOrigem.TabIndex = 8
        Me.btnOrigem.Text = "..."
        Me.btnOrigem.UseVisualStyleBackColor = False
        '
        'txtSource
        '
        Me.txtSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSource.Location = New System.Drawing.Point(109, 23)
        Me.txtSource.Name = "txtSource"
        Me.txtSource.Size = New System.Drawing.Size(594, 26)
        Me.txtSource.TabIndex = 7
        Me.txtSource.Text = "D:\Download\_Mobile Temp\MotoG3"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Image Sync & Organizer"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnDestino_Click(sender As Object, e As EventArgs) Handles btnDestino.Click

        'todo: buttons to explore folders
        txtTarget.Text = FS.GetFolder(txtTarget.Text, "Select target folder...")

    End Sub

    Private Sub btnOrigem_Click(sender As Object, e As EventArgs) Handles btnOrigem.Click

        txtSource.Text = FS.GetFolder(txtSource.Text, "Select source folder...")

    End Sub

    Private Sub frmImageSync_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = AppInfo.GetFullTitle()

        txtSource.Text = WinReg.Read("ImageOrganizerSource")
        txtTarget.Text = WinReg.Read("ImageOrganizerTarget")
        txtPattern.Text = WinReg.Read("ImageOrganizerPattern")

        Init()

    End Sub

    Private Sub cmdEvaluate_Click(sender As Object, e As EventArgs) Handles btnEvaluatePattern.Click
        lblEvaluate.Text = DateTime.Now.ToString(txtPattern.Text)
    End Sub

    Private Sub btnOrganize_Click(sender As Object, e As EventArgs) Handles btnOrganize.Click


        If String.IsNullOrEmpty(txtSource.Text) Or String.IsNullOrEmpty(txtTarget.Text) Then
            MessageBox.Show("Please check source and target folders!", "Organize command", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        Cursor.Hide()
        Dim targetDir As String = txtTarget.Text

        WinReg.Write("ImageOrganizerSource", txtSource.Text)
        WinReg.Write("ImageOrganizerTarget", targetDir)
        WinReg.Write("ImageOrganizerPattern", txtPattern.Text)

        'todo: write extensions to registry

        Dim extList() As String = lbExtensions.Items.Cast(Of String).ToArray()

        Dim fileList As New List(Of FileInfo)

        'todo: get extensions from list
        For Each ext As String In extList
            Dim fi = From f In (New DirectoryInfo(txtSource.Text)).GetFiles(ext).Cast(Of IO.FileInfo)() Order By f.CreationTime Select f
            fileList.AddRange(fi)
        Next

        ProgressBar1.Maximum = fileList.Count
        ProgressBar1.Minimum = 0
        ProgressBar1.Value = 0

        Dim newDirectory As String
        Dim filesMoved As Integer = 0
        Dim filesDeleted As Integer = 0

        For Each fileInfo As FileInfo In fileList
            newDirectory = targetDir & "\" & File.GetLastWriteTime(fileInfo.FullName).ToString(txtPattern.Text)
            If (Not Directory.Exists(newDirectory)) Then
                Directory.CreateDirectory(newDirectory)
            End If
            If (File.Exists(newDirectory & "\" & fileInfo.Name)) Then
                fileInfo.Delete()
                filesDeleted += 1
                Continue For
            End If
            File.Move(fileInfo.FullName, newDirectory & "\" & fileInfo.Name)
            filesMoved += 1
            ProgressBar1.Value += 1
            Application.DoEvents()
        Next

        Cursor.Show()

        MessageBox.Show(filesMoved & " images organized!", "Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Init()
        btnAdd.Enabled = False
        btnRemove.Enabled = False

        Dim aux As String = WinReg.Read("ImageOrganizerExtensions")

        If String.IsNullOrEmpty(aux) Then
            lbExtensions.Items.Add("*.jpg")
            lbExtensions.Items.Add("*.jpeg")
        Else
            'todo: split and add extensions
        End If

    End Sub

    Private Sub lbExtensions_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbExtensions.SelectedIndexChanged



        If lbExtensions.SelectedItem Is Nothing Then
            btnRemove.Enabled = False
        Else
            btnRemove.Enabled = True
        End If

    End Sub

    Private Sub txtExtension_TextChanged(sender As Object, e As EventArgs) Handles txtExtension.TextChanged

        If txtExtension.Text.Length < 3 Or txtExtension.Text.Length > 6 Then
            btnAdd.Enabled = False
        ElseIf txtExtension.Text.Substring(0, 2) <> "*." Then

            btnAdd.Enabled = False
        Else
            btnAdd.Enabled = True
        End If

    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        lbExtensions.Items.Remove(lbExtensions.SelectedItem)
        btnRemove.Enabled = False

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        lbExtensions.Items.Add(txtExtension.Text)
        btnAdd.Enabled = False
        txtExtension.Text = String.Empty

    End Sub
End Class
