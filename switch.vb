'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Imports System
Imports System.IO


Public Class switch
    Inherits System.Windows.Forms.Form
    Public Shared mystream As String
    Dim current_stream As String
    Public great2 As New splash
    Private mouse_offset As Point



#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Lookup_Tables As System.Windows.Forms.Button
    Friend WithEvents Check_Out_Lookup As System.Windows.Forms.Button
    Friend WithEvents Check_in_out As System.Windows.Forms.Button
    Friend WithEvents Check_in_out_click As System.Windows.Forms.Button
    Friend WithEvents check_in As System.Windows.Forms.Button
    Friend WithEvents history_search As System.Windows.Forms.Button
    Friend WithEvents min As System.Windows.Forms.Button
    Friend WithEvents database As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(switch))
        Me.Lookup_Tables = New System.Windows.Forms.Button
        Me.Check_Out_Lookup = New System.Windows.Forms.Button
        Me.min = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.check_in = New System.Windows.Forms.Button
        Me.history_search = New System.Windows.Forms.Button
        Me.database = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Lookup_Tables
        '
        Me.Lookup_Tables.BackColor = System.Drawing.Color.Silver
        Me.Lookup_Tables.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Lookup_Tables.Location = New System.Drawing.Point(88, 192)
        Me.Lookup_Tables.Name = "Lookup_Tables"
        Me.Lookup_Tables.Size = New System.Drawing.Size(88, 23)
        Me.Lookup_Tables.TabIndex = 2
        Me.Lookup_Tables.Text = "Lookup Tables"
        Me.Lookup_Tables.UseVisualStyleBackColor = False
        '
        'Check_Out_Lookup
        '
        Me.Check_Out_Lookup.BackColor = System.Drawing.Color.Silver
        Me.Check_Out_Lookup.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Check_Out_Lookup.Location = New System.Drawing.Point(424, 184)
        Me.Check_Out_Lookup.Name = "Check_Out_Lookup"
        Me.Check_Out_Lookup.Size = New System.Drawing.Size(75, 23)
        Me.Check_Out_Lookup.TabIndex = 3
        Me.Check_Out_Lookup.Text = "Check_Out Lookup"
        Me.Check_Out_Lookup.UseVisualStyleBackColor = False
        '
        'min
        '
        Me.min.BackgroundImage = CType(resources.GetObject("min.BackgroundImage"), System.Drawing.Image)
        Me.min.Location = New System.Drawing.Point(256, 72)
        Me.min.Name = "min"
        Me.min.Size = New System.Drawing.Size(24, 24)
        Me.min.TabIndex = 4
        '
        'Button6
        '
        Me.Button6.BackgroundImage = CType(resources.GetObject("Button6.BackgroundImage"), System.Drawing.Image)
        Me.Button6.Location = New System.Drawing.Point(312, 72)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(24, 24)
        Me.Button6.TabIndex = 5
        '
        'check_in
        '
        Me.check_in.BackColor = System.Drawing.Color.Silver
        Me.check_in.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.check_in.Location = New System.Drawing.Point(368, 144)
        Me.check_in.Name = "check_in"
        Me.check_in.Size = New System.Drawing.Size(128, 23)
        Me.check_in.TabIndex = 6
        Me.check_in.Text = "Check_In/Check_Out"
        Me.check_in.UseVisualStyleBackColor = False
        '
        'history_search
        '
        Me.history_search.BackColor = System.Drawing.Color.Silver
        Me.history_search.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.history_search.Location = New System.Drawing.Point(88, 144)
        Me.history_search.Name = "history_search"
        Me.history_search.Size = New System.Drawing.Size(88, 23)
        Me.history_search.TabIndex = 7
        Me.history_search.Text = "History Search"
        Me.history_search.UseVisualStyleBackColor = False
        '
        'database
        '
        Me.database.BackColor = System.Drawing.Color.Silver
        Me.database.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.database.Location = New System.Drawing.Point(248, 264)
        Me.database.Name = "database"
        Me.database.Size = New System.Drawing.Size(75, 23)
        Me.database.TabIndex = 8
        Me.database.Text = "Database"
        Me.database.UseVisualStyleBackColor = False
        '
        'switch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.WindowText
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(576, 352)
        Me.Controls.Add(Me.database)
        Me.Controls.Add(Me.history_search)
        Me.Controls.Add(Me.check_in)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.min)
        Me.Controls.Add(Me.Check_Out_Lookup)
        Me.Controls.Add(Me.Lookup_Tables)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "switch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Loaner PC Manager"
        Me.TransparencyKey = System.Drawing.Color.Black
        Me.ResumeLayout(False)

    End Sub

#End Region
    

    


    

    


    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

       


    End Sub



    Private Sub switch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim formop As New switch
        formop.Opacity = 1.0
        formop.TransparencyKey = Color.Black

        database.Enabled = False
        history_search.Enabled = True
        Lookup_Tables.Enabled = True
        check_in.Enabled = True
        Check_Out_Lookup.Enabled = True
    'AxSkin1.ApplySkin(0)
    'Dim great1 As New switch
    'great1.Visible = True
        great2.Show()

        Try

    ' Create an instance of StreamReader to read from a file.
    Dim sr As StreamReader = New StreamReader("log.txt")
    Dim line As String
    ' Read and display the lines from the file until the end 
    ' of the file is reached.
            mystream = sr.ReadLine()
            mystream = mystream.Remove(0, 1)
            mystream = mystream.Remove(mystream.Length - 1, 1)
    REM  MessageBox.Show(mystream, "HIO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            sr.Close()
        Catch ex As Exception
            MessageBox.Show("Cannot find database", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            database.Enabled = True
            history_search.Enabled = False
            Lookup_Tables.Enabled = False
            check_in.Enabled = False
            Check_Out_Lookup.Enabled = False

        End Try




    End Sub

    Private Sub close_form()
        great2.Close()

    End Sub

    Private Function AppPath() As String
        Return System.Windows.Forms.Application.StartupPath
    End Function

    Private Sub Form1_MouseDown(ByVal sender As Object, _
        ByVal e As System.Windows.Forms.MouseEventArgs) _
        Handles MyBase.MouseDown
        mouse_offset = New Point(-e.X, -e.Y)
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.MouseEventArgs) _
    Handles MyBase.MouseMove
        If e.Button = MouseButtons.Left Then
            Dim mousePos As Point = Control.MousePosition
            mousePos.Offset(mouse_offset.X, mouse_offset.Y)
            Location = mousePos
        End If
    End Sub











    Private Sub check_in_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles check_in.Click
        Dim check_in_out As New loner_pc
        check_in_out.Show()
    End Sub

    Private Sub Lookup_Tables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Lookup_Tables.Click

    End Sub

    Private Sub Lookup_Tables_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Lookup_Tables.Click
        Dim table_managment As New lookup_tables
        table_managment.Show()
    End Sub

    Private Sub Check_Out_Lookup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Check_Out_Lookup.Click
        Dim checked_out_view As New lookup_tables_check
        checked_out_view.Show()
    End Sub

    Private Sub history_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles history_search.Click
        Dim history1 As New history_lookup

        history1.Show()
    End Sub

    Private Sub min_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles min.Click
        Dim form5 As New switch
        WindowState = FormWindowState.Minimized

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub


    Private Sub database_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles database.Click
        Dim openFileDialog1 As New OpenFileDialog
        openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = "MS Access|*.mdb"
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            mystream = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='" & openFileDialog1.FileName() & "'"

            current_stream = AppPath() + "\log.txt"
            FileOpen(1, current_stream, OpenMode.Append)
            WriteLine(1, mystream)
            FileClose(1)
            database.Enabled = False
            history_search.Enabled = True
            Lookup_Tables.Enabled = True
            check_in.Enabled = True
            Check_Out_Lookup.Enabled = True
            Try

                ' Create an instance of StreamReader to read from a file.
                Dim sr As StreamReader = New StreamReader(current_stream)
                Dim line As String
                ' Read and display the lines from the file until the end 
                ' of the file is reached.
                mystream = sr.ReadLine()
                mystream = mystream.Remove(0, 1)
                mystream = mystream.Remove(mystream.Length - 1, 1)
                REM  MessageBox.Show(mystream, "HIO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                sr.Close()
            Catch ex As Exception
                MessageBox.Show("Cannot find database", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                database.Enabled = True
                history_search.Enabled = False
                Lookup_Tables.Enabled = False
                check_in.Enabled = False
                Check_Out_Lookup.Enabled = False

            End Try


        End If
    End Sub
End Class
