'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Class user_edit1
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents dept_text As System.Windows.Forms.TextBox
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents name1 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.dept_text = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.ok_button = New System.Windows.Forms.Button
        Me.cancel_button = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.name1 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'dept_text
        '
        Me.dept_text.Location = New System.Drawing.Point(16, 72)
        Me.dept_text.Name = "dept_text"
        Me.dept_text.Size = New System.Drawing.Size(272, 20)
        Me.dept_text.TabIndex = 15
        Me.dept_text.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 16)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Name"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Department"
        '
        'ok_button
        '
        Me.ok_button.BackColor = System.Drawing.SystemColors.Control
        Me.ok_button.Location = New System.Drawing.Point(24, 128)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.TabIndex = 25
        Me.ok_button.Text = "OK"
        '
        'cancel_button
        '
        Me.cancel_button.BackColor = System.Drawing.SystemColors.Control
        Me.cancel_button.Location = New System.Drawing.Point(200, 128)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.TabIndex = 26
        Me.cancel_button.Text = "Cancel"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(0, 112)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(304, 4)
        Me.Panel1.TabIndex = 27
        '
        'name1
        '
        Me.name1.Enabled = False
        Me.name1.Location = New System.Drawing.Point(16, 24)
        Me.name1.Name = "name1"
        Me.name1.Size = New System.Drawing.Size(272, 20)
        Me.name1.TabIndex = 28
        Me.name1.Text = ""
        '
        'user_edit1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(304, 166)
        Me.ControlBox = False
        Me.Controls.Add(Me.name1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dept_text)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "user_edit1"
        Me.Text = "User (Edit)"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ok_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ok_button.Click
        'Set the SQL string
        Dim objData As New DALBase
        Dim main As switch
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "edit_user"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection
        objData.AddParameter("@department", _
            OleDb.OleDbType.VarChar, 16, dept_text.Text)
        objData.AddParameter("@user", _
            OleDb.OleDbType.VarChar, 16, name1.Text)

        objData.Connection.ConnectionString = connect_string
        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery()
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("Add Task Failed")
            End If
            'Clear the input fields


        Catch ex As Exception

        End Try

        Me.Close()




    End Sub



    Private Sub user_edit1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ok_button.Enabled = False

    End Sub

    Private Sub name_text_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ok_button.Enabled = True
    End Sub

    Private Sub dept_text_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dept_text.TextChanged
        ok_button.Enabled = True
    End Sub

    Private Sub cancel_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel_button.Click
        Me.Close()

    End Sub
End Class
