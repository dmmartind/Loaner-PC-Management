'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Public Class pc_add
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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents name_text As System.Windows.Forms.TextBox
    Friend WithEvents mac_text As System.Windows.Forms.TextBox
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents tag_text As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.name_text = New System.Windows.Forms.TextBox
        Me.mac_text = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ok_button = New System.Windows.Forms.Button
        Me.cancel_button = New System.Windows.Forms.Button
        Me.tag_text = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'name_text
        '
        Me.name_text.Location = New System.Drawing.Point(8, 24)
        Me.name_text.Name = "name_text"
        Me.name_text.Size = New System.Drawing.Size(272, 20)
        Me.name_text.TabIndex = 12
        Me.name_text.Text = ""
        '
        'mac_text
        '
        Me.mac_text.Location = New System.Drawing.Point(160, 72)
        Me.mac_text.Name = "mac_text"
        Me.mac_text.Size = New System.Drawing.Size(128, 20)
        Me.mac_text.TabIndex = 14
        Me.mac_text.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 16)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Name"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Tag"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(160, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "MAC Address"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(-8, 112)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(304, 4)
        Me.Panel1.TabIndex = 21
        '
        'ok_button
        '
        Me.ok_button.BackColor = System.Drawing.SystemColors.Control
        Me.ok_button.Location = New System.Drawing.Point(16, 136)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.TabIndex = 23
        Me.ok_button.Text = "OK"
        '
        'cancel_button
        '
        Me.cancel_button.BackColor = System.Drawing.SystemColors.Control
        Me.cancel_button.Location = New System.Drawing.Point(200, 136)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.TabIndex = 24
        Me.cancel_button.Text = "Cancel"
        '
        'tag_text
        '
        Me.tag_text.Location = New System.Drawing.Point(8, 72)
        Me.tag_text.Name = "tag_text"
        Me.tag_text.Size = New System.Drawing.Size(128, 20)
        Me.tag_text.TabIndex = 25
        Me.tag_text.Text = ""
        '
        'pc_add
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(292, 174)
        Me.ControlBox = False
        Me.Controls.Add(Me.tag_text)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.mac_text)
        Me.Controls.Add(Me.name_text)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "pc_add"
        Me.Text = "PC (Add)"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub ok_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ok_button.Click
        'Set the SQL string
        Dim objData As New DALBase
        Dim main As switch
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "add_pc"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection
        objData.AddParameter("@name", _
            OleDb.OleDbType.VarChar, 16, name_text.Text)
        objData.AddParameter("@tag", _
            OleDb.OleDbType.VarChar, 16, tag_text.Text)
        objData.AddParameter("@mac", _
                    OleDb.OleDbType.VarChar, 16, mac_text.Text)
        'Open the database connection
        objData.Connection.ConnectionString = connect_string

        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery()
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("Delete Task Failed")
            End If
        Catch ex As Exception
            MessageBox.Show("This record has delete protection because it has relations with one or more tables.", "Relation Integrity Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        Me.Close()


    End Sub

    Private Sub pc_add_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ok_button.Enabled = False
    End Sub


    Private Sub name_text_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles name_text.TextChanged
        ok_button.Enabled = True

    End Sub

    Private Sub cancel_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel_button.Click
        Dim form As New pc_add
        Me.Close()
    End Sub

    Private Sub tag_text_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tag_text.TextChanged
        ok_button.Enabled = True
    End Sub

    Private Sub mac_text_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mac_text.TextChanged
        ok_button.Enabled = True
    End Sub
End Class
