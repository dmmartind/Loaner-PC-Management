'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Public Class pc_edit
    Inherits System.Windows.Forms.Form

    '////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////


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
    Friend WithEvents tag_text As System.Windows.Forms.TextBox
    Friend WithEvents mac_address As System.Windows.Forms.TextBox
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents name1 As System.Windows.Forms.TextBox
    Friend WithEvents id1 As System.Windows.Forms.TextBox
    Friend WithEvents tag_check As System.Windows.Forms.CheckBox
    Friend WithEvents mac_check As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(pc_edit))
        Me.tag_text = New System.Windows.Forms.TextBox
        Me.mac_address = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ok_button = New System.Windows.Forms.Button
        Me.cancel_button = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.name1 = New System.Windows.Forms.TextBox
        Me.id1 = New System.Windows.Forms.TextBox
        Me.tag_check = New System.Windows.Forms.CheckBox
        Me.mac_check = New System.Windows.Forms.CheckBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.SuspendLayout()
        '
        'tag_text
        '
        Me.tag_text.Location = New System.Drawing.Point(8, 72)
        Me.tag_text.Name = "tag_text"
        Me.tag_text.Size = New System.Drawing.Size(128, 20)
        Me.tag_text.TabIndex = 26
        Me.tag_text.Text = ""
        '
        'mac_address
        '
        Me.mac_address.Location = New System.Drawing.Point(168, 72)
        Me.mac_address.Name = "mac_address"
        Me.mac_address.Size = New System.Drawing.Size(128, 20)
        Me.mac_address.TabIndex = 27
        Me.mac_address.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 16)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Name"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Tag"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(168, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "MAC Address"
        '
        'ok_button
        '
        Me.ok_button.BackColor = System.Drawing.SystemColors.Control
        Me.ok_button.Location = New System.Drawing.Point(16, 128)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.TabIndex = 31
        Me.ok_button.Text = "OK"
        '
        'cancel_button
        '
        Me.cancel_button.BackColor = System.Drawing.SystemColors.Control
        Me.cancel_button.Location = New System.Drawing.Point(216, 128)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.TabIndex = 32
        Me.cancel_button.Text = "Cancel"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(-8, 104)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(328, 4)
        Me.Panel1.TabIndex = 33
        '
        'name1
        '
        Me.name1.Location = New System.Drawing.Point(8, 24)
        Me.name1.Name = "name1"
        Me.name1.Size = New System.Drawing.Size(288, 20)
        Me.name1.TabIndex = 34
        Me.name1.Text = ""
        '
        'id1
        '
        Me.id1.Location = New System.Drawing.Point(176, 112)
        Me.id1.Name = "id1"
        Me.id1.Size = New System.Drawing.Size(16, 20)
        Me.id1.TabIndex = 35
        Me.id1.Text = ""
        Me.id1.Visible = False
        '
        'tag_check
        '
        Me.tag_check.Location = New System.Drawing.Point(112, 56)
        Me.tag_check.Name = "tag_check"
        Me.tag_check.Size = New System.Drawing.Size(16, 16)
        Me.tag_check.TabIndex = 38
        '
        'mac_check
        '
        Me.mac_check.Location = New System.Drawing.Point(280, 56)
        Me.mac_check.Name = "mac_check"
        Me.mac_check.Size = New System.Drawing.Size(16, 16)
        Me.mac_check.TabIndex = 39
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(88, 48)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox1.TabIndex = 40
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(256, 48)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox2.TabIndex = 41
        Me.PictureBox2.TabStop = False
        '
        'pc_edit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(312, 158)
        Me.ControlBox = False
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.mac_check)
        Me.Controls.Add(Me.tag_check)
        Me.Controls.Add(Me.id1)
        Me.Controls.Add(Me.name1)
        Me.Controls.Add(Me.mac_address)
        Me.Controls.Add(Me.tag_text)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "pc_edit"
        Me.Text = "PC (Edit)"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ok_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ok_button.Click
        'Set the SQL string
        Dim objData As New DALBase
        Dim main As switch
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "edit_pc"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection

        objData.AddParameter("@tag", _
    OleDb.OleDbType.Integer, 16, tag_text.Text)
        objData.AddParameter("@mac", _
                    OleDb.OleDbType.VarChar, 16, mac_address.Text)
        objData.AddParameter("@id", _
                            OleDb.OleDbType.Integer, 16, id1.Text)
        objData.Connection.ConnectionString = connect_string
        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("EDIT Task Failed")
            End If
        Catch ex As Exception
            MessageBox.Show("This record has delete protection because it has relations with one or more tables.", "Relation Integrity Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ok_button.Enabled = True
    End Sub


    Private Sub cancel_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel_button.Click
        Me.Close()
    End Sub

    Private Sub pc_edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ok_button.Enabled = False
        tag_check.Checked = False
        mac_check.Checked = False
        tag_text.Enabled = False
        mac_address.Enabled = False

    End Sub

    Private Sub tag_text_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tag_text.TextChanged
        ok_button.Enabled = True
    End Sub

    Private Sub mac_address_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mac_address.TextChanged
        ok_button.Enabled = True
    End Sub

    Private Sub tag_check_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tag_check.CheckedChanged
        If tag_check.Checked = True Then
            tag_text.Enabled = True
        ElseIf tag_check.Checked = False Then
            tag_text.Enabled = False
        End If
    End Sub

    Private Sub mac_check_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mac_check.CheckedChanged
        If mac_check.Checked = True Then
            mac_address.Enabled = True
        ElseIf mac_check.Checked = False Then
            mac_address.Enabled = False
        End If
    End Sub
End Class
