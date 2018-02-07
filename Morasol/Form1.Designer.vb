<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxUsaNacionalidadSiNula = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBoxNacionalidadSiNula = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ListBoxHojas = New System.Windows.Forms.ListBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ButtonRutaExcel = New System.Windows.Forms.Button()
        Me.TextBoxRutaLibroExcel = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ListBoxDebug = New System.Windows.Forms.ListBox()
        Me.ButtonReservas = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBoxregistrosTratados = New System.Windows.Forms.TextBox()
        Me.ListBoxDebug2 = New System.Windows.Forms.ListBox()
        Me.ButtonDudasReservas = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPageReservas = New System.Windows.Forms.TabPage()
        Me.TabPageEquivalencias = New System.Windows.Forms.TabPage()
        Me.ButtonEquivalenciasEntidad = New System.Windows.Forms.Button()
        Me.ListBoxEquivalencias = New System.Windows.Forms.ListBox()
        Me.TabPageCardex = New System.Windows.Forms.TabPage()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPageReservas.SuspendLayout()
        Me.TabPageEquivalencias.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.CheckBoxUsaNacionalidadSiNula)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TextBoxNacionalidadSiNula)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.ListBoxHojas)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.ButtonRutaExcel)
        Me.GroupBox1.Controls.Add(Me.TextBoxRutaLibroExcel)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(725, 160)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'CheckBoxUsaNacionalidadSiNula
        '
        Me.CheckBoxUsaNacionalidadSiNula.AutoSize = True
        Me.CheckBoxUsaNacionalidadSiNula.Location = New System.Drawing.Point(116, 88)
        Me.CheckBoxUsaNacionalidadSiNula.Name = "CheckBoxUsaNacionalidadSiNula"
        Me.CheckBoxUsaNacionalidadSiNula.Size = New System.Drawing.Size(48, 17)
        Me.CheckBoxUsaNacionalidadSiNula.TabIndex = 29
        Me.CheckBoxUsaNacionalidadSiNula.Text = "Usar"
        Me.CheckBoxUsaNacionalidadSiNula.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 13)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "Nacionalidad si Nula"
        '
        'TextBoxNacionalidadSiNula
        '
        Me.TextBoxNacionalidadSiNula.Location = New System.Drawing.Point(116, 62)
        Me.TextBoxNacionalidadSiNula.Name = "TextBoxNacionalidadSiNula"
        Me.TextBoxNacionalidadSiNula.Size = New System.Drawing.Size(47, 20)
        Me.TextBoxNacionalidadSiNula.TabIndex = 27
        Me.TextBoxNacionalidadSiNula.Text = "ESP"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(493, 132)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(226, 13)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "IMEX=1  Lee los datos de Excel como TEXTO"
        '
        'ListBoxHojas
        '
        Me.ListBoxHojas.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBoxHojas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListBoxHojas.FormattingEnabled = True
        Me.ListBoxHojas.HorizontalExtent = 500
        Me.ListBoxHojas.HorizontalScrollbar = True
        Me.ListBoxHojas.Location = New System.Drawing.Point(267, 39)
        Me.ListBoxHojas.Name = "ListBoxHojas"
        Me.ListBoxHojas.Size = New System.Drawing.Size(263, 67)
        Me.ListBoxHojas.TabIndex = 25
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(203, 39)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Hoja Excel"
        '
        'ButtonRutaExcel
        '
        Me.ButtonRutaExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonRutaExcel.Image = CType(resources.GetObject("ButtonRutaExcel.Image"), System.Drawing.Image)
        Me.ButtonRutaExcel.Location = New System.Drawing.Point(536, 16)
        Me.ButtonRutaExcel.Name = "ButtonRutaExcel"
        Me.ButtonRutaExcel.Size = New System.Drawing.Size(44, 23)
        Me.ButtonRutaExcel.TabIndex = 21
        Me.ButtonRutaExcel.UseVisualStyleBackColor = True
        '
        'TextBoxRutaLibroExcel
        '
        Me.TextBoxRutaLibroExcel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxRutaLibroExcel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBoxRutaLibroExcel.Location = New System.Drawing.Point(71, 13)
        Me.TextBoxRutaLibroExcel.Name = "TextBoxRutaLibroExcel"
        Me.TextBoxRutaLibroExcel.Size = New System.Drawing.Size(459, 20)
        Me.TextBoxRutaLibroExcel.TabIndex = 19
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 13)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Libro Excel"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ListBoxDebug
        '
        Me.ListBoxDebug.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBoxDebug.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBoxDebug.FormattingEnabled = True
        Me.ListBoxDebug.HorizontalScrollbar = True
        Me.ListBoxDebug.ItemHeight = 14
        Me.ListBoxDebug.Location = New System.Drawing.Point(6, 6)
        Me.ListBoxDebug.Name = "ListBoxDebug"
        Me.ListBoxDebug.Size = New System.Drawing.Size(711, 172)
        Me.ListBoxDebug.TabIndex = 1
        '
        'ButtonReservas
        '
        Me.ButtonReservas.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonReservas.Location = New System.Drawing.Point(764, 202)
        Me.ButtonReservas.Name = "ButtonReservas"
        Me.ButtonReservas.Size = New System.Drawing.Size(75, 23)
        Me.ButtonReservas.TabIndex = 2
        Me.ButtonReservas.Text = "Reservas"
        Me.ButtonReservas.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 181)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Registros Tratados"
        '
        'TextBoxregistrosTratados
        '
        Me.TextBoxregistrosTratados.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TextBoxregistrosTratados.Location = New System.Drawing.Point(112, 181)
        Me.TextBoxregistrosTratados.Name = "TextBoxregistrosTratados"
        Me.TextBoxregistrosTratados.Size = New System.Drawing.Size(61, 20)
        Me.TextBoxregistrosTratados.TabIndex = 4
        '
        'ListBoxDebug2
        '
        Me.ListBoxDebug2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBoxDebug2.FormattingEnabled = True
        Me.ListBoxDebug2.Location = New System.Drawing.Point(12, 421)
        Me.ListBoxDebug2.Name = "ListBoxDebug2"
        Me.ListBoxDebug2.Size = New System.Drawing.Size(742, 108)
        Me.ListBoxDebug2.TabIndex = 5
        '
        'ButtonDudasReservas
        '
        Me.ButtonDudasReservas.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonDudasReservas.Location = New System.Drawing.Point(763, 157)
        Me.ButtonDudasReservas.Name = "ButtonDudasReservas"
        Me.ButtonDudasReservas.Size = New System.Drawing.Size(75, 39)
        Me.ButtonDudasReservas.TabIndex = 6
        Me.ButtonDudasReservas.Text = "Dudas Reservas"
        Me.ButtonDudasReservas.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPageReservas)
        Me.TabControl1.Controls.Add(Me.TabPageEquivalencias)
        Me.TabControl1.Controls.Add(Me.TabPageCardex)
        Me.TabControl1.Location = New System.Drawing.Point(12, 178)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(742, 237)
        Me.TabControl1.TabIndex = 7
        '
        'TabPageReservas
        '
        Me.TabPageReservas.Controls.Add(Me.ListBoxDebug)
        Me.TabPageReservas.Controls.Add(Me.Label3)
        Me.TabPageReservas.Controls.Add(Me.TextBoxregistrosTratados)
        Me.TabPageReservas.Location = New System.Drawing.Point(4, 22)
        Me.TabPageReservas.Name = "TabPageReservas"
        Me.TabPageReservas.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageReservas.Size = New System.Drawing.Size(734, 211)
        Me.TabPageReservas.TabIndex = 0
        Me.TabPageReservas.Text = "Reservas"
        Me.TabPageReservas.UseVisualStyleBackColor = True
        '
        'TabPageEquivalencias
        '
        Me.TabPageEquivalencias.Controls.Add(Me.ButtonEquivalenciasEntidad)
        Me.TabPageEquivalencias.Controls.Add(Me.ListBoxEquivalencias)
        Me.TabPageEquivalencias.Location = New System.Drawing.Point(4, 22)
        Me.TabPageEquivalencias.Name = "TabPageEquivalencias"
        Me.TabPageEquivalencias.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageEquivalencias.Size = New System.Drawing.Size(734, 211)
        Me.TabPageEquivalencias.TabIndex = 1
        Me.TabPageEquivalencias.Text = "Equivalencias"
        Me.TabPageEquivalencias.UseVisualStyleBackColor = True
        '
        'ButtonEquivalenciasEntidad
        '
        Me.ButtonEquivalenciasEntidad.Location = New System.Drawing.Point(6, 6)
        Me.ButtonEquivalenciasEntidad.Name = "ButtonEquivalenciasEntidad"
        Me.ButtonEquivalenciasEntidad.Size = New System.Drawing.Size(75, 23)
        Me.ButtonEquivalenciasEntidad.TabIndex = 1
        Me.ButtonEquivalenciasEntidad.Text = "Entidades"
        Me.ButtonEquivalenciasEntidad.UseVisualStyleBackColor = True
        '
        'ListBoxEquivalencias
        '
        Me.ListBoxEquivalencias.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBoxEquivalencias.FormattingEnabled = True
        Me.ListBoxEquivalencias.ItemHeight = 14
        Me.ListBoxEquivalencias.Location = New System.Drawing.Point(6, 33)
        Me.ListBoxEquivalencias.Name = "ListBoxEquivalencias"
        Me.ListBoxEquivalencias.Size = New System.Drawing.Size(722, 158)
        Me.ListBoxEquivalencias.TabIndex = 0
        '
        'TabPageCardex
        '
        Me.TabPageCardex.Location = New System.Drawing.Point(4, 22)
        Me.TabPageCardex.Name = "TabPageCardex"
        Me.TabPageCardex.Size = New System.Drawing.Size(734, 211)
        Me.TabPageCardex.TabIndex = 2
        Me.TabPageCardex.Text = "Cardex"
        Me.TabPageCardex.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(850, 539)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.ButtonDudasReservas)
        Me.Controls.Add(Me.ListBoxDebug2)
        Me.Controls.Add(Me.ButtonReservas)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPageReservas.ResumeLayout(False)
        Me.TabPageReservas.PerformLayout()
        Me.TabPageEquivalencias.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label5 As Label
    Friend WithEvents ButtonRutaExcel As Button
    Friend WithEvents TextBoxRutaLibroExcel As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents ListBoxHojas As ListBox
    Friend WithEvents ListBoxDebug As ListBox
    Friend WithEvents ButtonReservas As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBoxregistrosTratados As TextBox
    Friend WithEvents ListBoxDebug2 As ListBox
    Friend WithEvents ButtonDudasReservas As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBoxNacionalidadSiNula As TextBox
    Friend WithEvents CheckBoxUsaNacionalidadSiNula As CheckBox
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPageReservas As TabPage
    Friend WithEvents TabPageEquivalencias As TabPage
    Friend WithEvents ListBoxEquivalencias As ListBox
    Friend WithEvents ButtonEquivalenciasEntidad As Button
    Friend WithEvents TabPageCardex As TabPage
End Class
