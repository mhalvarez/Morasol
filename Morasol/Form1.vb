Imports System.IO
Imports System.Configuration
Public Class Form1



    Private m_DBMORASOL As C_DATOS.C_DatosOledb
    Private m_DBHOTEL As C_DATOS.C_DatosOledb

    '
    Private Linea As Integer
    Private SQL As String
    Private m_ResultInt As Integer
    Private m_ResultStr As String
    ' excel
    Dim WithEvents ExcelApp As Microsoft.Office.Interop.Excel.Application
    Dim ExcelBook As Microsoft.Office.Interop.Excel.Workbook
    Dim ExcelSheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim Rango As Microsoft.Office.Interop.Excel.Range
    Private m_File As FileInfo

    '
    Private Filegraba As StreamWriter
    Private FileEstaOk As Boolean = False

    Private mHayErrores As Boolean = False

    Private m_AuxStr As String
    Private m_AuxInt As Integer
    Private m_PosibleNombre As String
    Private m_PosibleApellido As String
    Private m_Hotel As String


    Private DS As System.Data.DataSet
    Private DT As System.Data.DataTable
    Private DR As System.Data.DataRow


    Private MyCommand As System.Data.OleDb.OleDbDataAdapter
    Private MyConnection As System.Data.OleDb.OleDbConnection



    Private Sub ButtonRutaExcel_Click(sender As Object, e As EventArgs) Handles ButtonRutaExcel.Click


        Try
            OpenFileDialog1.CheckPathExists = True
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|Excel files (*.xls)|*.xls|All files (*.*)|*.*"
            OpenFileDialog1.FilterIndex = 1
            OpenFileDialog1.RestoreDirectory = True

            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                If Me.OpenFileDialog1.FileName.Length > 0 Then
                    Me.TextBoxRutaLibroExcel.Text = Me.OpenFileDialog1.FileName
                    Me.Cursor = Cursors.AppStarting
                    Me.ContarHojasExcel(Me.TextBoxRutaLibroExcel.Text)
                End If
            End If



        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

            Me.Cursor = Cursors.Default

        End Try
    End Sub
    Private Sub ContarHojasExcel(ByVal vLibro As String)
        Try
            Me.ListBoxHojas.Items.Clear()
            Me.ListBoxHojas.Update()

            If IsNothing(Me.ExcelApp) = True Then
                Me.ExcelApp = New Microsoft.Office.Interop.Excel.Application
            End If

            Me.ExcelApp.IgnoreRemoteRequests = True


            '  Me.ExcelBook = Me.ExcelApp.Workbooks.Open(Me.TextBoxRutaLibroExcel.Text, , True, , , , True, )

            Me.ExcelBook = Me.ExcelApp.Workbooks.Add(vLibro)

            '  Me.ExcelApp.Visible = False
            Dim i As Integer
            Dim P As Integer

            For Each Me.ExcelSheet In Me.ExcelApp.Sheets
                Me.ListBoxHojas.Items.Add(Me.ExcelSheet.Name)
                i = i + 1

                For P = 1 To Me.ExcelSheet.Name.Length

                    If Mid(Me.ExcelSheet.Name, P, 1) = "." Then
                        MsgBox("No use punto(.) en el nombre de la Hoja de Excel a Tratar = " & Me.ExcelSheet.Name, MsgBoxStyle.Exclamation, "Atención")
                    End If

                Next P
            Next Me.ExcelSheet



        Catch e As System.Runtime.InteropServices.COMException
            MsgBox("Es posible que esta Computadora no tenga instalado Microsoft Excel " & vbCrLf & e.Message, MsgBoxStyle.Information, " Exepción COM Puede Continuar ...")

        Catch ex As Exception
            MsgBox("Es posible que esta Computadora no tenga instalado Microsoft Excel " & vbCrLf & ex.Message, MsgBoxStyle.Information, "Exepción Genérica Puede Continuar ...")
        End Try

    End Sub
    Private Sub LeeExcelReservas(ByVal vFileName As String, vHojaName As String)
        Try
            '  Dim DS As System.Data.DataSet
            '  Dim DT As System.Data.DataTable
            '  Dim DR As System.Data.DataRow


            '  Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            '  Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim StrConexion As String

            Dim Ind As Integer



            Me.Linea = 1

            Me.m_Hotel = InputBox("Hotel a Tratar", "AP MORASOL, ATLANTICO,SUITES", "SUITES")

            '  StrConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vFileName & ";Extended Properties=""Excel 8.0;HDR=Yes"""
            StrConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vFileName & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
            MyConnection = New System.Data.OleDb.OleDbConnection(StrConexion)
            DS = New System.Data.DataSet()
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & vHojaName & "$] where ([Estado reserva] = 'RESERVA' or [Estado reserva] = 'ENTRADA') AND [HOTEL] = '" & Me.m_Hotel & "'", MyConnection)
            '    "select * from [MAYO 2010$]", MyConnection)

            MyCommand.Fill(DS)
            DT = DS.Tables(0)

            ' DT.Locale.NumberFormat.NumberDecimalSeparator = "."
            ' DT.Locale.NumberFormat.NumberGroupSeparator = ","

            Dim Linea As String
            Dim LineaGraba As String = ""

            Dim r_Hotel As String
            Dim r_Ejercicio As String
            Dim r_Reserva As String
            Dim r_Desglose As String
            Dim r_ReservaAnterior As String
            Dim r_FechaCreacion As String
            Dim r_Cliente As String
            Dim r_TtoName As String
            Dim r_Agencia As String
            Dim r_Llegada As String
            Dim r_Salida As String
            Dim r_Hllegada As String
            Dim r_Hsalida As String


            Dim r_VueloLlegada As String
            Dim r_VueloSalida As String
            Dim r_Nacionalidad As String
            Dim r_Apellido As String
            Dim r_Trato As String
            Dim r_Bono As String

            Dim r_TipoHabUso As String
            Dim r_TipoHabFra As String
            Dim r_TipoRegimen As String

            Dim r_Adultos As String
            Dim r_Junior As String
            Dim r_Ninos As String
            Dim r_Cunas As String

            Dim r_TipoFactura As String

            Dim r_EstadoReserva As String

            Dim r_ObservacionContrato As String


            ' aux

            Dim T As TimeSpan

            Me.ListBoxDebug.Items.Clear()


            Me.m_DBMORASOL.IniciaTransaccion()


            Me.mHayErrores = False

            For Each DR In DT.Rows


                Ind = Ind + 1


                r_Hotel = DR(0).ToString
                r_Ejercicio = DR(1).ToString
                r_Reserva = DR(2).ToString
                r_Desglose = DR(3).ToString

                If CInt(r_Desglose) > 1 Then
                    MsgBox("Existen Reservas de Grupo Reserva = " & r_Reserva)
                End If

                '
                r_ReservaAnterior = r_Reserva & "/" & r_Ejercicio
                '

                r_FechaCreacion = DR(4).ToString
                r_Cliente = DR(6).ToString
                r_TtoName = DR(7).ToString
                r_Agencia = DR(8).ToString
                r_Llegada = DR(9).ToString
                r_Salida = DR(10).ToString




                If DR(11).ToString.Length > 0 Then

                    r_Hllegada = Format(CDate(r_Llegada), "dd/MM/yyyy") & " " & Format(CDate(DR(11)), "H:mm")
                Else
                    r_Hllegada = ""
                End If


                If DR(12).ToString.Length > 0 Then
                    r_Salida = Format(CDate(r_Salida), "dd/MM/yyyy") & " " & Format(CDate(DR(12)), "H:mm")
                Else
                    r_Hsalida = ""
                End If


                r_VueloLlegada = DR(13).ToString
                r_VueloSalida = DR(14).ToString
                r_Nacionalidad = DR(16).ToString
                r_Apellido = DR(17).ToString

                ' test
                r_Apellido = Me.ExtraeNombre(r_Apellido)
                '

                r_Trato = DR(18).ToString
                r_Bono = DR(19).ToString



                r_TipoHabUso = DR(20).ToString
                r_TipoHabFra = DR(21).ToString
                r_TipoRegimen = DR(22).ToString

                r_Adultos = DR(24).ToString
                r_Junior = DR(25).ToString
                r_Ninos = DR(26).ToString
                r_Cunas = DR(27).ToString
                r_TipoFactura = DR(28).ToString
                r_EstadoReserva = DR(29).ToString


                r_ObservacionContrato = DR(41).ToString



                Linea = Ind.ToString.PadRight(10, " ") & "|" & r_Hotel.PadRight(10, " ") & "|" & r_EstadoReserva.PadRight(10, " ") & "|" & r_Ejercicio.PadRight(10, " ") & "|" &
                r_Reserva.ToString.PadRight(10, " ") & "|" & r_FechaCreacion.PadRight(25, " ") & "|" & r_Cliente.PadRight(20, " ") & "|" &
                r_TtoName.ToString.PadRight(50, " ") & "|" & r_Agencia.ToString.PadRight(50, " ") & "|" & r_Llegada.ToString.PadRight(25, " ") & "|" &
                r_Salida.ToString.PadRight(25, " ") & "|" & r_Hllegada.ToString.PadRight(19, " ") & "|" & r_Hsalida.ToString.PadRight(19, " ") & "|" &
                r_VueloLlegada.ToString.PadRight(15, " ") & "|" & r_VueloSalida.ToString.PadRight(15, " ") & "|" & r_Nacionalidad.ToString.PadRight(10, " ") & "|" &
              DR(17).PadRight(50, " ") & "|" &
                r_Apellido.ToString.PadRight(50, " ") & "|" & r_Trato.ToString.PadRight(10, " ") & "|" & Mid(r_Bono.ToString, 1, 25).PadRight(25, " ") & "|" &
                r_TipoHabUso.ToString.PadRight(10, " ") & "|" & r_TipoHabFra.ToString.PadRight(10, " ") & "|" & r_TipoRegimen.ToString.PadRight(25, " ") & "|" &
                r_Adultos.ToString.PadRight(10, " ") & "|" & r_Junior.ToString.PadRight(10, " ") & "|" & r_Ninos.ToString.PadRight(25, " ") & "|" & r_Cunas.ToString.PadRight(25, " ") & "|"

                Me.ListBoxDebug.Items.Add(Linea)
                ' 







                ' ALGUNOS CONTROLES

                If r_Bono.Length > 25 Then
                    MessageBox.Show("Longitud del Numero de Bono excedida (25) en reserva" & r_ReservaAnterior & vbCrLf & r_Bono, "Atención FIN DEL PROCESO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.mHayErrores = True
                    Exit For
                End If



                SQL = "SELECT NACI_CODI FROM TNHT_NACI WHERE NACI_CISO = '" & r_Nacionalidad & "'"
                Me.m_ResultStr = Me.m_DBHOTEL.EjecutaSqlScalar(SQL)
                If IsNothing(Me.m_ResultStr) = True Then
                    SQL = "SELECT NACI_CODI FROM TNHT_NACI WHERE NACI_CODI = '" & r_Nacionalidad & "'"
                    Me.m_ResultStr = Me.m_DBHOTEL.EjecutaSqlScalar(SQL)
                End If

                If IsNothing(Me.m_ResultStr) = True And Me.CheckBoxUsaNacionalidadSiNula.Checked Then
                    Me.m_ResultStr = Me.TextBoxNacionalidadSiNula.Text
                End If



                If IsNothing(Me.m_ResultStr) = False Then
                    r_Nacionalidad = Me.m_ResultStr
                Else
                    MessageBox.Show("Nacionalidad No Válida " & r_Nacionalidad & " en reserva" & vbCrLf & r_ReservaAnterior, "Atención FIN DEL PROCESO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.mHayErrores = True
                    Exit For
                End If






                ' inserta oracle
                SQL = " INSERT INTO TRESE ( "
                SQL = SQL & "   TRESE_CODI, TRESE_ANCI,TRESE_DESGLOSE, TRESE_HOTEL, "
                SQL = SQL & "   TRESE_DACR, TRESE_CLIE, TRESE_TTOO,TRESE_AGENCIA, "
                SQL = SQL & "   TRESE_DAEN, TRESE_DASA,TNACI_CODI,TRESE_THABRESE,TRESE_THABOCUP,TRESE_TREG,TRESE_NUAD,TRESE_NUCR,TRESE_NUIN,TRESE_NUCU)  "
                SQL = SQL & "VALUES ("
                SQL = SQL & r_Reserva & ","
                SQL = SQL & r_Ejercicio & ","
                SQL = SQL & r_Desglose & ",'"
                SQL = SQL & r_Hotel & "','"
                SQL = SQL & Format(CDate(r_FechaCreacion), "dd/MM/yyyy") & "','"
                SQL = SQL & r_Cliente & "','"
                SQL = SQL & r_TtoName & "','"
                SQL = SQL & r_Agencia & "','"
                SQL = SQL & Format(CDate(r_Llegada), "dd/MM/yyyy") & "','"
                SQL = SQL & Format(CDate(r_Salida), "dd/MM/yyyy") & "','"
                SQL = SQL & r_Nacionalidad & "','"
                SQL = SQL & r_TipoHabFra & "','"
                SQL = SQL & r_TipoHabUso & "','"
                SQL = SQL & r_TipoRegimen & "',"

                SQL = SQL & r_Adultos & ","
                SQL = SQL & r_Ninos & ","
                SQL = SQL & r_Junior & ","
                SQL = SQL & r_Cunas & ")"





                Me.m_DBMORASOL.EjecutaSql(SQL)
                '    Me.m_DBMORASOL.
                If Me.m_DBMORASOL.HayError Then
                    Me.ListBoxDebug2.Items.Add(SQL)
                    Me.ListBoxDebug2.Items.Add(Me.m_DBMORASOL.StrError)
                    Me.mHayErrores = True
                    MessageBox.Show("Error de  Base de Datos" & vbCrLf & Me.m_DBMORASOL.StrError, "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    ' Me.m_DBMORASOL.CancelaTransaccion()
                    Exit For

                Else
                    ' graba fichero plano

                    LineaGraba = "I"

                    LineaGraba += r_ReservaAnterior.PadRight(10, " ")


                    r_Ejercicio = DR(1).ToString
                    r_Reserva = DR(2).ToString



                    LineaGraba += r_Bono.PadRight(25, " ")
                    LineaGraba += "01"
                    LineaGraba += Format(CDate(r_FechaCreacion), "dd/MM/yyyy")

                    If Me.m_PosibleNombre.Length > 0 Then
                        LineaGraba += Me.m_PosibleApellido.PadRight(50, " ")
                    Else
                        LineaGraba += r_Apellido.PadRight(50, " ")
                    End If



                    If Me.m_PosibleNombre.Length > 0 Then
                        LineaGraba += Me.m_PosibleNombre.PadRight(50, " ")
                    Else
                        LineaGraba += "NOMBRE".PadRight(50, " ")
                    End If



                    LineaGraba += Format(CDate(r_Llegada), "dd/MM/yyyy")

                    If r_Hllegada.Length > 0 Then
                        Me.m_AuxStr = Format(CDate(r_Llegada), "dd/MM/yyyy") & " " & Format(CDate(r_Hllegada), "HH:mm")
                        LineaGraba += Me.m_AuxStr
                    Else
                        LineaGraba += "                "
                    End If

                    LineaGraba += Format(CDate(r_Salida), "dd/MM/yyyy")

                    If r_Hsalida.Length > 0 Then
                        Me.m_AuxStr = Format(CDate(r_Salida), "dd/MM/yyyy") & " " & Format(CDate(r_Hsalida), "HH:mm")
                        LineaGraba += Me.m_AuxStr
                    Else
                        LineaGraba += "                "
                    End If



                    If r_Nacionalidad.Length > 0 Then
                        LineaGraba += r_Nacionalidad.PadRight(5, " ")
                    Else
                        LineaGraba += "".PadRight(5, " ")
                    End If

                    ' pax ( junuior = niños que pagan , niños = niños gratis



                    LineaGraba += r_Adultos.PadRight(2, " ")
                    LineaGraba += r_Junior.PadRight(2, " ")

                    ' bebes = niños + cunas ??? REVISAR
                    LineaGraba += CStr(CInt(r_Ninos) + CInt(r_Cunas)).PadRight(2, " ")



                    ' t hab 
                    LineaGraba += r_TipoHabFra.PadRight(10, " ")
                    LineaGraba += r_TipoHabUso.PadRight(10, " ")

                    ' Tipo de regomen
                    LineaGraba += r_TipoRegimen.PadRight(5, " ")
                    ' TTO
                    '   LineaGraba += r_TtoName.PadRight(50, " ")
                    LineaGraba += Mid(r_TtoName, 1, 10).PadRight(50, " ")
                    LineaGraba += r_Agencia.PadRight(50, " ")

                    ' obse
                    LineaGraba += r_ObservacionContrato.PadRight(400, " ")
                    Me.Filegraba.WriteLine(LineaGraba)
                End If


                Me.TextBoxregistrosTratados.Text = Ind
                Me.TextBoxregistrosTratados.Update()


            Next

            DS.Dispose()
            MyConnection.Close()
            Filegraba.Close()




            If Me.mHayErrores = False Then
                If Me.m_DBMORASOL.HayError = False Then
                    If MessageBox.Show("Confirma TRANSACCIÓN", "Atención", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                        Me.m_DBMORASOL.ConfirmaTransaccion()
                    Else
                        Me.m_DBMORASOL.CancelaTransaccion()
                    End If
                End If
            Else
                Me.m_DBMORASOL.CancelaTransaccion()
            End If


            If IsNothing(Me.ExcelApp) = False Then
                If Me.ExcelApp.Workbooks.Count > 0 Then
                    Me.ExcelApp.Workbooks.Close()
                End If
                Me.ExcelApp = Nothing
            End If

        Catch ex As Exception

            DS.Dispose()
            MyConnection.Close()
            Filegraba.Close()

            Me.m_DBMORASOL.CancelaTransaccion()
            MsgBox(ex.Message)
        Finally

        End Try
    End Sub
    Private Sub LeeExcelEntidades(ByVal vFileName As String, vHojaName As String)
        Try
            Dim DS As System.Data.DataSet
            Dim DT As System.Data.DataTable
            Dim DR As System.Data.DataRow


            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim StrConexion As String

            Dim Ind As Integer



            Me.Linea = 1

            Me.m_Hotel = InputBox("Hotel a Tratar", "AP MORASOL, ATLANTICO,SUITES", "SUITES")

            '  StrConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vFileName & ";Extended Properties=""Excel 8.0;HDR=Yes"""
            StrConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vFileName & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
            MyConnection = New System.Data.OleDb.OleDbConnection(StrConexion)
            DS = New System.Data.DataSet()
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & vHojaName & "$] where ([Estado reserva] = 'RESERVA' or [Estado reserva] = 'ENTRADA') AND [HOTEL] = '" & Me.m_Hotel & "'", MyConnection)
            '    "select * from [MAYO 2010$]", MyConnection)

            MyCommand.Fill(DS)
            DT = DS.Tables(0)

            ' DT.Locale.NumberFormat.NumberDecimalSeparator = "."
            ' DT.Locale.NumberFormat.NumberGroupSeparator = ","

            Dim Linea As String
            Dim LineaGraba As String = ""

            Dim r_Hotel As String

            Dim r_Cliente As String
            Dim r_TtoName As String
            Dim r_Agencia As String



            Me.ListBoxDebug.Items.Clear()



            SQL = "DELETE TENTIDADES"
            Me.m_ResultInt = Me.m_DBMORASOL.EjecutaSqlCommit(SQL)
            Me.ListBoxDebug2.Items.Add("Registros TENTIDADES Borrados " & Me.m_ResultInt)





            Me.m_DBMORASOL.IniciaTransaccion()


            Me.mHayErrores = False

            For Each DR In DT.Rows


                Ind = Ind + 1


                r_Hotel = DR(0).ToString

                r_Cliente = DR(6).ToString
                r_TtoName = DR(7).ToString
                r_Agencia = DR(8).ToString




                Linea = Ind.ToString.PadRight(10, " ") & "|" & r_Hotel.PadRight(10, " ") & r_Cliente.PadRight(20, " ") & "|" &
                r_TtoName.ToString.PadRight(50, " ") & "|" & r_Agencia.ToString.PadRight(50, " ")

                Me.ListBoxDebug.Items.Add(Linea)
                ' 







                ' inserta oracle
                SQL = " INSERT INTO TENTIDADES ( "
                SQL = SQL & "   HOTEL,TTOO,AGENCIA,CLIENTE )"

                SQL = SQL & "VALUES ('"
                SQL = SQL & r_Hotel & "','"
                SQL = SQL & r_TtoName & "','"
                SQL = SQL & r_Agencia & "','"
                SQL = SQL & r_Cliente & "')"



                Me.m_DBMORASOL.EjecutaSql(SQL)
                If Me.m_DBMORASOL.HayError Then
                    Me.ListBoxDebug2.Items.Add(SQL)
                    Me.ListBoxDebug2.Items.Add(Me.m_DBMORASOL.StrError)
                    Me.mHayErrores = True
                    MessageBox.Show("Error de  Base de Datos" & vbCrLf & Me.m_DBMORASOL.StrError, "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                    Exit For

                End If


                Me.TextBoxregistrosTratados.Text = Ind
                Me.TextBoxregistrosTratados.Update()


            Next

            DS.Dispose()
            MyConnection.Close()
            '




            If Me.mHayErrores = False Then
                If Me.m_DBMORASOL.HayError = False Then
                    If MessageBox.Show("Confirma TRANSACCIÓN", "Atención", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                        Me.m_DBMORASOL.ConfirmaTransaccion()
                    Else
                        Me.m_DBMORASOL.CancelaTransaccion()
                    End If
                End If
            Else
                Me.m_DBMORASOL.CancelaTransaccion()
            End If


            If IsNothing(Me.ExcelApp) = False Then
                If Me.ExcelApp.Workbooks.Count > 0 Then
                    Me.ExcelApp.Workbooks.Close()
                End If
                Me.ExcelApp = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try
    End Sub
    Private Function ExtraeNombre(vTexto As String) As String
        Try

            Me.m_PosibleNombre = ""
            Me.m_PosibleApellido = ""
            For I As Integer = vTexto.Length To 1 Step -1

                If Mid(vTexto, I, 1) = " " Or Mid(vTexto, I, 1) = "/" Then
                    Me.m_PosibleNombre = Mid(vTexto, I + 1, vTexto.Length - I)
                    Me.m_PosibleApellido = Mid(vTexto, 1, I - 1).Replace("/", "")
                    Return Mid(vTexto, 1, I - 1) & "," & Mid(vTexto, I + 1, vTexto.Length - I)


                    Exit For
                End If
            Next
            Return vTexto
        Catch ex As Exception
            MsgBox(ex.Message)
            Return ""
        End Try
    End Function
    Private Function EsUnNUmero(ByVal valor As String) As Boolean
        Try
            Dim style As Globalization.NumberStyles
            Dim culture As Globalization.CultureInfo
            Dim number As Double


            style = Globalization.NumberStyles.AllowDecimalPoint Or Globalization.NumberStyles.AllowThousands _
            Or Globalization.NumberStyles.AllowLeadingSign Or Globalization.NumberStyles.AllowTrailingSign
            culture = Globalization.CultureInfo.CreateSpecificCulture("es-ES")
            If Double.TryParse(valor, style, culture, number) Then
                '  Console.WriteLine("Converted '{0}' to {1}.", valor, number)
                Return True
            Else
                'Console.WriteLine("Unable to convert '{0}'.", valor)
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function
    Private Function ConvierteNumero(ByVal valor As String) As Double
        Try
            Dim style As Globalization.NumberStyles
            Dim culture As Globalization.CultureInfo
            Dim number As Double


            style = Globalization.NumberStyles.AllowDecimalPoint Or Globalization.NumberStyles.AllowThousands _
              Or Globalization.NumberStyles.AllowLeadingSign Or Globalization.NumberStyles.AllowTrailingSign
            culture = Globalization.CultureInfo.CreateSpecificCulture("es-ES")
            If Double.TryParse(valor, style, culture, number) Then
                '  Console.WriteLine("Converted '{0}' to {1}.", valor, number)
                Return number
            Else
                'Console.WriteLine("Unable to convert '{0}'.", valor)
                Return 0
            End If
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Private Sub ButtonReservas_Click(sender As Object, e As EventArgs) Handles ButtonReservas.Click
        Try
            If Me.ListBoxHojas.Items.Count > 0 Then
                If Me.ListBoxHojas.SelectedItems.Count > 0 Then


                    SQL = "DELETE TRESE"
                    Me.m_ResultInt = Me.m_DBMORASOL.EjecutaSqlCommit(SQL)
                    Me.ListBoxDebug2.Items.Add("Registros TRESE Borrados " & Me.m_ResultInt)

                    Me.FileEstaOk = False
                    Me.CrearFichero()
                    If FileEstaOk Then
                        Me.Cursor = Cursors.WaitCursor
                        Me.LeeExcelReservas(Me.TextBoxRutaLibroExcel.Text, Me.ListBoxHojas.SelectedItem)
                    Else
                        MessageBox.Show("No hay acceso al Fichero Plano  Destino de Reservas", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If


                End If


            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            If IsNothing(Me.ExcelApp) = False Then
                If Me.ExcelApp.Workbooks.Count > 0 Then
                    Me.ExcelApp.Workbooks.Close()
                End If
                Me.ExcelApp = Nothing
            End If

            If IsNothing(Me.m_DBHOTEL) = False Then
                Me.m_DBHOTEL.CerrarConexion()
            End If

            If IsNothing(Me.m_DBMORASOL) = False Then
                Me.m_DBMORASOL.CerrarConexion()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.LeerparametrosAppConfig()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Form Load")
        End Try
    End Sub

    Private Sub LeerparametrosAppConfig()
        Try
            Try

                Dim appSettings = ConfigurationManager.AppSettings
                Dim result As String

                ' CADENAS DE CONEXION

                result = appSettings("StrConexionMorasol")
                If IsNothing(result) Then
                    Me.ListBoxDebug2.Items.Add("Cadena de Conexión (MORASOL) no localizada en app.config")
                Else
                    Me.ListBoxDebug2.Items.Add(result)
                    Me.m_DBMORASOL = New C_DATOS.C_DatosOledb(result, False)

                    If Me.m_DBMORASOL.EstadoConexion = ConnectionState.Open Then
                        Me.m_DBMORASOL.EjecutaSqlCommit("ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'")
                        Me.ListBoxDebug2.Items.Add("Open Morasol ok")


                    Else
                        MessageBox.Show("No hay acceso a la Base de Datos", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                End If


                result = appSettings("StrConexionHotel")
                If IsNothing(result) Then
                    Me.ListBoxDebug2.Items.Add("Cadena de Conexión (RMC) no localizada en app.config")
                Else
                    Me.ListBoxDebug2.Items.Add(result)
                    Me.m_DBHOTEL = New C_DATOS.C_DatosOledb(result, False)
                    If Me.m_DBHOTEL.EstadoConexion = ConnectionState.Open Then
                        Me.m_DBHOTEL.EjecutaSqlCommit("ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'")
                        Me.ListBoxDebug2.Items.Add("Open RMC ok")
                    Else
                        MessageBox.Show("No hay acceso a la Base de Datos", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                End If


            Catch ec As ConfigurationErrorsException

                MessageBox.Show("Error reading app settings")
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, "LeerparametrosAppConfig")
        End Try
    End Sub

    Private Sub CrearFichero()

        Try

            Dim appSettings = ConfigurationManager.AppSettings
            Dim result As String

            ' CADENAS DE CONEXION

            result = appSettings("FileReservas")
            If IsNothing(result) Then
                Me.ListBoxDebug2.Items.Add("Ficehro destino de Reservass NO localizao en App.config")
            Else
                Me.ListBoxDebug2.Items.Add("Destino de Reservas = " & result)
                'Filegraba = New StreamWriter(vFile, False, System.Text.Encoding.UTF8)
                Filegraba = New StreamWriter(result, False, System.Text.Encoding.ASCII)
                FileEstaOk = True
            End If


        Catch ex As Exception
            FileEstaOk = False
            MsgBox("No dispone de acceso al Fichero Plano Destino de Reservas" & vbCrLf & ex.Message, MsgBoxStyle.Information, "Atención")
        End Try
    End Sub

    Private Sub ButtonDudasReservas_Click(sender As Object, e As EventArgs) Handles ButtonDudasReservas.Click
        Try
            Dim Texto As String


            Texto = "Como determinar el Número der Pax " & vbCrLf
            Texto += " Adultos = Adultos en Nav" & vbCrLf
            Texto += " Niños (Que pagan) = Junior en Nav" & vbCrLf

            Texto += " Niños (Que NO pagan) = Niños  en Nav  , Como pax Free ?" & vbCrLf
            Texto += " Cunas = Cunas  en Nav" & vbCrLf


            MsgBox(Texto)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function xConvertToDateTime(value As String) As String
        Dim convertedDate As Date
        Try
            convertedDate = Convert.ToDateTime(value)
            Return CStr(convertedDate)

        Catch e As FormatException
            Return ""
        End Try
    End Function

    Private Sub ButtonEquivalenciasEntidad_Click(sender As Object, e As EventArgs) Handles ButtonEquivalenciasEntidad.Click
        Try
            Me.LeeExcelEntidades(Me.TextBoxRutaLibroExcel.Text, Me.ListBoxHojas.SelectedItem)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
