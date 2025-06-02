Option Explicit On
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Text


Public Class Util

    '  Public Shared oBapiControl As SAPBAPIControlLib.SAPBAPIControl
    Private Shared oConnection As Object
    Public Shared DesarrolloSAP As Integer = 0
    Public Shared ProduccionSAP As Integer = 1

    Public Structure Turnos
        Public Const Mañana As String = "Mañana"
        Public Const Tarde As String = "Tarde"
        Public Const Noche As String = "Noche"
    End Structure

    Public Class EstructuraEnum

        Friend Ubicacion As Object

        Public Property Codigo As String
        Public Property Nombre As String


        Public Overrides Function Equals(obj As Object) As Boolean
            Try
                If obj.GetType() <> GetType(EstructuraEnum) Then
                    Return False
                End If
                Return Me.Codigo = CType(obj, EstructuraEnum).Codigo
            Catch ex As Exception
                Equals = False
            End Try
        End Function

        Public Overrides Function GetHashCode() As Integer
            Try
                Dim hashItemName = If(Me.Nombre Is Nothing, 0, Me.Nombre.GetHashCode())
                Dim hashItemCode = Me.Codigo.GetHashCode()
                Return hashItemName Xor hashItemCode
            Catch ex As Exception
                GetHashCode = 0
            End Try
        End Function
    End Class

    'Public Property BapiControl() As SAPBAPIControlLib.SAPBAPIControl
    '    Get
    '        BapiControl = oBapiControl
    '    End Get
    '    Set(ByVal value As SAPBAPIControlLib.SAPBAPIControl)
    '        Util.oBapiControl = value
    '    End Set
    'End Property

    Shared Sub esperar(ByVal iTiempo As Integer)

        'Proc. que espera el logintimeout antes de cerrar la BD
        Dim Inicio, Fin As Integer

        Inicio = CInt(DateAndTime.Timer)
        Do
            Fin = CInt(DateAndTime.Timer)
            If Fin - Inicio >= iTiempo Then
                Exit Do
            End If
        Loop
    End Sub

    Shared Function ValidarCuentaCorreo(ByVal CuentaCorreo As String) As Boolean
        Try
            ValidarCuentaCorreo = False
            If CuentaCorreo.Trim.Length > 0 Then
                Dim re As Regex = New Regex("^[\w._%-]+@[\w.-]+\.[a-zA-Z]{2,4}$")
                Dim m As Match = re.Match(CuentaCorreo)
                ValidarCuentaCorreo = (m.Captures.Count <> 0)
            End If
        Catch ex As Exception
            ValidarCuentaCorreo = False
        End Try
    End Function

    ''' <summary>
    ''' Redondea un Numero Entero en funcion de la posicion del redondeo, por ejemplo: 9256 -> 10000, si la posicion es el 4º digito
    ''' </summary>
    ''' <param name="ValorARedondear">Valor que se pasa para redondear</param>
    ''' <param name="PosicionRedondeo">Posicion del digito que se va a redondear</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function RedondearAlAlza(ByVal ValorARedondear As Integer, ByVal PosicionRedondeo As Integer) As Integer
        Try
            Dim retorno As Integer = 0
            Dim min As Double = Math.Pow(10, PosicionRedondeo - 1)

            RedondearAlAlza = Convert.ToInt32(Math.Ceiling((Convert.ToDouble(ValorARedondear) / min)) * min)
        Catch ex As Exception
            RedondearAlAlza = 0
        End Try

    End Function

    ''' <summary>
    ''' Redondea un Numero Entero en funcion del redondeo, por ejemplo: 9256 -> 9000, si la posision es el 4º digito
    ''' </summary>
    ''' <param name="ValorARedondear">Valor que se pasa para redondear</param>
    ''' <param name="PosicionRedondeo">Posicion del digito que se va a redondear</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function RedondearALaBaja(ByVal ValorARedondear As Integer, ByVal PosicionRedondeo As Integer) As Integer
        Try
            Dim retorno As Integer = 0
            Dim min As Double = Math.Pow(10, PosicionRedondeo - 1)

            RedondearALaBaja = Convert.ToInt32(Math.Floor((Convert.ToDouble(ValorARedondear) / min)) * min)
        Catch ex As Exception
            RedondearALaBaja = 0
        End Try

    End Function

    Shared Function EsFectivo(ByVal Fecha As Date) As Boolean
        Try
            EsFectivo = False

            If Fecha.DayOfWeek = DayOfWeek.Sunday Then
                EsFectivo = True
            Else
                Select Case Fecha.Month
                    Case 1
                        'año nuevo y reyes
                        If Fecha.Day = 1 Or Fecha.Day = 6 Then
                            EsFectivo = True
                        End If
                    Case 4
                        'san marcos
                        If Fecha.Day = 25 Then
                            EsFectivo = True
                        End If

                        'faltaria semana santa

                    Case 5
                        If Fecha.Day = 1 Then
                            EsFectivo = True
                        End If

                    Case 8
                        'la feria
                        If Fecha.Day = 15 Or Fecha.Day = 16 Then
                            EsFectivo = True
                        End If

                    Case 9
                        'dia extremadura
                        If Fecha.Day = 8 Then
                            EsFectivo = True
                        End If

                    Case 10
                        If Fecha.Day = 12 Then
                            EsFectivo = True
                        End If
                    Case 11
                        If Fecha.Day = 1 Then
                            EsFectivo = True
                        End If

                    Case 12
                        If Fecha.Day = 6 Or Fecha.Day = 8 Or Fecha.Day = 24 Or Fecha.Day = 25 Or
                            Fecha.Day = 31 Then
                            EsFectivo = True
                        End If
                    Case Else
                        EsFectivo = False
                End Select

            End If

        Catch ex As Exception
            EsFectivo = False
        End Try
    End Function

    Public Enum Tipo
        Entero = 1
        Texto = 2
        DateTime = 3
        Float = 4
        Lng = 5
    End Enum

    ''' <summary>
    ''' Funcion que devuelve un valor inicializado por si el valor pasado es nulo
    ''' </summary>
    ''' <param name="ArVar">Valor del parametro</param>
    ''' <param name="arTipo">Tipo del parametro: A->Cadena;N->Integer,Long;D->Double;F->Date;DT->DateTime;FG ->FechaGlobal</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function NoNull(ByVal ArVar As Object, ByVal arTipo As String) As Object
        'Dim auxfecha As Date
        Try
            'para columnas vacias sin datos
            If ArVar.Equals(System.DBNull.Value) Then
                Select Case arTipo
                    Case "A"
                        NoNull = ""
                    Case "N"
                        NoNull = 0
                    Case "D"
                        NoNull = 0
                    Case "F"
                        NoNull = CDate("00:00:0000")
                    Case "DT"
                        NoNull = New DateTime(1, 1, 1)
                    Case "FG"
                        NoNull = #1/1/1900#
                    Case Else
                        NoNull = " "
                End Select
                Exit Function
            End If

            If Len(ArVar.ToString) > 0 Then
                Select Case arTipo
                    Case "A"
                        NoNull = ArVar
                    Case "N"
                        NoNull = Val(ArVar)
                    Case "D"
                        If IsNumeric(ArVar) Then
                            NoNull = CDec(ArVar)
                        Else
                            NoNull = 0
                        End If
                    Case "F"
                        If ArVar Is "00/00/0000" Then
                            NoNull = ""
                        Else
                            If InStr(ArVar.ToString, "/") > 0 Then
                                NoNull = ArVar
                            Else
                                NoNull = Format(ArVar, "00/00/0000")
                            End If
                        End If
                    Case Else
                        NoNull = ArVar
                End Select
            Else
                Select Case arTipo
                    Case "A"
                        NoNull = ""
                    Case "N"
                        NoNull = 0
                    Case "D"
                        NoNull = 0
                    Case "F"
                        NoNull = CDate("00:00:0000")
                    Case Else
                        NoNull = " "
                End Select
            End If
        Catch ex As Exception
            Select Case arTipo
                Case "A"
                    NoNull = ""
                Case "N"
                    NoNull = 0
                Case "D"
                    NoNull = 0
                Case "F"
                    NoNull = CDate("00:00:0000")
                Case Else
                    NoNull = " "
            End Select
        End Try
    End Function

    Shared Function CadenaSinCeros(ByVal sEntrada As String) As String
        Dim i As Integer, j As Integer, sIntermedio As String
        For j = 1 To Len(sEntrada)
            If Mid(sEntrada, j, 1) <> "0" Then
                sIntermedio = Mid(sEntrada, j)
                Exit For
            End If
        Next
        CadenaSinCeros = sIntermedio
    End Function
    Shared Function PuntoComa(ByVal Cadena As Object) As String
        Try
            PuntoComa = Replace(Cadena.ToString, ",", ".", 1, 1)
        Catch ex As Exception
            PuntoComa = Cadena.ToString
        End Try
    End Function

    Shared Function UTrim(ByVal sCadena As Object) As String
        Try
            If NoNull(sCadena, "A").ToString.Trim <> "" Then
                UTrim = Trim(UCase(NoNull(sCadena, "A").ToString))
            Else
                UTrim = NoNull(sCadena, "A").ToString.Trim
            End If

        Catch ex As Exception
            UTrim = NoNull(sCadena, "A").ToString
        End Try
    End Function

    Public Shared Function EliminarEspeciales(ByVal s As String,
                                                Optional ByVal Filtro As String = "{}[]!""#$%&/()=?¡'¿|*+¨´:.;,<>") As String
        Try
            Dim I As Integer
            For I = 1 To Len(Filtro)
                s = Replace(s, Mid(Filtro, I, 1), "")
            Next
            EliminarEspeciales = s
        Catch ex As Exception
            Return s
        End Try
    End Function


    ''' <summary>
    ''' Funcion que devuelve una variable de tipo date de una variable numerica o cadena
    ''' </summary>
    ''' <param name="Fecha">La fecha debe de venir formateada de la forma YYYYMMDD</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function FechaAformatoDate(ByVal Fecha As String) As Date
        Try
            FechaAformatoDate = CDate(DateTime.ParseExact(Fecha, "yyyyMMdd", Nothing).ToString("dd\/MM\/yyyy"))
        Catch ex As Exception
            FechaAformatoDate = Nothing
        End Try
    End Function

    Shared Function DameFechaLarga(ByVal Fecha As String) As String
        Try
            'sDameFechaLarga = Format(Format(sFecha, "####/##/##"), "dd/MM/yyyy")
            DameFechaLarga = Mid(Fecha, 1, 4) & "/" &
                              Mid(Fecha, 5, 2) & "/" &
                              Mid(Fecha, 7, 2)
        Catch ex As Exception
            DameFechaLarga = Fecha
        End Try

    End Function

    Shared Function sDameFechaCorta(ByVal dFecha As Date) As String
        Try
            'sDameFechaCorta = Format(dFecha, "yyyyMMdd")
            sDameFechaCorta = Format(dFecha.Year, "0000") & Format(dFecha.Month, "00") &
                             Format(dFecha.Day, "00")
        Catch ex As Exception
            sDameFechaCorta = ""
        End Try
    End Function

    Shared Function sDameDiaSemana(ByVal dFecha As Date) As String
        Try
            sDameDiaSemana = ""
            Select Case Weekday(dFecha)
                Case 1 'Domingo
                    sDameDiaSemana = "D"
                Case 2 'Lunes
                    sDameDiaSemana = "L"
                Case 3 'Martes
                    sDameDiaSemana = "M"
                Case 4 'Miércoles
                    sDameDiaSemana = "X"
                Case 5 'Jueves
                    sDameDiaSemana = "J"
                Case 6 'Viernes
                    sDameDiaSemana = "V"
                Case 7 'Sábado
                    sDameDiaSemana = "S"
            End Select

        Catch ex As Exception
            sDameDiaSemana = ""
        End Try

    End Function

    Shared Function sPasar_Segundos_a_Horas(ByVal Segundos As Long,
                                            Optional ByVal MostrarSegundos As Boolean = True) As String

        Try
            Dim iMinutos As Integer
            Dim iHoras As Integer
            Dim iSegundos As Integer
            Dim lSegundosHora As Integer = 3600

            iHoras = CInt(Segundos \ lSegundosHora)
            iMinutos = CInt((Segundos Mod lSegundosHora) \ 60)
            iSegundos = CInt((Segundos Mod lSegundosHora) Mod 60)

            sPasar_Segundos_a_Horas = iHoras & ":" &
                                      Format(iMinutos, "00")
            If MostrarSegundos = True Then
                sPasar_Segundos_a_Horas = sPasar_Segundos_a_Horas & ":" &
                                          Format(iSegundos, "00")
            End If

        Catch ex As Exception
            sPasar_Segundos_a_Horas = "00:00:00"
        End Try

    End Function

    Shared Function ConvertirListaObjectosADatatable(Of T)(ByVal list As List(Of T)) As DataTable
        Try
            Dim Tabla As New DataTable

            Dim Fila As DataRow = Nothing
            Dim _itemProperties() As PropertyInfo =
                 list.Item(0).GetType().GetProperties()
            '    
            ' Meta Data. 
            '
            _itemProperties = list.Item(0).GetType().GetProperties()
            For Each p As PropertyInfo In _itemProperties
                Tabla.Columns.Add(p.Name,
                          p.GetGetMethod.ReturnType())
            Next
            '
            ' Data
            '
            For Each item As T In list
                '
                ' Get the data from this item into a DataRow
                ' then add the DataRow to the DataTable.
                ' Eeach items property becomes a colunm.
                '
                _itemProperties = item.GetType().GetProperties()
                Fila = Tabla.NewRow()
                For Each p As PropertyInfo In _itemProperties
                    Fila(p.Name) = p.GetValue(item, Nothing)
                Next
                Tabla.Rows.Add(Fila)
            Next

            ConvertirListaObjectosADatatable = Tabla
        Catch ex As Exception
            ConvertirListaObjectosADatatable = New DataTable
        End Try
    End Function

    Shared Function sDameNombreMes(ByVal iMes As Integer) As String
        Try
            Select Case iMes
                Case 1 : sDameNombreMes = "Enero"
                Case 2 : sDameNombreMes = "Febrero"
                Case 3 : sDameNombreMes = "Marzo"
                Case 4 : sDameNombreMes = "Abril"
                Case 5 : sDameNombreMes = "Mayo"
                Case 6 : sDameNombreMes = "Junio"
                Case 7 : sDameNombreMes = "Julio"
                Case 8 : sDameNombreMes = "Agosto"
                Case 9 : sDameNombreMes = "Septiembre"
                Case 10 : sDameNombreMes = "Octubre"
                Case 11 : sDameNombreMes = "Noviembre"
                Case 12 : sDameNombreMes = "Diciembre"
                Case Else : sDameNombreMes = ""
            End Select

        Catch ex As Exception
            sDameNombreMes = ""
        End Try

    End Function

    Shared Function Meter_Blancos(ByVal sCad As String,
                                  ByVal iLong As Integer) As String
        Try
            Meter_Blancos = Trim(sCad) & Space(iLong - Len(Trim(sCad)))
        Catch ex As Exception
            Meter_Blancos = Space(iLong)
        End Try

    End Function

    Shared Function sConvertir_Caracter_A_Hexadecimal(ByVal sCadena As Object) As String
        Try
            Dim i As Integer

            sConvertir_Caracter_A_Hexadecimal = ""
            For i = 1 To Len(sCadena)
                sConvertir_Caracter_A_Hexadecimal = sConvertir_Caracter_A_Hexadecimal &
                                                   Format(Hex(Asc(Mid(sCadena.ToString, i, 1))), "00") & " "

            Next
        Catch ex As Exception
            sConvertir_Caracter_A_Hexadecimal = ""
        End Try
    End Function

    Shared Function sConvertir_Hexadecimal_A_Caracter(ByVal sCadena As Object) As String
        Try
            Dim i As Integer

            sConvertir_Hexadecimal_A_Caracter = ""
            For i = 1 To Len(sCadena)
                sConvertir_Hexadecimal_A_Caracter = sConvertir_Hexadecimal_A_Caracter &
                                                    Chr(Val(CChar("0x" + Mid(sCadena.ToString, i, 1))))

            Next
        Catch ex As Exception
            sConvertir_Hexadecimal_A_Caracter = ""
        End Try
    End Function

    Shared Function ComprobarNumero(ByVal caracter As Char,
                                    ByVal bTengoComa As Boolean) As Char
        Try
            Select Case caracter
                Case CChar("0") To CChar("9")
                    ComprobarNumero = caracter
                Case CChar(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                    If bTengoComa = False Then
                        ComprobarNumero = caracter
                    Else
                        ComprobarNumero = CChar("")
                    End If
                Case CChar(vbBack)
                    ComprobarNumero = caracter
                Case CChar("-")
                    ComprobarNumero = caracter
                Case Else
                    ComprobarNumero = CChar("")
            End Select
        Catch ex As Exception
            ComprobarNumero = CChar("-1")
        End Try
    End Function

    Shared Function ComprobarDecimal(ByVal Caracter As Char,
                                     ByVal Numero As String) As Char
        Try
            Select Case Caracter
                Case CChar("0") To CChar("9")
                    ComprobarDecimal = Caracter
                Case CChar(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                    If CBool(InStr(Numero, System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) Then
                        ComprobarDecimal = CChar("")
                    Else
                        ComprobarDecimal = Caracter
                    End If

                Case CChar(vbBack)
                    ComprobarDecimal = Caracter
                Case CChar("-")
                    ComprobarDecimal = Caracter
                Case Else
                    ComprobarDecimal = CChar("")
            End Select
        Catch ex As Exception
            ComprobarDecimal = CChar("-1")
        End Try
    End Function

    Shared Function ComprobarEntero(ByVal caracter As Char) As Char
        Try
            Select Case caracter
                Case CChar("0") To CChar("9")
                    ComprobarEntero = caracter
                Case CChar(vbBack)
                    ComprobarEntero = caracter
                Case CChar("-")
                    ComprobarEntero = caracter
                Case Else
                    ComprobarEntero = CChar("")
            End Select
        Catch ex As Exception
            ComprobarEntero = CChar("-1")
        End Try
    End Function

    Shared Function EjecutarArchivo(ByVal sRutaArchivo As String) As Boolean
        Try
            System.Diagnostics.Process.Start(Trim(sRutaArchivo))
            EjecutarArchivo = True
        Catch ex As Exception
            EjecutarArchivo = False
        End Try
    End Function

    'Shared Function bFormulario_Activo(ByVal lCodOperForm As Long) As Boolean
    '    Try
    '        Dim miForm As Form = Nothing

    '        bFormulario_Activo = False

    '        For Each miForm In System.Windows.Forms.Application.OpenForms
    '            If CBool(CLng(NoNull(miForm.Tag, "D")) = lCodOperForm And miForm.Tag IsNot Nothing) Then
    '                bFormulario_Activo = True
    '                miForm.Activate()

    '                Exit Function
    '            End If
    '        Next

    '    Catch ex As Exception
    '        bFormulario_Activo = False
    '    End Try
    'End Function

    'Shared Function bFormulario_Activo(ByVal sNombreForm As String) As Boolean
    '    Try
    '        Dim miForm As Form = Nothing

    '        bFormulario_Activo = False

    '        For Each miForm In System.Windows.Forms.Application.OpenForms
    '            If Trim(UCase(miForm.Name)) = Trim(UCase(sNombreForm)) Then
    '                bFormulario_Activo = True
    '                miForm.Activate()

    '                Exit Function
    '            End If
    '        Next

    '    Catch ex As Exception
    '        bFormulario_Activo = False
    '    End Try
    'End Function

    'Shared Function bFormulario_Activo(ByVal TabControl As MdiTabControl.TabControl, ByVal Formulario As Form) As Boolean
    '    Try
    '        Dim miForm As Form = Nothing
    '        Dim miPagina As MdiTabControl.TabPage

    '        bFormulario_Activo = False

    '        For Each miForm In System.Windows.Forms.Application.OpenForms
    '            If miForm.Name.Equals(Formulario.Name) Then
    '                bFormulario_Activo = True

    '                miPagina = miForm.Tag

    '                miForm.Activate()
    '                For Each miPagina In TabControl.TabPages
    '                    If Formulario.Name.Equals(miPagina.Form.name) Then
    '                        miPagina.Select()
    '                        Exit For
    '                    End If
    '                Next

    '                Exit Function
    '            End If
    '        Next

    '    Catch ex As Exception
    '        bFormulario_Activo = False
    '    End Try
    'End Function

    Shared Function bCrear_carpeta(ByVal sCarpeta As String) As Boolean
        Try
            If IO.Directory.Exists(sCarpeta) = False Then
                My.Computer.FileSystem.CreateDirectory(sCarpeta)
            End If

            bCrear_carpeta = True

        Catch ex As Exception
            bCrear_carpeta = False
        End Try
    End Function

    Shared Function bCopiarFichero(ByVal sFichOrigen As String,
                                   ByVal sRuta_Destino As String) As Boolean
        Try
            System.IO.File.Copy(sFichOrigen, sRuta_Destino, True)
            bCopiarFichero = True
        Catch ex As Exception
            bCopiarFichero = False
        End Try
    End Function



    'Shared Function bLogOnSAP() As Boolean
    '    Try
    '        bLogOnSAP = False

    '        'oBapiControl = CreateObject("SAP.BAPI.1")
    '        oBapiControl = New SAPBAPIControlLib.SAPBAPIControl

    '        oConnection = oBapiControl.Connection
    '        oConnection.MessageServer = ""

    '        oConnection.ApplicationServer = "192.168.205.206"
    '        oConnection.System = "CLD"
    '        'oConnection.ApplicationServer = "1.0.0.25"
    '        'oConnection.System = "D36"
    '        oConnection.client = "200"
    '        oConnection.User = "MOV7000"
    '        oConnection.Password = "crislay"
    '        oConnection.Language = "ES"

    '        If Not oConnection.LogOn(0, True) Then
    '            oConnection = Nothing
    '            bLogOnSAP = False
    '            MessageBox.Show("Logon failed.", "Interface Stock 7000-SAP", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Else
    '            bLogOnSAP = True
    '        End If

    '        Exit Function
    '    Catch ex As Exception
    '        MessageBox.Show("ERROR: " & ex.Message & " -- bLogOnSAP", "Administrador", _
    '                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '    End Try

    'End Function

    ' ''' <summary>
    ' ''' Crea una conexion con el Sistema SAP
    ' ''' </summary>
    ' ''' <param name="TipoServidorSAP">Nos dice el tipo de servidor al que queremos conectarnos. 0->Desarrollo; 1->Produccion</param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Shared Function bLogOnSAP(ByVal TipoServidorSAP As Integer) As Boolean
    '    Try
    '        bLogOnSAP = False

    '        'oBapiControl = CreateObject("SAP.BAPI.1")
    '        oBapiControl = New SAPBAPIControlLib.SAPBAPIControl

    '        oConnection = oBapiControl.Connection
    '        oConnection.MessageServer = ""

    '        If TipoServidorSAP = DesarrolloSAP Then
    '            'oConnection.ApplicationServer = "1.0.0.57"
    '            oConnection.ApplicationServer = "192.168.205.204"
    '            oConnection.User = "RESINF"
    '            oConnection.Password = "champion"
    '        Else
    '            'oConnection.ApplicationServer = "1.0.0.25"
    '            oConnection.ApplicationServer = "192.168.205.206"
    '            oConnection.User = "MOV7000"
    '            oConnection.Password = "crislay"
    '        End If

    '        oConnection.System = "D36"
    '        oConnection.System = "CLD"
    '        oConnection.client = "200"
    '        oConnection.Language = "ES"

    '        If Not oConnection.LogOn(0, True) Then
    '            oConnection = Nothing
    '            bLogOnSAP = False
    '            MessageBox.Show("Logon failed.", "Interface Stock 7000-SAP", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Else
    '            bLogOnSAP = True
    '        End If

    '        Exit Function
    '    Catch ex As Exception
    '        MessageBox.Show("ERROR: " & ex.Message & " -- bLogOnSAP", "Administrador", _
    '                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '    End Try

    'End Function

    ' ''' <summary>
    ' ''' Crea una conexion con el Sistema SAP
    ' ''' </summary>
    ' ''' <param name="TipoServidorSAP">Nos dice el tipo de servidor al que queremos conectarnos. 0->Desarrollo; 1->Produccion</param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Shared Function LogOnSAP(ByVal DireccionServidor As String, _
    '                          ByVal Usuario As String, _
    '                          ByVal Clave As String, _
    '                          ByVal Sistema As String, _
    '                          ByVal Cliente As String, _
    '                          ByVal NumeroInstancia As String) As Boolean
    '    Try
    '        LogOnSAP = False

    '        oBapiControl = New SAPBAPIControlLib.SAPBAPIControl

    '        oConnection = oBapiControl.Connection
    '        oConnection.MessageServer = ""

    '        oConnection.ApplicationServer = DireccionServidor
    '        oConnection.User = Usuario
    '        oConnection.Password = Clave

    '        oConnection.System = Sistema
    '        oConnection.client = Cliente
    '        oConnection.SystemNumber = NumeroInstancia
    '        oConnection.Language = "ES"

    '        If Not oConnection.LogOn(0, True) Then
    '            oConnection = Nothing
    '            LogOnSAP = False
    '            MessageBox.Show("Logon failed.", "Interface Stock 7000-SAP", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Else
    '            LogOnSAP = True
    '        End If

    '        Exit Function
    '    Catch ex As Exception
    '        LogOnSAP = False
    '        MessageBox.Show("ERROR: " & ex.Message & " -- LogOnSAP", "Administrador", _
    '                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '    End Try
    'End Function

    Shared Function PrimerDiaMes(ByVal Mes As Integer, ByVal Año As Integer) As Date
        Try
            PrimerDiaMes = New Date(Año, Mes, 1)

        Catch ex As Exception
            PrimerDiaMes = New Date(1900, 1, 1)
        End Try
    End Function

    Shared Function UltimoDiaMes(ByVal Mes As Integer, ByVal Año As Integer) As Date
        Try
            UltimoDiaMes = New Date(Año, Mes, 1).AddMonths(1).AddDays(-1)
        Catch ex As Exception
            UltimoDiaMes = New Date(1900, 1, 1)
        End Try
    End Function

    Public Shared Sub EscribirEventLog(ByVal texto As String, ByVal tipo_entrada As EventLogEntryType)
        Try
            Dim Maquina As String = "."
            Dim Origen As String = My.Application.Info.AssemblyName
            'Escribimos en los Registros de Aplicación
            Dim Elog As EventLog
            Elog = New EventLog("Application", Maquina, Origen)
            Elog.WriteEntry(texto, tipo_entrada, 100, CType(50, Short))
            Elog.Close()
            Elog.Dispose()
        Catch ex As Exception
            'Aquí no podemos hacer nada tiene que tener permisos para hacerlo.
        End Try
    End Sub

    'Calcula la letra del NIF a partir de sus Números
    Private Shared Function CalculaNIF(ByVal strA As String) As String
        Try
            Const cCADENA As String = "TRWAGMYFPDXBNJZSQVHLCKE"
            Const cNUMEROS As String = "0123456789"
            Dim a, b, c, NIF As Integer
            Dim sb As New StringBuilder

            strA = Trim(strA)
            If Len(strA) = 0 Then Return ""

            ' Dejar sólo los números
            For i As Integer = 0 To strA.Length - 1
                If cNUMEROS.IndexOf(strA(i)) > -1 Then
                    sb.Append(strA(i))
                End If
            Next

            strA = sb.ToString
            a = 0
            NIF = CInt(Val(strA))
            Do
                b = CInt(Int(NIF / 24))
                c = NIF - (24 * b)
                a = a + c
                NIF = b
            Loop While b <> 0
            b = CInt(Int(a / 23))
            c = a - (23 * b)

            Return strA & Mid(cCADENA, CInt(c + 1), 1)
        Catch ex As Exception
            CalculaNIF = ""
        End Try
    End Function

    'Verifica si un CIF introducido es correcto, devuelve true si es correcto y false en caso contrario
    Public Shared Function Verificar_CIF(ByVal valor As String) As Boolean
        Try
            Dim strLetra As String, strNumero As String, strDigit As String
            Dim strDigitAux As String
            Dim auxNum As Integer
            Dim i As Integer
            Dim suma As Integer
            Dim letras As String

            letras = "ABCDEFGHKLMPQSXVJ"

            valor = UCase(valor)

            If Len(valor) < 9 OrElse Not IsNumeric(Mid(valor, 2, 7)) Then
                Return False
            End If

            strLetra = Mid(valor, 1, 1)     ' letra del CIF
            strNumero = Mid(valor, 2, 7)    ' Código de Control
            strDigit = Mid(valor, 9)        ' CIF menos primera y última posición

            If InStr(letras, strLetra) = 0 Then ' comprobamos la letra del CIF (1ª posición)
                Return False
            End If

            For i = 1 To 7
                If i Mod 2 = 0 Then
                    suma = suma + CInt(Mid(strNumero, i, 1))
                Else
                    auxNum = CInt(Mid(strNumero, i, 1)) * 2
                    suma = suma + (auxNum \ 10) + (auxNum Mod 10)
                End If
            Next
            suma = (10 - (suma Mod 10)) Mod 10

            Select Case strLetra
                Case "K", "P", "Q", "S"
                    suma = suma + 64
                    strDigitAux = Chr(suma)
                Case "X"
                    strDigitAux = Mid(CalculaNIF(strNumero), 8, 1)
                Case Else
                    strDigitAux = CStr(suma)
            End Select

            If strDigit = strDigitAux Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function Verificar_CIFIntracomunitario(ByVal valor As String,
                                                         ByVal acronimoPais As String,
                                                         ByVal patronCIF As String) As Boolean
        Try
            Select Case acronimoPais
                Case "ES"
                    Return Verificar_CIF(valor)
                Case Else
                    If patronCIF.Length > 0 Then
                        Return Regex.IsMatch(valor, patronCIF, RegexOptions.IgnoreCase)
                    Else
                        Return True
                    End If
            End Select
        Catch ex As Exception
            Verificar_CIFIntracomunitario = False
        End Try
    End Function

    'Verifica si un NIF introducido es correcto, devuelve true si es correcto y false en caso contrario
    Public Shared Function Verificar_NIF(ByVal valor As String) As Boolean
        Try
            Dim aux As String

            valor = valor.ToUpper ' ponemos la letra en mayúscula
            aux = valor.Substring(0, valor.Length - 1) ' quitamos la letra del NIF

            If aux.Length >= 7 AndAlso IsNumeric(aux) Then
                aux = CalculaNIF(aux) ' calculamos la letra del NIF para comparar con la que tenemos
            Else
                Return False
            End If

            If valor <> aux Then ' comparamos las letras
                Return False
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function DameNumeroDeLaSemana(Fecha As Date) As Integer
        Try
            DameNumeroDeLaSemana = MsgBox(DatePart(DateInterval.WeekOfYear, Fecha, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1))
        Catch ex As Exception
            DameNumeroDeLaSemana = 0
        End Try
    End Function

    Public Shared Function DameTurno(Fecha As DateTime) As String
        Try
            DameTurno = Turnos.Mañana

            If Fecha.TimeOfDay > New TimeSpan(6, 0, 0) And Fecha.TimeOfDay < New TimeSpan(14, 0, 0) Then
                DameTurno = Turnos.Mañana
            End If

            If Fecha.TimeOfDay > New TimeSpan(14, 0, 0) And Fecha.TimeOfDay <= New TimeSpan(22, 0, 0) Then
                DameTurno = Turnos.Tarde
            End If

            If (Fecha.TimeOfDay > New TimeSpan(0, 0, 0) And Fecha.TimeOfDay <= New TimeSpan(6, 0, 0)) Or
                Fecha.TimeOfDay > New TimeSpan(22, 0, 0) And Fecha.TimeOfDay <= New TimeSpan(23, 59, 59) Then
                DameTurno = Turnos.Noche
            End If
        Catch ex As Exception
            DameTurno = Turnos.Mañana
        End Try
    End Function

    Public Shared ReadOnly Property Desviacion(ValorInicial As Double, ValorFinal As Double) As String
        Get
            Dim iDesviacion As Integer = 0
            iDesviacion = CInt(ValorFinal - ValorInicial)
            If iDesviacion > 0 Then
                Desviacion = "+" & iDesviacion
            Else
                Desviacion = iDesviacion.ToString
            End If
        End Get
    End Property

    Public Shared ReadOnly Property PorcDesviacion(ValorInicial As Double, ValorFinal As Double) As Double
        Get
            If (ValorInicial = 0 And ValorFinal = 0) Or (ValorInicial = ValorFinal) Then
                PorcDesviacion = 0
            Else
                If ValorInicial = 0 Then
                    PorcDesviacion = 100
                Else
                    If ValorFinal = 0 Then
                        PorcDesviacion = 0
                    Else
                        PorcDesviacion = Math.Round((ValorFinal - ValorInicial) / ValorInicial, 4) * 100
                    End If
                End If
            End If
        End Get
    End Property


End Class
