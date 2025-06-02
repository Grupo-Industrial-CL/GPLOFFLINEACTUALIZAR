Imports SAP.Middleware.Connector
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP



Imports GPLOFFLINEACTUALIZAR.DatosGPP


Public Class DatosSAPConexion
    Public Shared misDatosSAP As DatosSAPConexion
    'Private wsSAP As WS_ZWS_SAPX.ZWS_WMClient
    Private miConfiguracion As ECCDestinationConfig
    Private miDestino As RfcDestination

    Public Shared ReadOnly Property DatosSAP As DatosSAPConexion
        Get
            If misDatosSAP Is Nothing Then
                misDatosSAP = New DatosSAPConexion
            End If
            Return misDatosSAP
        End Get
    End Property

    Public Function DameStockMARD(ByVal CENTRO As String,
                                  ByVal ALMACEN As String,
                                  ByVal MATERIAL As String,
                                  ByVal GTIN As Integer,
                                  ByVal Tipo_material As String) As Integer
        Try
            Dim misRegistrosMard As IRfcTable
            Dim tablaMard As List(Of stockAltualSAP)
            Dim Reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(Reintentos), RfcDestination)
            Dim iTotalSotck As Integer = 0

            iTotalSotck = 0

            If Not IsNothing(DatosSAP) Then
                Dim funcion As IRfcFunction = DameFuncion("ZDATOSSTOCKMARD_PER")
                funcion.SetValue("CENTRO", CENTRO)
                funcion.SetValue("ALMACEN", ALMACEN)
                funcion.SetValue("MATERIAL", MATERIAL)

                If Tipo_material.Trim.Length > 0 Then
                    funcion.SetValue("TIPOMATERIAL", Tipo_material.Trim)
                End If

                funcion.SetValue("MANDANTE", MandanteSAP)
                funcion.Invoke(DatosSAP)
                misRegistrosMard = funcion.GetTable("DATOS")
                If Not IsNothing(misRegistrosMard) AndAlso misRegistrosMard.RowCount > 0 Then

                    'tablaMard = New List(Of stockAltualSAP)
                    For Each miDato As SAP.Middleware.Connector.IRfcStructure In misRegistrosMard
                        iTotalSotck += CInt(miDato.GetDecimal("KILOS_LD"))
                    Next
                End If
            End If

            Return iTotalSotck

        Catch ex As Exception
            DameStockMARD = 0
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function DameForeCastSAP(ByVal Centro As String, ByVal Mes As Integer, ByVal Anio As Integer) As List(Of BeanForeCastSap)
        Try
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            Dim funcion As IRfcFunction = Nothing
            Dim TablaForeCastSAP As IRfcTable = Nothing

            DameForeCastSAP = New List(Of BeanForeCastSap)

            If Not IsNothing(DatosSAP) Then
                funcion = DameFuncion("ZDAMEDATOS_PLANIFICACION")

                funcion.SetValue("I_CENTRO", Centro)
                funcion.SetValue("I_ANYO", Anio)
                funcion.SetValue("I_MES", Mes)
                funcion.SetValue("I_AGRUPAR_STOCK", "X")

                funcion.Invoke(DatosSAP)

                TablaForeCastSAP = funcion.GetTable("ET_DATOS_PLANIFICACION")
                If Not IsNothing(TablaForeCastSAP) Then


                    DameForeCastSAP = (From elemento In TablaForeCastSAP
                                       Select New BeanForeCastSap(
                                               CodMaterial:=CStr(NoNull(elemento.GetValue("MATNR"), "A")),
                                               NombreMaterial:=CStr(NoNull(elemento.GetValue("NOMBRE_MATERIAL"), "A")),
                                               ClaseNecesidad:=CStr(NoNull(elemento.GetValue("BEDAE"), "A")),
                                               Mes:=CInt(NoNull(elemento.GetValue("MES"), "D")),
                                               Anio:=CInt(NoNull(elemento.GetValue("ANYO"), "D")),
                                               CantidadPlan:=CInt(NoNull(elemento.GetValue("PLNMG"), "D")),
                                               Unidad:=CStr(NoNull(elemento.GetValue("UNIDAD"), "A")),
                                               StockActual:=CInt(NoNull(elemento.GetValue("STOCK_ACTUAL"), "D")),
                                               FechaRotura:=elemento.GetValue("FECHA_ROTURA").ToString(),
                                               FechaRoturaEntradas:=elemento.GetValue("FECHA_ROTURA_ENTRADAS").ToString(),
                                                Fecha:=elemento.GetValue("PDATU").ToString(),
                                                miStockBloquedo:=CInt(NoNull(elemento.GetValue("STOCK_BLOQUEADO"), "D")),
                                                miEstatus:=CStr(NoNull(elemento.GetValue("STATUS"), "A")),
                                                FechaCorta:=CStr(NoNull(elemento.GetValue("PDATU"), "A")),
                                               Necesidad:=CStr(NoNull(elemento.GetValue("BEDAE"), "A"))
                                               )).ToList()

                End If

                'DameForeCastSAP = DameForeCastSAP.Where(Function(w) w.Anio = Anio And w.Mes = Mes).ToList()
                'If DameForeCastSAP.Where(Function(w) w.CodMaterial = "70905252").ToList().Count > 0 Or DameForeCastSAP.Where(Function(w) w.CodMaterial = "70905383").ToList().Count > 0 Or
                '    DameForeCastSAP.Where(Function(w) w.CodMaterial = "70906505").ToList().Count > 0 Or DameForeCastSAP.Where(Function(w) w.CodMaterial = "70908028").ToList().Count > 0 Or
                '    DameForeCastSAP.Where(Function(w) w.CodMaterial = "70908029").ToList().Count > 0 Then
                '    Dim h = 9
                'End If

                If DameForeCastSAP.Where(Function(w) w.CodMaterial = "70905465").ToList().Count > 0 Then
                    Dim h = 9
                End If



                Dim valores = DameForeCastSAP.Where(Function(w) w.CodMaterial = "70905465").ToList()
                Dim Meses = DameForeCastSAP.Select(Function(s) s.Mes).Distinct().ToList()

            End If

        Catch ex As Exception
            DameForeCastSAP = New List(Of BeanForeCastSap)
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Sub ReiniciarConexion()
        Try
            miConfiguracion = Nothing
            miDestino = Nothing
            misDatosSAP = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Public Function ComprobarConexionSAP() As Boolean
        Try
            Dim reintentos As Integer = ConstantesGPP.ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            If Not IsNothing(DatosSAP) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Conecta con SAP si no lo ha hecho ya y si falla la conexión reconecta.
    ''' </summary>
    ''' <param name="reintentos"></param> 
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConectarConSAP(ByRef reintentos As Integer,
                                   Optional ByVal ConComprobacion As Boolean = False) As Object
        Try
            If Not ping() Then
                If ("172.25.205.12".Length > 0 Or
                    "MOV7000".Length > 0 Or
                    "crislay".Length > 0 Or
                    "CLP".Length > 0) Then
                    Me.obtenerDestino()
                    If ping() Then
                        Return Me.miDestino
                    Else
                        reintentos = reintentos - 1
                        Me.miDestino = Nothing
                        If reintentos > 0 Then
                            Return ConectarConSAP(reintentos)
                        Else
                            Return Nothing
                        End If
                    End If
                Else
                    ' EscribirEventLog("Error - no hay parámetros de conexión SAP", EventLogEntryType.Information, 999)
                    Throw New Exception("Error - no hay parámetros de conexión SAP")
                End If
            Else
                Return Me.miDestino
            End If
        Catch ex As Exception
            If ConComprobacion Then
                Return "ERROR - NO HAY BAPIS"
            End If
            reintentos = reintentos - 1
            Me.miDestino = Nothing
            If reintentos > 0 Then
                'Si da un error en la conexión se descuenta el número de reintentos y se llama de forma recursiva al método
                Return ConectarConSAP(reintentos)
            Else
                ConectarConSAP = Nothing
            End If
            If IsNothing(ConectarConSAP) Then
                'Throw New NegocioDatosExcepction("Error Conexión - espere unos minutos y vuelva a intentarlo: Reintentos: " & reintentos, ex)
            End If
        End Try
    End Function



    Public Class ECCDestinationConfig
        Implements IDestinationConfiguration

        Public Function ChangeEventsSupported() As Boolean Implements IDestinationConfiguration.ChangeEventsSupported
            Return False
        End Function

        Public Event ConfigurationChanged(destinationName As String, args As RfcConfigurationEventArgs) Implements IDestinationConfiguration.ConfigurationChanged

        Public Function GetParameters(destinationName As String) As RfcConfigParameters Implements IDestinationConfiguration.GetParameters
            Try
                Dim misParametros As New RfcConfigParameters
                If True Then
                    misParametros.Add(RfcConfigParameters.Name, "Nueva Configuracion...")
                    misParametros.Add(RfcConfigParameters.User, "MOV7000")
                    misParametros.Add(RfcConfigParameters.Password, "crislay")
                    misParametros.Add(RfcConfigParameters.Client, "020")
                    misParametros.Add(RfcConfigParameters.Language, "S")
                    misParametros.Add(RfcConfigParameters.AppServerHost, "172.25.205.12")
                    misParametros.Add(RfcConfigParameters.SystemID, "CLP")
                    misParametros.Add(RfcConfigParameters.SystemNumber, "00")
                    misParametros.Add(RfcConfigParameters.PeakConnectionsLimit, "100")

                    '<add key = "gpl.Servidor.sap" value="172.25.205.12"/>
                    '  <add key = "gpl.Instancia.sap" value="00"/>
                    '  <add key = "gpl.Sistema.sap" value="CLP"/>
                    '  <add key = "gpl.Usuario.sap" value="mov7000"/>
                    '  <add key = "gpl.Clave.sap" value="crislay "/>
                    '  <add key = "gpl.Cliente.sap" value="020"/>

                End If
                Return misParametros
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
    End Class

    Private Sub obtenerDestino()
        Try
            If IsNothing(Me.miConfiguracion) Then
                Me.miConfiguracion = New ECCDestinationConfig()
                RfcDestinationManager.RegisterDestinationConfiguration(Me.miConfiguracion)
            End If
            If IsNothing(Me.miDestino) Then
                Me.miDestino = RfcDestinationManager.GetDestination("Nueva Configuracion...")
            End If
        Catch ex As Exception
            '  EscribirEventLog(ex.Message, EventLogEntryType.Information)
            ' Throw New NegocioDatosExcepction(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub
    Public Function DameFuncion(ByVal nombre As String) As IRfcFunction
        Try
            Return Me.miDestino.Repository.CreateFunction(nombre)
        Catch ex As Exception
            Me.miDestino = Nothing
            Try
                DatosSAP.obtenerDestino()
                If ping() Then
                    Return Me.miDestino.Repository.CreateFunction(nombre)
                Else
                    ' Throw New NegocioDatosExcepction(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
                End If
            Catch ex1 As Exception
                'Throw New NegocioDatosExcepction(ex1.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
            End Try
        End Try
    End Function

    Private Function ping() As Boolean
        Try
            If IsNothing(Me.miDestino) Then Return False
            Me.miDestino.Ping()
            ping = True
        Catch ex As Exception
            ping = False
        End Try
    End Function

    Public Function Dame_CantidadNotificada_OrdenNew(ByVal Orden As String) As Integer
        Try
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            'Dim funcion As IRfcFunction = Nothing
            Dim tablaDatos As IRfcTable = Nothing

            Dame_CantidadNotificada_OrdenNew = 0

            If Not IsNothing(DatosSAP) Then
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("Z_DAME_CANTIDAD_NOTIF")
                funcion.SetValue("I_ORDEN", Orden.PadLeft(12, "0"))

                funcion.Invoke(DatosSAP)


                Dame_CantidadNotificada_OrdenNew = CInt(NoNull(funcion.GetValue("E_CANT_NOTIFICADA"), "D"))



            End If

        Catch ex As Exception
            Dame_CantidadNotificada_OrdenNew = 0
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function GuardarFechasPrevistasXOrden(ByVal Orden As String, ByVal FechaIni As Date, ByVal FechaFin As Date) As Integer
        Dim Resultado = 0
        Try
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            'Dim funcion As IRfcFunction = Nothing
            Dim tablaDatos As IRfcTable = Nothing



            If Not IsNothing(DatosSAP) Then
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZMODIFICA_FECHAS_ORDEN")
                funcion.SetValue("I_ORDEN", Orden.PadLeft(12, "0"))
                If FechaIni < Date.Now Then
                    FechaIni = Date.Now
                End If

                If FechaFin < Date.Now Then
                    FechaFin = Date.Now
                End If
                funcion.SetValue("I_FECHA_INI_EXTREMA", FechaIni)
                funcion.SetValue("I_FECHA_FIN_EXTREMA", FechaFin)
                'funcion.SetValue("I_FECHA_INI_EXTREMA", sDameFechaCorta(FechaIni))
                'funcion.SetValue("I_FECHA_FIN_EXTREMA", sDameFechaCorta(FechaFin))

                funcion.Invoke(DatosSAP)


                Resultado = CInt(NoNull(funcion.GetValue("E_RC"), "D"))
                If Resultado = 0 Then

                Else
                    Dim i = 0

                End If



            End If

            Return Resultado

        Catch ex As Exception
            Return Resultado
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function Dame_StockDeGranel(ByVal Material As String) As Integer
        Dim Stock = 0
        Try
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            'Dim funcion As IRfcFunction = Nothing
            Dim tablaDatos As IRfcTable = Nothing



            If Not IsNothing(DatosSAP) Then
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZPPPF0101")
                funcion.SetValue("I_MATNR", Material)
                funcion.SetValue("I_WERKS", "12")

                funcion.Invoke(DatosSAP)


                Stock = CInt(NoNull(funcion.GetValue("E_STOCK"), "D"))



            End If

            Return Stock

        Catch ex As Exception
            Return Stock
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function VerificarDisponibilidadXOrden(ByRef ListaOrdenesFabricacion As List(Of DatosOrdenesSAP)) As String
        Dim listaFalta = ""
        Try
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            'Dim funcion As IRfcFunction = Nothing
            Dim tablaDatos As IRfcTable = Nothing
            Dim tOrdenes As IRfcTable = Nothing
            Dim miTablaFaltaMateriales As IRfcTable
            Dim miTablaFaltaEstatus As IRfcTable

            If Not IsNothing(DatosSAP) Then
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("Z_VERIFICACION_MATERIAL")

                tOrdenes = funcion.GetTable("IT_ORDENES")

                For Each midatoOrden In ListaOrdenesFabricacion
                    tOrdenes.Append()
                    tOrdenes.SetValue("ORDEN", midatoOrden.NumOrden.PadLeft(12, "0"))
                Next


                funcion.Invoke(DatosSAP)


                miTablaFaltaMateriales = funcion.GetTable("T_FALTAS")


                If Not IsNothing(miTablaFaltaMateriales) AndAlso miTablaFaltaMateriales.RowCount > 0 Then
                    For Each miDato In miTablaFaltaMateriales
                        Dim orderNum = CStr(NoNull(miDato.GetValue("ORDER_NUMBER"), "A"))
                        Dim Mensaje = CStr(NoNull(miDato.GetValue("MESSAGE"), "A"))
                        If Mensaje.Trim() <> "" Then
                            Dim o = 0
                        End If
                        If ListaOrdenesFabricacion.Select(Function(s) s.NumOrden.PadLeft(12, "0")).ToList().Contains(orderNum) Then
                            If ListaOrdenesFabricacion.Where(Function(w) w.NumOrden.PadLeft(12, "0") = orderNum).ToList().Count = 1 Then
                                ListaOrdenesFabricacion.Where(Function(w) w.NumOrden.PadLeft(12, "0") = orderNum).FirstOrDefault().Faltantes += CStr(NoNull(miDato.GetValue("MESSAGE"), "A")) & " | "
                            Else
                                For Each ordenLocal In ListaOrdenesFabricacion.Where(Function(w) w.NumOrden.PadLeft(12, "0") = orderNum).ToList()
                                    ordenLocal.Faltantes += CStr(NoNull(miDato.GetValue("MESSAGE"), "A")) & " | "
                                Next
                            End If
                        End If
                    Next
                End If

                miTablaFaltaEstatus = funcion.GetTable("ET_STATUS")


                If Not IsNothing(miTablaFaltaEstatus) AndAlso miTablaFaltaEstatus.RowCount > 0 Then
                    For Each miDato In miTablaFaltaEstatus
                        Dim orderNum = CStr(NoNull(miDato.GetValue("ORDEN"), "A"))
                        Dim Status = CStr(NoNull(miDato.GetValue("STATUS"), "A"))
                        If Status.Trim() <> "" Then
                            Dim o = 0
                        End If
                        If ListaOrdenesFabricacion.Select(Function(s) s.NumOrden.PadLeft(12, "0")).ToList().Contains(orderNum) Then
                            If ListaOrdenesFabricacion.Where(Function(w) w.NumOrden.PadLeft(12, "0") = orderNum).ToList().Count = 1 Then
                                ListaOrdenesFabricacion.Where(Function(w) w.NumOrden.PadLeft(12, "0") = orderNum).FirstOrDefault().EsLiberada += CStr(NoNull(miDato.GetValue("STATUS"), "A")) & " | "
                            Else
                                For Each ordenLocal In ListaOrdenesFabricacion.Where(Function(w) w.NumOrden.PadLeft(12, "0") = orderNum).ToList()
                                    ordenLocal.EsLiberada += CStr(NoNull(miDato.GetValue("STATUS"), "A")) & " | "
                                Next
                            End If
                        End If
                    Next
                End If
                'ListaOrdenesFabricacion.Where(Function(w) w.Faltantes.Trim.Length > 0).ToList().ForEach(Sub(s) s.Faltantes = s.Faltantes.Substring(0, s.Faltantes.Length - 3))
                'If Not IsNothing(miTablaFaltaMateriales) AndAlso miTablaFaltaMateriales.RowCount > 0 Then
                '    For Each miDato In miTablaFaltaMateriales
                '        If miDato Is miTablaFaltaMateriales.LastOrDefault() Then
                '            listaFalta += CStr(NoNull(miDato.GetValue("MESSAGE"), "A"))
                '        Else
                '            listaFalta += CStr(NoNull(miDato.GetValue("MESSAGE"), "A")) & " | "
                '        End If
                '    Next
                'End If

            End If

            Return listaFalta

        Catch ex As Exception
            DatosProduccion.FnlogApp("ERROR SAP - Verificación de  Disponibilidad : " & ex.Message)
            Return listaFalta
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function Dame_CantidadNotificada_Orden(ByVal Orden As String) As Integer
        Try
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            'Dim funcion As IRfcFunction = Nothing
            Dim tablaDatos As IRfcTable = Nothing

            Dame_CantidadNotificada_Orden = 0

            If Not IsNothing(DatosSAP) Then
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZDAMENOTIFORDEN")
                funcion.SetValue("ORDEN", Orden.PadLeft(12, "0"))

                funcion.Invoke(DatosSAP)

                tablaDatos = funcion.GetTable("NOTIFICACIONES")

                If IsNothing(tablaDatos) OrElse tablaDatos.RowCount = 0 Then
                    Exit Function
                End If

                For Each miDato In tablaDatos
                    If miDato.GetValue("MOVE_TYPE") = TipoMovimiento.Mov101 Then
                        Dame_CantidadNotificada_Orden += CInt(NoNull(miDato.GetValue("ENTRY_QNT"), "D"))
                    End If
                Next

                For Each miDato In tablaDatos
                    If miDato.GetValue("MOVE_TYPE") = TipoMovimiento.Mov102 Then
                        Dame_CantidadNotificada_Orden -= CInt(NoNull(miDato.GetValue("ENTRY_QNT"), "D"))
                    End If
                Next

            End If

        Catch ex As Exception
            Dame_CantidadNotificada_Orden = 0
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function Dame_HIGIENECOSMETICA_Orden(ByVal Orden As String) As String
        Dim Resultado = ""
        Try
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            'Dim funcion As IRfcFunction = Nothing
            Dim tablaDatos As IRfcTable = Nothing
            Dim tOrdenes As IRfcTable = Nothing

            Resultado = ""

            If Not IsNothing(DatosSAP) Then
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZPPPF0103")


                'funcion.SetValue("IT_AUFNR", Orden.PadLeft(12, "0"))
                tOrdenes = funcion.GetTable("IT_AUFNR")

                'For Each midatoOrden In ListaOrdenesFabricacion
                tOrdenes.Append()
                tOrdenes.SetValue("AUFNR", Orden.PadLeft(12, "0"))
                'Next

                funcion.Invoke(DatosSAP)

                tablaDatos = funcion.GetTable("ET_LINEA")

                If IsNothing(tablaDatos) OrElse tablaDatos.RowCount = 0 Then
                    Return Resultado
                    Exit Function
                End If

                For Each miDato In tablaDatos
                    'If miDato.GetValue("MOVE_TYPE") = TipoMovimiento.Mov101 Then                    
                    Resultado += CStr(NoNull(miDato.GetValue("LINEA"), "A"))
                    'End If
                Next



            End If
            Return Resultado

        Catch ex As Exception
            Return Resultado
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function DamePedidosVentaV2(Optional ByVal Centro As String = "",
                                 Optional ByVal FechaInicio As String = "",
                                 Optional ByVal FechaFin As String = "",
                                 Optional ByVal Material As String = "",
                                 Optional ByVal CodClienteSolic As Integer = 0,
                                 Optional ByVal CodClienteDest As Integer = 0,
                                 Optional ByVal Grupo As String = "",
                                 Optional Quitar_Concluidos As Boolean = False) As List(Of PedidosVenta)
        Try
            Dim dFecha As Date = Now
            Dim campoObservaciones As String = ""
            Dim reintentos As Integer = 60
            Dim DatosSAP As RfcDestination = CType(ConectarConSAP(reintentos), RfcDestination)
            Dim funcion As IRfcFunction = Nothing
            Dim tablaDatos As IRfcTable = Nothing

            DamePedidosVentaV2 = New List(Of PedidosVenta)

            If Not IsNothing(DatosSAP) Then
                funcion = DameFuncion("ZDAMEPEDIDOSVENTA_V2")
                'funcion.SetValue("MANDANTE", Datos.Sistema.ClienteSAP)
                funcion.SetValue("I_CENTRO", Centro)
                funcion.SetValue("I_ORGANIZACION_VENTAS", Centro)
                funcion.SetValue("I_IDIOMA", "S")
                If FechaInicio.Length > 0 Then
                    funcion.SetValue("I_FECHAINICIO", CStr(FechaInicio))
                    funcion.SetValue("I_FECHAFIN", CStr(FechaFin))
                End If

                If Material <> "0" AndAlso Material <> "" Then
                    funcion.SetValue("I_MATERIAL", Material)
                End If

                If CodClienteSolic > 0 Then
                    funcion.SetValue("I_CLIENTESOLIC", CodClienteSolic.ToString("0000000000"))
                End If

                funcion.Invoke(DatosSAP)

                tablaDatos = funcion.GetTable("ET_DATOS")

                If IsNothing(tablaDatos) OrElse tablaDatos.RowCount = 0 Then
                    Exit Function
                End If

                '.Kilos = CInt(IIf(NoNull(elemento.GetValue("KILOS_PTES"), "D") = 0, NoNull(elemento.GetValue("KILOS"), "D"), NoNull(elemento.GetValue("KILOS_PTES"), "D"))),

                DamePedidosVentaV2 = (From elemento In tablaDatos.Where(Function(P) IsNumeric(P.GetValue("MATERIAL")))
                                      Select New PedidosVenta With {.FechaPrevista = CDate(NoNull(elemento.GetValue("FECHAPREVENT"), "DT")),
                                                        .FechaReal = CStr(NoNull(elemento.GetValue("FECHAENTREAL"), "A")),
                                                        .CodClienteSolic = CInt(NoNull(elemento.GetValue("CLIENTESOLIC"), "D")),
                                                        .CodClienteDest = CInt(NoNull(elemento.GetValue("CLIENTEDEST"), "D")),
                                                        .Material = CInt(UTrim(elemento.GetValue("MATERIAL"))).ToString,
                                                        .Grupo = CStr(NoNull(elemento.GetValue("GRUPO"), "A")).Trim,
                                                        .Kilos = CInt(NoNull(elemento.GetValue("CANTIDAD_PEDIDO"), "D")),
                                                        .KilosPtes = CInt(NoNull(elemento.GetValue("KILOS_PTES"), "D")) - CInt(NoNull(elemento.GetValue("KILOS"), "D")),
                                                        .UnidadesPtes = CInt(NoNull(elemento.GetValue("CANTIDAD_PEDIDO"), "D")) - CInt(NoNull(elemento.GetValue("CANTIDAD_SERV_UMV"), "D")),
                                                        .KilosEnt = CInt(NoNull(elemento.GetValue("CANTIDAD_SERV_UMV"), "D")),
                                                        .Unidad = If(CStr(NoNull(elemento.GetValue("UNIDAD"), "A")).Trim = "ST", LiteralesSAP.Unidad, CStr(NoNull(elemento.GetValue("UNIDAD"), "A")).Trim),
                                                        .ClaseEntrega = CStr(NoNull(elemento.GetValue("CLASEENTREGA"), "A")),
                                                        .Pedido = CadenaSinCeros(CStr(NoNull(elemento.GetValue("NUMPEDIDO"), "D"))),
                                                        .TipoEnvio = CStr(NoNull(elemento.GetValue("ESCONSIGNA"), "A")),
                                                        .Centro = CStr(NoNull(elemento.GetValue("CENTRO"), "A")),
                                                        .Almacen = CStr(NoNull(elemento.GetValue("ALMACEN"), "A")),
                                                        .TipoPosicion = CStr(NoNull(elemento.GetValue("TIPOPOSICION"), "A")),
                                                        .LineaPedido = CadenaSinCeros(CStr(NoNull(elemento.GetValue("POSICION"), "A"))),
                                                        .OrdenTransporte = CStr(NoNull(elemento.GetValue("NUMTRANSPORTE"), "A")),
                                                        .EstadoOrdenTpte = CShort(NoNull(elemento.GetValue("ESTADOTPTE"), "D")).ToString,
                                                        .EntregaPendiente = .FechaReal = "0000-00-00",
                                                        .NombreMaterial = CStr(NoNull(elemento.GetValue("NOMBREMATERIAL"), "A")).Trim,
                                                        .NombreCliente = CStr(NoNull(elemento.GetValue("NOMBRECLIENTE"), "A")).Trim,
                                                        .StatusGLobal = enumToString(CStr(NoNull(elemento.GetValue("STATUS_GLOBAL"), "A")).Trim, GetType(Estatus_Pedido_Venta)),
                                                        .StatusEntrega = enumToString(CStr(NoNull(elemento.GetValue("STATUS_ENTREGA"), "A")).Trim, GetType(Estatus_Pedido_Venta))}).ToList
            End If


        Catch ex As Exception
            DamePedidosVentaV2 = New List(Of PedidosVenta)
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function



End Class
