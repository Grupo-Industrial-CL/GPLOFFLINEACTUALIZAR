
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP
Imports GPLOFFLINEACTUALIZAR.DatosProduccion
Imports System.Net
Imports System.IO
Imports Newtonsoft.Json
Imports SAP.Middleware.Connector

Module Module1

    Private miListaMaterialesPullSystem As List(Of Material)
    Private Cantidad As Integer = 0
    Sub Main()



        DatosProduccion.FnlogApp("........................INICIO TAREA.........................")

        DatosProduccion.FnlogApp("Actualizar_Cab_ListaMateriales")
        Actualizar_Cab_ListaMateriales()

        DatosProduccion.FnlogApp("Actualizar_ListaMateriales")
        Actualizar_ListaMateriales()

        DatosProduccion.FnlogApp("ActualizarMateriales")
        ActualizarMateriales()



        DatosProduccion.FnlogApp("Actualizar_ListaMaterialxMat")
        Actualizar_ListaMaterialxMat()

        DatosProduccion.FnlogApp("Actualizar_Clientes")
        Actualizar_Clientes()

        DatosProduccion.FnlogApp("........................FIN TAREA.........................")

        'Cantidad = 0
        'llamadaApiRest(OrdenEnvasado:="20067610", IdMaquina:="10000095")
        'Dim tempPost = New With {Key .status = "", Key .message = "", Key .data = 0}
        'Dim lst = JsonConvert.DeserializeAnonymousType(valor, tempPost)
        'Dim h = Cantidad

    End Sub

    Private Sub Actualizar_Clientes()
        Dim Actualizar_Clientes = False
        Try
            Actualizar_Clientes = False
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(DatosSAPConexion.DatosSAP.ConectarConSAP(reintentos), RfcDestination)
            Dim retorno As IRfcTable
            Dim sSql As String
            Dim DTDatos As New DataTable

            If Not IsNothing(DatosSAP) Then
                'Nos traemos la funcion 
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZDAMECLIENTESGPL")

                funcion.Invoke(DatosSAP)
                retorno = funcion.GetTable("DATOS")

                If IsNothing(retorno) OrElse retorno.RowCount = 0 Then
                    Exit Sub
                End If

                For Each elemento In retorno

                    sSql = "SELECT COUNT(*) AS CUANTOS " &
                            " FROM CLIENTES " &
                            " WHERE clCod = " & CInt(NoNull(elemento.GetValue("CODCLIENTE"), "D"))

                    If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                        If CInt(NoNull(DTDatos.Rows(0).Item("CUANTOS"), "D")) > 0 Then
                            sSql = " UPDATE CLIENTES " &
                                     " SET clNombre = '" & Replace(UTrim(elemento.GetValue("NOMCLIENTE")), "'", "") & "'," &
                                     " clPais = '" & UTrim(elemento.GetValue("PAIS")) & "'," &
                                     " clCP = '" & UTrim(elemento.GetValue("CP")) & "'," &
                                     " clPoblacion = '" & Replace(Replace(UTrim(elemento.GetValue("POBLACION")), "'", ""), "`", "") & "' " &
                                     " WHERE clCod = " & CInt(NoNull(elemento.GetValue("CODCLIENTE"), "D"))
                        Else
                            sSql = "INSERT INTO CLIENTES (clCod,clNombre,clPais,clCP,clPoblacion) VALUES (" &
                                   CInt(NoNull(elemento.GetValue("CODCLIENTE"), "D")) & ",'" &
                                   Replace(UTrim(elemento.GetValue("NOMCLIENTE")), "'", "") & "','" &
                                   CStr(UTrim(elemento.GetValue("PAIS"))) & "','" &
                                   CStr(UTrim(elemento.GetValue("CP"))) & "','" &
                                   CStr(EliminarEspeciales(UTrim(elemento.GetValue("POBLACION")))) & "')"

                        End If


                        Datos.CGPL.EjecutarConsulta(sSql)
                    End If

                Next

                Actualizar_Clientes = True

            Else
                Throw New Exception("ERROR - CONEXION CON SAP - REINICIAR APPLICACIÓN")
            End If

        Catch ex As Exception
            Actualizar_Clientes = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name &
                                                            "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Private Sub Actualizar_ListaMaterialxMat()
        Dim Actualizar_ListaMaterialxMat = False
        Try
            Actualizar_ListaMaterialxMat = False
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(DatosSAPConexion.DatosSAP.ConectarConSAP(reintentos), RfcDestination)
            Dim retorno As IRfcTable
            Dim i As Integer = 1
            Dim sSql As String
            Dim DTDatos As New DataTable


            If Not IsNothing(DatosSAP) Then
                'Nos traemos la funcion 
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZDAMELMMATGPL")

                funcion.SetValue("Centro", "0300")

                funcion.Invoke(DatosSAP)
                retorno = funcion.GetTable("DATOS")

                If IsNothing(retorno) OrElse retorno.RowCount = 0 Then
                    Exit Sub
                End If

                For Each elemento In retorno

                    sSql = " SELECT COUNT(*) AS CUANTOS " &
                           " FROM  ListaxMat" &
                           " WHERE lmMat = '" & UTrim(elemento.GetValue("MATERIAL")).TrimStart(CChar("0")) & "'" &
                           " AND lmLista = '" & CStr(UTrim(elemento.GetValue("LISTA"))) & "'"

                    If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                        If CInt(NoNull(DTDatos.Rows(0).Item("CUANTOS"), "D")) = 0 Then
                            sSql = "INSERT INTO ListaxMat (lmMat,lmLista) VALUES ('" &
                                     UTrim(elemento.GetValue("MATERIAL")).TrimStart(CChar("0")) & "','" &
                                     CStr(UTrim(elemento.GetValue("LISTA"))) & "')"
                            Datos.CGPL.EjecutarConsulta(sSql)
                        End If
                    End If

                Next

            Else
                Throw New Exception("ERROR - CONEXION CON SAP - REINICIAR APPLICACIÓN")
            End If

            Actualizar_ListaMaterialxMat = True

        Catch ex As Exception
            Actualizar_ListaMaterialxMat = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name &
                                                            "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Private Sub Actualizar_Cab_ListaMateriales()
        Dim Actualizar_Cab_ListaMateriales = False
        Try
            Actualizar_Cab_ListaMateriales = False
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(DatosSAPConexion.DatosSAP.ConectarConSAP(reintentos), RfcDestination)
            Dim retorno As IRfcTable
            Dim i As Integer = 1
            Dim sSql As String
            Dim DTDatos As New DataTable

            If Not IsNothing(DatosSAP) Then
                'Nos traemos la funcion 
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZDAMECABLMGPL")

                funcion.SetValue("tipoLM", "M")

                funcion.Invoke(DatosSAP)
                retorno = funcion.GetTable("DATOS")

                If IsNothing(retorno) OrElse retorno.RowCount = 0 Then
                    Exit Sub
                End If

                'sSql = "TRUNCATE TABLE ListaMaterialesCab"

                'If Datos.CGPL.EjecutarConsulta(sSql) = True Then
                For Each elemento In retorno
                    sSql = " SELECT COUNT(*) AS CUANTOS " &
                           " FROM ListaMaterialesCab " &
                           " WHERE clLista = '" & UTrim(elemento.GetValue("LISTA")) & "' "

                    If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                        If CInt(NoNull(DTDatos.Rows(0).Item("CUANTOS"), "D")) > 0 Then
                            sSql = " UPDATE ListaMaterialesCab " &
                                   " SET clCantidad = " & PuntoComa(CDbl(NoNull(elemento.GetValue("CANTIDADBASE"), "D"))) & "," &
                                   " clUM = '" & UTrim(elemento.GetValue("UNIDAD")) & "' " &
                                   " WHERE UPPER(RTRIM(clLista)) = '" & UTrim(elemento.GetValue("LISTA")) & "'"

                        Else
                            sSql = "INSERT INTO ListaMaterialesCab (clLista,clCantidad,clUM) VALUES ('" &
                               CStr(UTrim(elemento.GetValue("LISTA"))) & "'," &
                               PuntoComa(CDbl(NoNull(elemento.GetValue("CANTIDADBASE"), "D"))) & ",'" &
                               CStr(UTrim(elemento.GetValue("UNIDAD"))) & "')"
                        End If
                    End If
                    Datos.CGPL.EjecutarConsulta(sSql)
                Next
                'End If

                Actualizar_Cab_ListaMateriales = True

            Else
                Throw New Exception("ERROR - CONEXION CON SAP - REINICIAR APPLICACIÓN")
            End If

        Catch ex As Exception
            Actualizar_Cab_ListaMateriales = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name &
                                                            "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Private Sub ActualizarMateriales()
        Dim ActualizarMateriales = False
        Try
            ActualizarMateriales = False
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(DatosSAPConexion.DatosSAP.ConectarConSAP(reintentos), RfcDestination)
            Dim misDatos As IRfcTable
            Dim misTipos As IRfcTable
            Dim sSql As String = ""
            Dim DTDatos As New DataTable

            If Not IsNothing(DatosSAP) Then
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZDAMEMATERIALESGPL")

                misTipos = funcion.GetTable("TIPOMATERIALES")

                misTipos.Append()
                misTipos(0).SetValue("TIPO", "FERT")
                misTipos.Append()
                misTipos(1).SetValue("TIPO", "HALB")
                misTipos.Append()
                misTipos(2).SetValue("TIPO", "ROH")
                misTipos.Append()
                misTipos(3).SetValue("TIPO", "VERP")

                If misTipos.RowCount > 0 Then
                    funcion.SetValue("TIPOMATERIALES", misTipos)
                    funcion.SetValue("CENTRO", Centros_SAP.Plastiverd)

                    funcion.Invoke(DatosSAP)
                    misDatos = funcion.GetTable("DATOS")

                    If IsNothing(misDatos) = False AndAlso misDatos.RowCount > 0 Then
                        For Each miDato As SAP.Middleware.Connector.IRfcStructure In misDatos

                            sSql = " SELECT COUNT(*) AS CUANTOS " &
                                   " FROM Materiales " &
                                    " WHERE UPPER(RTRIM(maCod)) = '" & UTrim(miDato.GetValue("CODIGO")).TrimStart(CChar("0")) & "'"
                            If UTrim(miDato.GetValue("CODIGO")).TrimStart(CChar("0")) = "105925" Then
                                Dim h = 0
                            End If

                            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                                If CInt(NoNull(DTDatos.Rows(0).Item("CUANTOS"), "D")) > 0 Then
                                    sSql = " UPDATE Materiales " &
                                      " SET maTipoMat = '" & UTrim(miDato.GetValue("TIPO")) & "'," &
                                      " maGrupoArt = '" & UTrim(miDato.GetValue("GRUPO")) & "'," &
                                      " maUMBase = '" & UTrim(miDato.GetValue("UNIDAD")) & "'," &
                                      " maTipoEmbalaje = " & CShort(NoNull(miDato.GetValue("TIPOEMBALAJE"), "D")) & "," &
                                      " maNombre = '" & UTrim(miDato.GetValue("NOMBRE")) & "' " &
                                      " WHERE UPPER(RTRIM(maCod)) = '" & UTrim(miDato.GetValue("CODIGO")).TrimStart(CChar("0")) & "'"
                                Else
                                    sSql = "INSERT INTO Materiales (maCod,matipoMat,maGrupoArt,maUMBase,maTipoEmbalaje,maNombre) " &
                                 " VALUES ('" & UTrim(miDato.GetValue("CODIGO")).TrimStart(CChar("0")) & "','" &
                                                UTrim(miDato.GetValue("TIPO")) & "','" &
                                                UTrim(miDato.GetValue("GRUPO")) & "','" &
                                                UTrim(miDato.GetValue("UNIDAD")) & "'," &
                                                CShort(NoNull(miDato.GetValue("TIPOEMBALAJE"), "D")) & ",'" &
                                                UTrim(miDato.GetValue("NOMBRE")) & "')"
                                End If
                                Datos.CGPL.EjecutarConsulta(sSql)
                            End If
                        Next

                        ActualizarMateriales = True
                    End If
                End If
            End If
        Catch ex As Exception
            ActualizarMateriales = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Private Sub Actualizar_ListaMateriales()
        Dim Actualizar_ListaMateriales = False
        Try
            Actualizar_ListaMateriales = False
            Dim reintentos As Integer = ReintentosConexionSap
            Dim DatosSAP As RfcDestination = CType(DatosSAPConexion.DatosSAP.ConectarConSAP(reintentos), RfcDestination)
            Dim retorno As IRfcTable
            Dim i As Integer = 1
            Dim sSql As String
            Dim DTDatos As New DataTable

            If Not IsNothing(DatosSAP) Then
                'Nos traemos la funcion 
                Dim funcion As IRfcFunction = DatosSAPConexion.DatosSAP.DameFuncion("ZDAMELMGPL")

                funcion.SetValue("TIPOLM", "M")

                funcion.Invoke(DatosSAP)
                retorno = funcion.GetTable("DATOS")

                If IsNothing(retorno) OrElse retorno.RowCount = 0 Then
                    Exit Sub
                End If

                For Each elemento In retorno

                    Try
                        sSql = " SELECT COUNT(*) AS CUANTOS " &
                           " FROM ListaMateriales " &
                           " WHERE dlLista = '" & UTrim(elemento.GetValue("LISTA")) & "' " &
                           " AND dlNodo = " & CShort(NoNull(elemento.GetValue("NODO"), "D"))

                        If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                            If CInt(NoNull(DTDatos.Rows(0).Item("CUANTOS"), "D")) > 0 Then
                                sSql = "UPDATE LISTAMATERIALES " &
                                          " SET dlMaterial = '" & UTrim(elemento.GetValue("MATERIAL")) & "'," &
                                          " dlPosicion = '" & UTrim(elemento.GetValue("POSICION")) & "'," &
                                          " dlCantidad = " & PuntoComa(CDbl(NoNull(elemento.GetValue("CANTIDAD"), "D"))) & "," &
                                          " dlUM = '" & UTrim(elemento.GetValue("UNIDAD")) & "'," &
                                          " dlTipoPos = '" & UTrim(elemento.GetValue("TIPOPOSICION")) & "' " &
                                          " WHERE dlLista = '" & UTrim(elemento.GetValue("LISTA")) & "' " &
                                          " AND dlNodo = " & CShort(NoNull(elemento.GetValue("NODO"), "D"))
                            Else
                                sSql = "INSERT INTO ListaMateriales (dlLista,dlNodo,dlMaterial,dlPosicion,dlCantidad,dlUM,dlTipoPos) VALUES ('" &
                                        CStr(UTrim(elemento.GetValue("LISTA"))) & "'," &
                                        CDbl(NoNull(elemento.GetValue("NODO"), "D")) & ",'" &
                                        CStr(UTrim(elemento.GetValue("MATERIAL"))) & "','" &
                                        CStr(UTrim(elemento.GetValue("POSICION"))) & "'," &
                                        PuntoComa(CDbl(NoNull(elemento.GetValue("CANTIDAD"), "D"))) & ",'" &
                                        CStr(UTrim(elemento.GetValue("UNIDAD"))) & "','" &
                                        CStr(UTrim(elemento.GetValue("TIPOPOSICION"))) & "')"
                            End If
                        End If

                        Datos.CGPL.EjecutarConsulta(sSql)
                    Catch ex As Exception
                        Continue For
                    End Try

                Next

                Actualizar_ListaMateriales = True

            Else
                Throw New Exception("ERROR - CONEXION CON SAP - REINICIAR APPLICACIÓN")
            End If

        Catch ex As Exception
            Actualizar_ListaMateriales = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name &
                                                            "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Private Sub CargarFabricaciones()
        Try


            DatosProduccion.FnlogApp("........................INICIO Carga de Fabricaciones Mostradas en CONTROL DE PRODUCCION.........................")
            Dim FechaInicial As DateTime = DateTime.Now



            Dim ListaPuestosTrabajo = DamePuestosTrabajo(TipoPuestoTrabajo:="'" & ConstantesGPP.TipoPuestoTrabajo.Maquina & "'", Recursos:=False)
            Dim ListaOrdenesFabricacion = New List(Of DatosOrdenesSAP)
            Parallel.ForEach(Of PuestosTrabajo)(ListaPuestosTrabajo, Sub(Puesto As PuestosTrabajo)
                                                                         'DatosProduccion.FnlogApp("........................INICIO Carga de Fabricaciones para Puesto " & Puesto.Nombre & " .........................")
                                                                         CargarDatos(ListaOrdenesFabricacion, Puesto.CodigoPuestoTrabajo)
                                                                         'DatosProduccion.FnlogApp("........................FIN Carga de Fabricaciones para Puesto " & Puesto.Nombre & " .........................")

                                                                     End Sub)
            'For Each Puesto As PuestosTrabajo In ListaPuestosTrabajo
            '    DatosProduccion.FnlogApp("........................INICIO Carga de Fabricaciones para Puesto " & Puesto.Nombre & " .........................")
            '    CargarDatos(ListaOrdenesFabricacion, Puesto.CodigoPuestoTrabajo)
            '    DatosProduccion.FnlogApp("........................FIN Carga de Fabricaciones para Puesto " & Puesto.Nombre & " .........................")
            'Next

            DatosProduccion.FnlogApp("........................Verificamos Disponibilidad de Materiales de las : " & ListaOrdenesFabricacion.Count & " Ordenes .........................")
            ListaOrdenesFabricacion = ListaOrdenesFabricacion.OrderBy(Function(o) o.FechaInicioPrev).ToList()
            DatosSAPConexion.DatosSAP.VerificarDisponibilidadXOrden(ListaOrdenesFabricacion:=ListaOrdenesFabricacion)

            DatosProduccion.FnlogApp("........................Guardamos Disponibilidad de Materiales de las : " & ListaOrdenesFabricacion.Count & " Ordenes .........................")

            Parallel.ForEach(Of DatosOrdenesSAP)(ListaOrdenesFabricacion, Sub(miDato As DatosOrdenesSAP)

                                                                              If miDato.Tipo = "OrdenFabSAP" Then
                                                                                  Dim Resul = ModificarOrdenFabricacionSAP(CodigoFabricacion:=miDato.CodigoFabricacion.ToString(), FaltanteLista:=miDato.Faltantes, esLiberada:=miDato.EsLiberada, esNumorden:=miDato.NumOrden)
                                                                              End If
                                                                              If miDato.Tipo = "OrdenEnvSAP" Then
                                                                                  Dim Resul = ModificarOrdenEnvasadoSAP(CodigoFabricacion:=miDato.CodigoFabricacion.ToString(), FaltanteLista:=miDato.Faltantes, esLiberada:=miDato.EsLiberada, esNumorden:=miDato.NumOrden)
                                                                              End If
                                                                          End Sub)

            'For Each miDato As DatosOrdenesSAP In ListaOrdenesFabricacion
            '    If miDato.Tipo = "OrdenFabSAP" Then
            '        Dim Resul = ModificarOrdenFabricacionSAP(CodigoFabricacion:=miDato.CodigoFabricacion.ToString(), FaltanteLista:=miDato.Faltantes)
            '    End If
            '    If miDato.Tipo = "OrdenEnvSAP" Then
            '        Dim Resul = ModificarOrdenEnvasadoSAP(CodigoFabricacion:=miDato.CodigoFabricacion.ToString(), FaltanteLista:=miDato.Faltantes)
            '    End If
            'Next

            'Guardamos las Fechas de Inicio y Fin de la Verificación de Materiales  
            ModificarFechaInicioFinVerificacionSAP(FechaInicial, DateTime.Now)

            DatosProduccion.FnlogApp("........................Hemos guardado Correctamente las : " & ListaOrdenesFabricacion.Count & " Ordenes .........................")

            DatosProduccion.FnlogApp("........................FIN Carga de Fabricaciones Mostradas en CONTROL DE PRODUCCION.........................")

        Catch ex As Exception
            DatosProduccion.FnlogApp(ex.Message)
        End Try
    End Sub

    Public Sub CargarDatos(ByRef ListaOrdenesFabricacion As List(Of DatosOrdenesSAP), ByVal Optional puestoTrabajo As Integer = 0)
        Try
            'dgvControlProduccion.DataSource = Nothing
            'dgvControlProduccion.ForceInitialize()
            Dim listaFabricaciones = New List(Of Fabricaciones)


            Dim sFormatoAnterior As String = ""
            Dim bCambioFormato As Boolean
            Dim dFechaInicio As Date
            Dim ContadoRegistros As Integer = 0
            Dim CantidadEnvasadaRegistro As Integer = 0
            Dim dtFabricaciones As New DataTable("Fabricaciones")

            Dim primeraVezFutura = True


            Dim sDiaSemPrev As String = ""
            Dim sDiaSemPrevInicial As String = ""
            Dim dFechaFinAnt As Date = Now
            Dim iVelActMaq As Integer = 150
            Dim sTiempoFabricacion As String = ""
            Dim listaRecursos As String = String.Empty
            dFechaInicio = DateSerial(Now.Year, Now.Month, Now.Day)
            Dim MinutosFab As Integer = 0
            ' con que campo determino la fecha para el filtro 
            dtFabricaciones.Rows.Clear()

            listaFabricaciones = DameFabricaciones(OrdenarOrdenFabrcacion:=True,
                                                   PuestoTrabajo:=puestoTrabajo,
                                                   bPlanifFuturo:=True)

            '2023/07/18 Funcionalidad para ordenar las fabricaciones
            Dim ListaFabricacionesTEmporal = New List(Of Fabricaciones)
            ListaFabricacionesTEmporal.AddRange(listaFabricaciones.Where(Function(w) w.EnMarcha = 2).OrderBy(Function(o) o.FechaInicio).ThenBy(Function(o) o.OrdenMaq))
            ListaFabricacionesTEmporal.AddRange(listaFabricaciones.Where(Function(w) w.EnMarcha = 1).OrderBy(Function(o) o.OrdenMaq))
            ListaFabricacionesTEmporal.AddRange(listaFabricaciones.Where(Function(w) w.EnMarcha = 0).OrderBy(Function(o) o.OrdenMaq))
            ListaFabricacionesTEmporal.AddRange(listaFabricaciones.Where(Function(w) w.EnMarcha = 3).OrderBy(Function(o) o.FechaFuturoInicio))
            'ListaFabricacionesTEmporal.AddRange(listaFabricaciones.Where(Function(w) w.EnMarcha = 0 And w.OrdenEnvSAP = 0).OrderBy(Function(o) o.FechaPreFin))

            '2023-07-25 Solo las ordenes que tienen numero de Lote
            'If miPtConsulta IsNot Nothing Then
            '    listaFabricaciones = ListaFabricacionesTEmporal.Where(Function(w) w.NumeroLoteSAP.Trim <> "").ToList
            'Else
            listaFabricaciones = ListaFabricacionesTEmporal.ToList
            'End If

            '2024-08-09
            'Haremos reparto de la cantidad a Fabricar en Registros con el mismo OrdenEnvSAP
            Dim ListaFabricacionesParaRepartoCantidadFabricar = New List(Of Fabricaciones)
            Dim ListaOrdeneEnvSapARepartir As List(Of Integer) = New List(Of Integer)
            Dim ordenesTemp = listaFabricaciones.Where(Function(w) w.OrdenEnvSAP > 0).Select(Function(s) s.OrdenEnvSAP).Distinct().ToList()

            For Each ordenLoc As Integer In ordenesTemp
                If listaFabricaciones.Where(Function(w) w.OrdenEnvSAP = ordenLoc).ToList().Count > 1 Then
                    ListaOrdeneEnvSapARepartir.Add(ordenLoc)
                End If
            Next

            'DatosProduccion.FnlogApp("Para este Puesto de actualizaran " & listaFabricaciones.Count.ToString() & " Fabricaciones")

            For Each RegistroFabrica In listaFabricaciones
                If RegistroFabrica.Material.Codigo.Trim = "70901818" Then
                    Dim val = 5
                End If
                'DatosProduccion.FnlogApp("Actualizamos Fabricación con Material : " & RegistroFabrica.Material.Codigo.Trim)
                '2024-06-05 Obtenemos la cantidad que se ha fabricado
                '2024-06-05 Obtenemos la cantidad que se ha fabricado
                RegistroFabrica.CantidadFabricada = 0
                If RegistroFabrica.OrdenEnvSAP > 0 And puestoTrabajo > 0 Then
                    'If RegistroFabrica.OrdenEnvSAP = 20066650 Then
                    '    Dim h = 0
                    'End If
                    If RegistroFabrica.EnMarcha = ConstantesGPP.EstadoFabricacion.EnMarcha Then
                        Cantidad = 0
                        llamadaApiRest(OrdenEnvasado:=RegistroFabrica.OrdenEnvSAP.ToString(), IdMaquina:=puestoTrabajo.ToString())

                        If Cantidad > 0 Then
                            '2024-10-04 
                            'ver como podemos gestionar cuando las referencias están creadas por cajas, ya que las fotocélulas leen unidades
                            If RegistroFabrica.UnidadesPorCaja > 1 Then
                                Cantidad = CInt(Cantidad / RegistroFabrica.UnidadesPorCaja)
                            End If

                            RegistroFabrica.CantidadFabricada = Cantidad

                        Else
                            'LLamaremos a SAP para obtener la cantidad Notificada con la BAPI Z_DAME_CANTIDAD_NOTIF
                            Cantidad = DatosSAPConexion.DatosSAP.Dame_CantidadNotificada_OrdenNew(Orden:=RegistroFabrica.OrdenEnvSAP.ToString.Trim())

                            If Cantidad > 0 Then
                                '2024-10-04 
                                'ver como podemos gestionar cuando las referencias están creadas por cajas, ya que las fotocélulas leen unidades
                                If RegistroFabrica.UnidadesPorCaja > 1 Then
                                    Cantidad = CInt(Cantidad / RegistroFabrica.UnidadesPorCaja)
                                End If

                                RegistroFabrica.CantidadFabricada = Cantidad

                            End If

                        End If
                    End If


                    'If RegistroFabrica.CantidadFabricada > 0 Then
                    If ListaFabricacionesParaRepartoCantidadFabricar.Where(Function(w) w.OrdenEnvSAP = RegistroFabrica.OrdenEnvSAP).ToList().Count = 0 Then
                        Dim nuevoElemento As New Fabricaciones()
                        nuevoElemento.OrdenEnvSAP = RegistroFabrica.OrdenEnvSAP
                        nuevoElemento.CantidadFabricada = RegistroFabrica.CantidadFabricada
                        ListaFabricacionesParaRepartoCantidadFabricar.Add(nuevoElemento)
                    End If


                    If ListaOrdeneEnvSapARepartir.Contains(RegistroFabrica.OrdenEnvSAP) Then
                        Dim cantidadRestanteTotal = ListaFabricacionesParaRepartoCantidadFabricar.Where(Function(w) w.OrdenEnvSAP = RegistroFabrica.OrdenEnvSAP).FirstOrDefault().CantidadFabricada
                        If cantidadRestanteTotal > 0 Then
                            If RegistroFabrica.CantidadPlanificada >= cantidadRestanteTotal Then
                                RegistroFabrica.CantidadFabricada = cantidadRestanteTotal
                                cantidadRestanteTotal = 0
                            Else
                                RegistroFabrica.CantidadFabricada = RegistroFabrica.CantidadPlanificada
                                cantidadRestanteTotal = cantidadRestanteTotal - RegistroFabrica.CantidadPlanificada
                            End If
                        Else
                            RegistroFabrica.CantidadFabricada = 0
                        End If

                        ListaFabricacionesParaRepartoCantidadFabricar.Where(Function(w) w.OrdenEnvSAP = RegistroFabrica.OrdenEnvSAP).FirstOrDefault().CantidadFabricada = cantidadRestanteTotal
                    End If

                    'End If


                End If
                'Guardamos la Cantidad Fabricacada como se calcula en Control de Produccion
                'Para mostrarlo en OFFLINE en Control de Produccion 
                ModificarOrdenEnvasadoSAP_CantidadFabricadaGeneral(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(),
                                                                                   CantidadFabricadaSAP:=RegistroFabrica.CantidadFabricada)

                'Guardaremos la Cantidad Fabricada , pero solo la que proviene de SAP,
                'Esto con el fin de restar esta cantidad a lo Planificado y mostrarlo en los Modulos de
                'Fabricaciones contra Stock y Pedidos de Venta


                If RegistroFabrica.OrdenEnvSAP > 0 And puestoTrabajo > 0 Then
                    If RegistroFabrica.OrdenEnvSAP = 20070323 Then
                        Dim h = 0
                    End If
                    Dim CantidadFab = 0

                    'LLamaremos a SAP para obtener la cantidad Notificada con la BAPI Z_DAME_CANTIDAD_NOTIF
                    CantidadFab = DatosSAPConexion.DatosSAP.Dame_CantidadNotificada_OrdenNew(Orden:=RegistroFabrica.OrdenEnvSAP.ToString.Trim())
                    'If CantidadFab > 0 Then
                    Dim Resul = ModificarOrdenEnvasadoSAP_CantidadFabricada(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(), CantidadFabricadaSAP:=CantidadFab)
                    'End If

                End If


                With RegistroFabrica
                    If sFormatoAnterior <> .Formato Then
                        bCambioFormato = True
                        sFormatoAnterior = .Formato
                    Else
                        bCambioFormato = False
                    End If

                    If .EnMarcha = ConstantesGPP.EstadoFabricacion.Finalizada Then
                        dFechaFinAnt = .FechaFin
                        sTiempoFabricacion = Pasar_Segundos_a_Horas(Segundos:=DateDiff(DateInterval.Second, .FechaInicio, .FechaFin),
                                                                    bMostrarSegundos:=True)
                        sDiaSemPrev = ""
                        sDiaSemPrevInicial = ""

                    ElseIf .EnMarcha = ConstantesGPP.EstadoFabricacion.EnMarcha Then
                        '2024-07-31 Se requiere que la fecha de inicio sea a partir de la fecha Actual
                        dFechaFinAnt = DatosProduccion.DameHoraFin(SegundosFab:= .Minutos_Tiempo_FabricacionEnMarcha(bCambioFormato) * 60,
                                                                   FechaInicio:=CDate(Date.Now),
                                                                   CodPuestoTrabajo:= .CodigoPuestoTrabajo)

                        'dFechaFinAnt = DatosProduccion.DameHoraFin(SegundosFab:= .Minutos_Tiempo_Fabricacion(bCambioFormato) * 60,
                        '                                           FechaInicio:=CDate(IIf(.CantidadFabBuenas = 0,
                        '                                                                  dFechaFinAnt,
                        '                                                                  IIf(.FechaInicio = FechaGlobal,
                        '                                                                      .PuestoTrabajo.Proxima_FechaInicio_Turno,
                        '                                                                      .FechaInicio))),
                        '                                           CodPuestoTrabajo:= .CodigoPuestoTrabajo)
                        'dFechaFinAnt = .Fecha_Fin_Previsto_PedidoenMarcha
                        sTiempoFabricacion = Pasar_Segundos_a_Horas(Segundos:=RegistroFabrica.Minutos_Tiempo_FabricacionEnMarcha(bCambioFormato) * 60, bMostrarSegundos:=True)
                        'sTiempoFabricacion = Pasar_Segundos_a_Horas(Segundos:=RegistroFabrica.Minutos_Tiempo_Fabricacion(bCambioFormato) * 60, bMostrarSegundos:=True)
                        'MinutosFab = RegistroFabrica.Minutos_Para_Cambio(bCambioFormato)
                        'sTiempoFabricacion = Pasar_Segundos_a_Horas(Segundos:=RegistroFabrica.Minutos_Para_Cambio(bCambioFormato) * 60,
                        '                                            bMostrarSegundos:=True)
                        sDiaSemPrev = dFechaFinAnt.ToString

                        'Fecha Inicio nueva
                        Dim NuevaFechaInicio = CDate(Date.Now)

                        Dim FechaIniPrev = .FechaInicio.ToString() ' - (spanTiempoFabricacion - spanTiempoPreparacion)

                        sDiaSemPrevInicial = FechaIniPrev.ToString
                        'If .FechaInicio.Year = 1900 Then
                        '    Dim h = 0
                        'End If
                        'Guardamos las Fechas Previstas y Tiempo
                        ModificarFechasPrevYTiempo(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(),
                                                   FechaInicioPrev:=NuevaFechaInicio, FechaFinPrev:=dFechaFinAnt, Tiempo:=sTiempoFabricacion)

                        If RegistroFabrica.OrdenEnvSAP > 0 Then
                            DatosSAPConexion.DatosSAP.GuardarFechasPrevistasXOrden(Orden:=RegistroFabrica.OrdenEnvSAP.ToString.Trim(), NuevaFechaInicio, dFechaFinAnt)
                        End If

                        If RegistroFabrica.OrdenFabSAP > 0 Then
                            DatosSAPConexion.DatosSAP.GuardarFechasPrevistasXOrden(Orden:=RegistroFabrica.OrdenFabSAP.ToString.Trim(), NuevaFechaInicio.AddDays(-2), dFechaFinAnt.AddDays(-2))
                        End If
                        'RegistroFabrica.FechaInicio
                        If RegistroFabrica.OrdenFabSAP > 0 And RegistroFabrica.CantidadFabricada = 0 Then
                            If Not ListaOrdenesFabricacion.Select(Function(s) s.NumOrden.ToString().Trim()).ToList().Contains(RegistroFabrica.OrdenFabSAP.ToString().Trim()) Then
                                Dim miDato = New DatosOrdenesSAP()
                                miDato.NumOrden = RegistroFabrica.OrdenFabSAP.ToString()
                                miDato.Tipo = "OrdenFabSAP"
                                miDato.CodigoFabricacion = RegistroFabrica.CodigoFabricacion.ToString()
                                miDato.FechaInicioPrev = NuevaFechaInicio
                                ListaOrdenesFabricacion.Add(miDato)
                            End If

                            'Dim faltanteLista = DatosSAPConexion.DatosSAP.VerificarDisponibilidadXOrden(Orden:=RegistroFabrica.OrdenFabSAP.ToString.Trim())
                            'If faltanteLista.Trim() <> "" Then
                            '    Dim Resul = ModificarOrdenFabricacionSAP(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(), FaltanteLista:=faltanteLista)
                            'End If
                        End If

                        If RegistroFabrica.OrdenEnvSAP > 0 Then
                            If Not ListaOrdenesFabricacion.Select(Function(s) s.NumOrden.ToString().Trim()).ToList().Contains(RegistroFabrica.OrdenEnvSAP.ToString().Trim()) Then
                                Dim miDato = New DatosOrdenesSAP()
                                miDato.NumOrden = RegistroFabrica.OrdenEnvSAP.ToString()
                                miDato.Tipo = "OrdenEnvSAP"
                                miDato.CodigoFabricacion = RegistroFabrica.CodigoFabricacion.ToString()
                                miDato.FechaInicioPrev = NuevaFechaInicio
                                ListaOrdenesFabricacion.Add(miDato)
                            End If

                            'Dim faltanteLista = DatosSAPConexion.DatosSAP.VerificarDisponibilidadXOrden(Orden:=RegistroFabrica.OrdenEnvSAP.ToString.Trim())
                            ''Actualizamos el faltante X Orden 
                            'If faltanteLista.Trim() <> "" Then
                            '    Dim Resul = ModificarOrdenEnvasadoSAP(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(), FaltanteLista:=faltanteLista)
                            'End If
                        End If



                    ElseIf .EnMarcha = ConstantesGPP.EstadoFabricacion.PteFabricar Then


                        Dim TiempoPreparacion = 0
                        If bCambioFormato Then
                            TiempoPreparacion = .Minutos_Tiempo_Preparacion(True)
                        End If

                        Dim spanTiempoPreparacion As TimeSpan = New TimeSpan(0, TiempoPreparacion, 0)

                        'Dim spanTiempoFabricacion As TimeSpan = New TimeSpan(0, 0, RegistroFabrica.Minutos_Tiempo_Fabricacion(bCambioFormato) * 60)

                        Dim FechaIniPrev = dFechaFinAnt + spanTiempoPreparacion

                        sDiaSemPrevInicial = FechaIniPrev.ToString


                        dFechaFinAnt = DatosProduccion.DameHoraFin(SegundosFab:= .Minutos_Tiempo_Fabricacion(bCambioFormato) * 60,
                                                                   FechaInicio:=dFechaFinAnt,
                                                                   CodPuestoTrabajo:= .CodigoPuestoTrabajo)

                        sTiempoFabricacion = Pasar_Segundos_a_Horas(Segundos:=RegistroFabrica.Minutos_Tiempo_Fabricacion(bCambioFormato) * 60,
                                                                  bMostrarSegundos:=True)



                        sDiaSemPrev = dFechaFinAnt.ToString 'sDameDiaSemana(dFechaFinAnt) & " " & dFechaFinAnt.ToString("HH:mm")

                        'Guardamos las Fechas Previstas y Tiempo
                        ModificarFechasPrevYTiempo(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(),
                                                   FechaInicioPrev:=FechaIniPrev, FechaFinPrev:=dFechaFinAnt, Tiempo:=sTiempoFabricacion)

                        If RegistroFabrica.OrdenEnvSAP > 0 Then
                            If RegistroFabrica.OrdenEnvSAP = 20068678 Then
                                Dim g = 0
                            End If
                            DatosSAPConexion.DatosSAP.GuardarFechasPrevistasXOrden(Orden:=RegistroFabrica.OrdenEnvSAP.ToString.Trim(), FechaIniPrev, dFechaFinAnt)
                        End If

                        If RegistroFabrica.OrdenFabSAP > 0 Then
                            If RegistroFabrica.OrdenEnvSAP = 20068678 Then
                                Dim g = 0
                            End If
                            DatosSAPConexion.DatosSAP.GuardarFechasPrevistasXOrden(Orden:=RegistroFabrica.OrdenFabSAP.ToString.Trim(), FechaIniPrev.AddDays(-2), dFechaFinAnt.AddDays(-2))
                        End If


                        If RegistroFabrica.OrdenFabSAP > 0 And RegistroFabrica.CantidadFabricada = 0 Then
                            If Not ListaOrdenesFabricacion.Select(Function(s) s.NumOrden.ToString().Trim()).ToList().Contains(RegistroFabrica.OrdenFabSAP.ToString().Trim()) Then
                                Dim miDato = New DatosOrdenesSAP()
                                miDato.NumOrden = RegistroFabrica.OrdenFabSAP.ToString()
                                miDato.Tipo = "OrdenFabSAP"
                                miDato.CodigoFabricacion = RegistroFabrica.CodigoFabricacion.ToString()
                                miDato.FechaInicioPrev = FechaIniPrev
                                ListaOrdenesFabricacion.Add(miDato)
                            End If

                            'Dim faltanteLista = DatosSAPConexion.DatosSAP.VerificarDisponibilidadXOrden(Orden:=RegistroFabrica.OrdenFabSAP.ToString.Trim())
                            'If faltanteLista.Trim() <> "" Then
                            '    Dim Resul = ModificarOrdenFabricacionSAP(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(), FaltanteLista:=faltanteLista)
                            'End If
                        End If

                        If RegistroFabrica.OrdenEnvSAP > 0 Then
                            If Not ListaOrdenesFabricacion.Select(Function(s) s.NumOrden.ToString().Trim()).ToList().Contains(RegistroFabrica.OrdenEnvSAP.ToString().Trim()) Then
                                Dim miDato = New DatosOrdenesSAP()
                                miDato.NumOrden = RegistroFabrica.OrdenEnvSAP.ToString()
                                miDato.Tipo = "OrdenEnvSAP"
                                miDato.CodigoFabricacion = RegistroFabrica.CodigoFabricacion.ToString()
                                miDato.FechaInicioPrev = FechaIniPrev
                                ListaOrdenesFabricacion.Add(miDato)
                            End If

                            'Dim faltanteLista = DatosSAPConexion.DatosSAP.VerificarDisponibilidadXOrden(Orden:=RegistroFabrica.OrdenEnvSAP.ToString.Trim())
                            'If faltanteLista.Trim() <> "" Then
                            '    Dim Resul = ModificarOrdenEnvasadoSAP(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(), FaltanteLista:=faltanteLista)
                            'End If
                        End If



                    ElseIf .EnMarcha = ConstantesGPP.EstadoFabricacion.PlanFuturo Then

                        '2024-10-18 
                        'Algoritmo que Absorbe Fabricaciones Futuras con la ultima fecha de las fabricaciones Pendientes
                        Dim TiempoHolguraInicial = New TimeSpan(24, 0, 0)
                        Dim TiempoHolguraEntreFabricaciones = New TimeSpan(24, 0, 0)

                        Dim FechaEvaluar = ConstantesGPP.FechaGlobal

                        If primeraVezFutura Then
                            FechaEvaluar = dFechaFinAnt + TiempoHolguraInicial
                            primeraVezFutura = False
                        Else
                            FechaEvaluar = dFechaFinAnt + TiempoHolguraEntreFabricaciones

                        End If


                        If FechaEvaluar >= .FechaFuturoInicio Then
                            '
                            dFechaFinAnt = .FechaFuturoFin
                            .OrdenMaq = listaFabricaciones.Where(Function(w) w.EnMarcha = 0).Max(Function(o) o.OrdenMaq) + 1
                            'Convertimos la Fabricacion Futura en una Fabricacion Pendiente
                            .EnMarcha = CByte(EstadoFabricacion.PteFabricar)
                            .Modificar()

                        End If


                        ''Fecha Inicial Prevista
                        'sDiaSemPrevInicial = .FechaFuturoInicio.ToString()
                        ''Fecha Final Prevista                        
                        'sDiaSemPrev = .FechaFuturoFin.ToString()

                        'sTiempoFabricacion = Pasar_Segundos_a_Horas(Segundos:=DateDiff(DateInterval.Second, .FechaFuturoInicio, .FechaFuturoFin),
                        '                                          bMostrarSegundos:=True)

                    End If
                End With

                ' contatenamos los recursos que contenga este material

                For Each hojaRuta In RegistroFabrica.Material.HojasDeRuta
                    For Each recursoLista In hojaRuta.PuestosTrabajoHojaRuta
                        If recursoLista.Recurso = True Then
                            listaRecursos = recursoLista.Nombre.Trim + "   "
                        Else
                            listaRecursos = String.Empty
                        End If
                    Next
                Next

                'consultamos la cantidad fabricada en SAP si existeria
                If RegistroFabrica.OrdenFabSAP > 0 Then
                    'RegistroFabrica.CantidadFabSAP = DatosSAPConexion.DatosSAP.Dame_CantidadNotificada_Orden(Orden:=RegistroFabrica.OrdenFabSAP.ToString.Trim())

                    If RegistroFabrica.OrdenEnvSAP > 0 Then
                        CantidadEnvasadaRegistro = (From registros In listaFabricaciones
                                                    Where registros.OrdenFabSAP = RegistroFabrica.OrdenFabSAP And registros.OrdenEnvSAP = RegistroFabrica.OrdenEnvSAP
                                                    Select registros.CantidadFabBuenas).Sum
                    Else
                        CantidadEnvasadaRegistro = 0
                    End If
                Else
                    CantidadEnvasadaRegistro = 0
                End If

                '2024-08-07 Apartir de la cantidad Planificada en UNIDADES obtenemos la cantidad Planificada en KG
                Dim CantidadPlanificadaKg = DameCantidadPlanificadaEnKg(RegistroFabrica.Material, RegistroFabrica.CantidadPlanificada)

                '2024-10-02
                'Obtenemos el Stock de granel en control de producción
                Dim codMaterialFabricacion = ""
                Try
                    codMaterialFabricacion = If(RegistroFabrica.CodGranel.Trim() = "", If(RegistroFabrica.Material.ProductoFormula.Count > 0, RegistroFabrica.Material.ProductoFormula(0).Codigo.Trim, ""), RegistroFabrica.CodGranel.Trim())
                Catch ex As Exception
                    codMaterialFabricacion = ""
                End Try

                'LLamaremos a SAP para obtener la cantidad Notificada con la BAPI ZPPPF0101
                Dim StockGranelLocal = 0
                If codMaterialFabricacion.Trim() <> "" Then
                    StockGranelLocal = DatosSAPConexion.DatosSAP.Dame_StockDeGranel(Material:=codMaterialFabricacion.Trim())
                End If

                Dim MaterialPadreV = If(RegistroFabrica.CodGranel.Trim() = "", If(RegistroFabrica.Material.ProductoFormula.Count > 0, RegistroFabrica.Material.ProductoFormula(0).Codigo.Trim, ""), RegistroFabrica.CodGranel.Trim())

                '2024-11-22
                'Guardamos Variables de la Fabricaciones
                ModificarVariables(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(),
                                                   listaRecursos:=listaRecursos, CantidadEnvasadaRegistro:=CantidadEnvasadaRegistro,
                                                   CantidadPlanificadaKg:=CantidadPlanificadaKg, StockGranelLocal:=StockGranelLocal,
                                                   MaterialPadreV:=MaterialPadreV)


                'consultamos la cantidad fabricada en SAP si existeria
                If RegistroFabrica.OrdenFabSAP > 0 Then
                    RegistroFabrica.CantidadFabSAP = DatosSAPConexion.DatosSAP.Dame_CantidadNotificada_Orden(Orden:=RegistroFabrica.OrdenFabSAP.ToString.Trim())

                    Dim Resul = ModificarOrdenFabricacionSAP_CantidadFabricada(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(), CantidadFabricadaSAP:=RegistroFabrica.CantidadFabSAP)
                    'If RegistroFabrica.OrdenEnvSAP > 0 Then
                    '    CantidadEnvasadaRegistro = (From registros In listaFabricaciones
                    '                                Where registros.OrdenFabSAP = RegistroFabrica.OrdenFabSAP And registros.OrdenEnvSAP = RegistroFabrica.OrdenEnvSAP
                    '                                Select registros.CantidadFabBuenas).Sum
                    'Else
                    '    CantidadEnvasadaRegistro = 0
                    'End If
                Else
                    CantidadEnvasadaRegistro = 0
                End If

                'consultamos HIGIENECOSMETICA en SAP si existeria
                If RegistroFabrica.OrdenFabSAP > 0 Then
                    Dim HIGIENECOSMETICA = DatosSAPConexion.DatosSAP.Dame_HIGIENECOSMETICA_Orden(Orden:=RegistroFabrica.OrdenFabSAP.ToString.Trim())
                    If HIGIENECOSMETICA.Trim() <> "" Then
                        Dim Resul = ModificarOrdenFabricacionSAP_HIGIENECOSMETICA(CodigoFabricacion:=RegistroFabrica.CodigoFabricacion.ToString(), HIGIENECOSMETICA:=HIGIENECOSMETICA)
                    End If


                End If


                '   If RegistroFabrica.or Then

                ContadoRegistros = ContadoRegistros + 1


            Next





        Catch ex As Exception

        Finally



        End Try
    End Sub

    Private Function DameCantidadPlanificadaEnKg(MaterialDetalle As Material, cantidadPlanificadaUN As Integer) As Double
        Dim TotalFabKilogramos As Double = 0
        Try

            Dim CantidadKgX1000UN As Double = 0
            For Each itemListaMaterial As ListaMaterial In MaterialDetalle.CabLista_Material.MaterialesLista
                If itemListaMaterial.UnidadMedida.ToUpper().Trim() = "KG" And itemListaMaterial.Material.Tipo = "1205" Then
                    CantidadKgX1000UN = itemListaMaterial.Cantidad
                    Exit For
                End If

            Next

            If CantidadKgX1000UN = 0 Then
                Dim listaSemielaborados = MaterialDetalle.CabLista_Material.MaterialesLista
                For Each semi As ListaMaterial In listaSemielaborados
                    Dim MaterialDetalleSemi = New Material(semi.CodigoMaterial)
                    '(New System.Collections.Generic.Mscorlib_CollectionDebugView(Of NegocioGPP.ListaMaterial)(MaterialDetalle.CabLista_Material.MaterialesLista).Items(0)).Cantidad	3000	Double


                    For Each itemListaMaterial As ListaMaterial In MaterialDetalleSemi.CabLista_Material.MaterialesLista
                        If itemListaMaterial.UnidadMedida.ToUpper().Trim() = "KG" And itemListaMaterial.Material.Tipo = "1205" Then
                            CantidadKgX1000UN = itemListaMaterial.Cantidad
                            Exit For
                        End If
                    Next
                    If CantidadKgX1000UN > 0 Then
                        Exit For
                    End If
                Next


            End If

            If MaterialDetalle.CabLista_Material.MaterialesLista.Where(Function(w) w.Material.Tipo = "1205").Count > 1 Then

                CantidadKgX1000UN = MaterialDetalle.CabLista_Material.MaterialesLista.Where(Function(w) w.Material.Tipo = "1205").Sum(Function(s) s.Cantidad)

            End If

            Dim CantidadBase = MaterialDetalle.CabLista_Material.CantidadBase


            TotalFabKilogramos = Redondeo((cantidadPlanificadaUN * CantidadKgX1000UN) / CantidadBase, 0)

            Return TotalFabKilogramos

        Catch ex As Exception
            Return TotalFabKilogramos
        End Try
    End Function

    Private Async Sub llamadaApiRest(ByVal OrdenEnvasado As String, ByVal IdMaquina As String)
        Dim respuesta As String = Await GetHttp(OrdenEnvasado:=OrdenEnvasado, idMaquina:=IdMaquina)
        Dim tempPost = New With {Key .status = "", Key .message = "", Key .data = 0}
        'Dim lst As List(Of DatosOrdenes) = JsonConvert.DeserializeObject(Of List(Of DatosOrdenes))(respuesta)
        Dim lst = JsonConvert.DeserializeAnonymousType(respuesta, tempPost)
        Cantidad = lst.data
    End Sub

    Private Async Function GetHttp(ByVal OrdenEnvasado As String, ByVal idMaquina As String) As Task(Of String)
        Try


            Dim oRequest As WebRequest = WebRequest.Create("https://gmao.clgrupoindustrial.com/Iot/botesperseida?apikey=QXLkLqd7p8u-hTFrSM3LZ5iyss-DoUrLvdzK3lPljVQ&ordenenvasado=" + OrdenEnvasado + "&maquinagpp=" + idMaquina)
            Dim oResponse As WebResponse = oRequest.GetResponse()
            Dim sr As StreamReader = New StreamReader(oResponse.GetResponseStream())
            Return Await sr.ReadToEndAsync()

        Catch ex As Exception

        End Try
    End Function

    Private Sub CargarFabricacionesContraStock()
        Try
            DatosProduccion.FnlogApp("........................INICIO Fabricaciones Contra Stock..........................")

            Dim fechaInicial = New Date(Date.Now.Year, Date.Now.Month, 1) 'New Date(CType(Me.dlAnio.Value, Integer), CType(Me.dlMes.Value, Integer), 1)
            Dim miListaPullSystem = New List(Of PullSystem)
            Dim DiasControl = 15
            Dim dlMesFin = 3
            Dim txtDiasLaborales = 21
            Dim miListaPullSystemDetalle = New List(Of PullSystem)
            Dim agruparForeCast = True
            Dim incluirPedVentas = True
            Dim ListaPullsystem As New List(Of BeanPullSystem)
            Dim miListaPedidosVenta As List(Of PedidosVenta)

            Dim listaMeses As New List(Of Integer)({1, 2, 3, 4, 5, 6, 7})

            'Borramos lo del mes pasado y Versiones Nulas
            Dim FechaBorrar = fechaInicial.AddMonths(-1)
            Dim VersionBorrar = FechaBorrar.Year.ToString() & FechaBorrar.Month.ToString().PadLeft(2, "0"c)
            DatosProduccion.EliminarPullSystemPasados(Version:=VersionBorrar) 'YA


            For Each numMes As Integer In listaMeses

                miListaPullSystem = New List(Of PullSystem)
                ListaPullsystem = New List(Of BeanPullSystem)
                dlMesFin = numMes

                Dim Version = fechaInicial.Year.ToString() & fechaInicial.Month.ToString().PadLeft(2, "0"c) & dlMesFin.ToString().PadLeft(2, "0"c)

                For index = 0 To CInt(dlMesFin) - 1

                    miListaPullSystem.AddRange(DameRegistrosPullSystemSAP2(MesPS:=fechaInicial.AddMonths(index).Month,
                                                                       AnioPS:=fechaInicial.AddMonths(index).Year,
                                                                       iDiasControl:=DiasControl))

                Next

                DatosProduccion.FnlogApp("Obtenemos : " & miListaPullSystem.Count & " Registros  del Año : " & fechaInicial.Year.ToString() & " y Del mes : " & fechaInicial.Month & " Versión : " & Version)


                miListaPullSystemDetalle = miListaPullSystem

                miListaPullSystem = miListaPullSystem.Where(Function(w) w.Material.MostrarEnPedidosVenta = False).ToList()  ' filtro mostrar en pedidos de venta y no en Pullsystem

                miListaPullSystem = miListaPullSystem.Where(Function(w) w.Material.Tipo = "1206").ToList()

                miListaPullSystem = (From registro In miListaPullSystem
                                     Group registro By
                                                   registro.CodigoMaterial,
                                                   registro.Stock Into grupoRegistro = Group
                                     Select New PullSystem(
                                                       sCodigoMaterial:=CodigoMaterial,
                                                       iMes:=0,
                                                       iAño:=0,
                                                       iCantidad:=grupoRegistro.Sum(Function(p) p.Cantidad),
                                                       idiasControl:=CInt(DiasControl),
                                                       StockActual:=Stock,
                                                       FechaRotura:=grupoRegistro.Select(Function(s) s.FechaRoturaEntradas).FirstOrDefault(),
                                                       miStockBloqueado:=grupoRegistro.Select(Function(s) s.StockBloqueado).FirstOrDefault(),
                                                       miEstatus:=grupoRegistro.Select(Function(s) s.Estatus).FirstOrDefault())).ToList()





                For Each registroPullsystem In miListaPullSystem


                    If registroPullsystem.CodigoMaterial = "70905460" Then
                        'Debugger.Break()
                    End If


                    'Comentado temporalmente
                    'If infome.Rows.Count > 0 Then
                    '    registroPullsystem.fecha_Fin_Previsto = (From dato In infome
                    '                                             Where dato.CodMaterial.Trim() = registroPullsystem.CodigoMaterial.Trim()
                    '                                             Select dato.FechaFinPrev).DefaultIfEmpty(FechaGlobal).First
                    'End If

                    Try
                        ListaPullsystem.Add(New BeanPullSystem(diasControl:=CInt(DiasControl),
                                                        stockActual:=registroPullsystem.Stock,
                                                        fechaPrevistaFin:=registroPullsystem.fecha_Fin_Previsto,
                                                        cantidadFC:=registroPullsystem.Cantidad,
                                                        codMaterial:=registroPullsystem.CodigoMaterial.Trim,
                                                        MesPS:=registroPullsystem.Mes,
                                                        AnioPS:=registroPullsystem.Año,
                                                        DiasLaborables:=CInt(txtDiasLaborales),
                                                        FechaRotura:=registroPullsystem.FechaRotura,
                                                        StockBloqueado:=registroPullsystem.StockBloqueado,
                                                        Estatus:=registroPullsystem.Estatus,
                                                        DetallePullSystem:=(From registro In miListaPullSystemDetalle
                                                                            Where registro.CodigoMaterial = registroPullsystem.CodigoMaterial.Trim
                                                                            Select New BeanPullSystem(
                                                                                diasControl:=CInt(DiasControl),
                                                                                stockActual:=registro.Stock,
                                                                                fechaPrevistaFin:=registro.fecha_Fin_Previsto,
                                                                                cantidadFC:=registro.Cantidad,
                                                                                codMaterial:=registro.CodigoMaterial.Trim,
                                                                                MesPS:=registro.Mes,
                                                                                AnioPS:=registro.Año,
                                                                                DiasLaborables:=CInt(txtDiasLaborales),
                                                                                DetallePullSystem:=New List(Of BeanPullSystem), Necesidad:=registro.Necesidad, fechaCorta:=registro.FechaCorta
                                                                                                                 )).ToList()
                                                        ))
                    Catch ex As Exception
                        Dim g = ex.Message
                    End Try




                Next

                'Asignar Nombres Puesto de trabajo
                For Each Pedido In ListaPullsystem
                    Dim ListaNombres = New List(Of String)
                    If Pedido.Codigo = "7090778" Then
                        Dim g = 8
                        'Dim materialSeleccionado As New Material(Pedido.Codigo)
                        'ListaNombres = materialSeleccionado.HojasDeRuta.Select(Function(s) s.Nombre).Distinct().ToList()


                    End If

                    Dim materialSeleccionado As New Material(Pedido.Codigo)
                    For Each hoja In materialSeleccionado.HojasDeRuta
                        ListaNombres.AddRange(hoja.PuestosTrabajoHojaRuta.Select(Function(s) s.Nombre.Trim))
                    Next
                    ListaNombres = ListaNombres.Distinct().ToList()
                    'ListaNombres.AddRange(Pedido.NombresPuestoTrabajo.Select(Function(s) s.Nombre.Trim()))
                    If ListaNombres.Distinct().Count = 1 Then
                        Pedido.NombrePuestoTrabajo = ListaNombres(0)
                        Pedido.PuestoTrabajo = ListaNombres(0)
                    Else
                        Pedido.NombrePuestoTrabajo = Pedido.NombrePuestoTrabajo.Trim()
                        Pedido.PuestoTrabajo = Pedido.NombrePuestoTrabajo.Trim()
                    End If

                    Dim sCodigoMaterialFab As String = ""
                    If materialSeleccionado.ProductoFormula.Count > 0 Then
                        sCodigoMaterialFab = materialSeleccionado.ProductoFormula(0).Codigo.Trim
                    End If
                    If sCodigoMaterialFab = "" Then
                        For Each semi As ListaMaterial In materialSeleccionado.CabLista_Material.MaterialesLista
                            If semi.Material.Tipo = "1215" Then
                                Dim MaterialDetalleSemielaborado = New Material(semi.Material.Codigo)
                                sCodigoMaterialFab = MaterialDetalleSemielaborado.ProductoFormula(0).Codigo.Trim
                                'Dim MaterialDetalleSemielaboradoV2 = New Material(semi.CodigoMaterial)
                                'sCodigoMaterialFab = MaterialDetalleSemielaboradoV2.ProductoFormula(0).Codigo.Trim
                            End If
                        Next
                    End If
                    Pedido.CodigoMaterialFab = sCodigoMaterialFab
                    Pedido.NuevaFabricacion = CalculaNuevaFabricacionUN(Pedido)
                Next

#Region "Proceso Guardado Detalle"

                DatosProduccion.EliminarPullSystemTemporal() 'YA

                DatosProduccion.InsertarPullSystemTMP(miListaPullSystemDetalle, DiasControl:=DiasControl, AñoActualizacion:=fechaInicial.Year, MesActualizacion:=fechaInicial.Month, Version:=Version) 'NO

                DatosProduccion.EliminarPullSystem(Año:=fechaInicial.Year, Mes:=fechaInicial.Month, Version:=Version) 'YA                

                DatosProduccion.InsertarPullSystem() 'YA

                DatosProduccion.FnlogApp("Guardamos : " & miListaPullSystemDetalle.Count & " Registros en la tabla FabricacionesContraStockOFFLINE, Versión : " & Version)

#End Region

#Region "Proceso Guardado Agrupado"

                DatosProduccion.EliminarPullSystemTemporalAgrupado() 'YA

                DatosProduccion.InsertarPullSystemTMPAgrupado(ListaPullsystem, DiasControl:=DiasControl, AñoActualizacion:=fechaInicial.Year, MesActualizacion:=fechaInicial.Month, Version:=Version) 'NO

                DatosProduccion.EliminarPullSystemAgrupado(Año:=fechaInicial.Year, Mes:=fechaInicial.Month, Version:=Version) 'YA

                DatosProduccion.InsertarPullSystemAgrupado() 'YA

                DatosProduccion.FnlogApp("Guardamos : " & ListaPullsystem.Count & " Registros en la tabla PullSystemOFFLINEAgrupado, Versión : " & Version)

#End Region

            Next


            DatosProduccion.FnlogApp("........................FIN Fabricaciones Contra Stock..........................")

        Catch ex As Exception
            DatosProduccion.FnlogApp("ERROR Carga de Fabricaciones Contra Stock : " & ex.Message)
        End Try
    End Sub

    Private Function CalculaNuevaFabricacionUN(miPullSystem As BeanPullSystem) As Double
        Dim UNFabricacion As Double = 0
        Try
            Dim MaterialDetalle = New Material(miPullSystem.Codigo)


            Dim CantidadKgX1000UN As Double = 0
            For Each itemListaMaterial As ListaMaterial In MaterialDetalle.CabLista_Material.MaterialesLista
                If itemListaMaterial.UnidadMedida.ToUpper().Trim() = "KG" And itemListaMaterial.Material.Tipo = "1205" Then
                    CantidadKgX1000UN = itemListaMaterial.Cantidad
                    Exit For
                End If

            Next

            Dim ValorAMultimplicarNecesario = 1
            For Each semi As ListaMaterial In MaterialDetalle.CabLista_Material.MaterialesLista
                If semi.Material.Tipo = "1215" Then
                    ValorAMultimplicarNecesario = CInt(semi.Cantidad / 1000)
                End If
            Next


            'miPullSystem.KgNuevaFabricacion = ValorAMultimplicarNecesario * miPullSystem.KgNuevaFabricacion

            'Dim NuevaFabricacionKg As Integer = CType(miPullSystem.KgNuevaFabricacion, Integer)
            Dim NuevaFabricacionKg = ValorAMultimplicarNecesario * miPullSystem.KgNuevaFabricacion

            If NuevaFabricacionKg > 0 Then
                If CantidadKgX1000UN = 0 Then
                    Dim listaSemielaborados = MaterialDetalle.CabLista_Material.MaterialesLista
                    For Each semi As ListaMaterial In listaSemielaborados
                        Dim MaterialDetalleSemi = New Material(semi.CodigoMaterial)
                        '(New System.Collections.Generic.Mscorlib_CollectionDebugView(Of NegocioGPP.ListaMaterial)(MaterialDetalle.CabLista_Material.MaterialesLista).Items(0)).Cantidad	3000	Double


                        For Each itemListaMaterial As ListaMaterial In MaterialDetalleSemi.CabLista_Material.MaterialesLista
                            If itemListaMaterial.UnidadMedida.ToUpper().Trim() = "KG" And itemListaMaterial.Material.Tipo = "1205" Then
                                CantidadKgX1000UN = itemListaMaterial.Cantidad
                                Exit For
                            End If
                        Next
                        If CantidadKgX1000UN > 0 Then
                            Exit For
                        End If
                    Next


                End If

                If MaterialDetalle.CabLista_Material.MaterialesLista.Where(Function(w) w.Material.Tipo = "1205").Count > 1 Then

                    CantidadKgX1000UN = MaterialDetalle.CabLista_Material.MaterialesLista.Where(Function(w) w.Material.Tipo = "1205").Sum(Function(s) s.Cantidad)

                End If

                Dim CantidadBase = MaterialDetalle.CabLista_Material.CantidadBase

                '2024-07-04 
                '1206 Producto terminado
                'Viene en UN

                '1205 Granel
                'en KG
                CantidadKgX1000UN = CantidadKgX1000UN * ValorAMultimplicarNecesario
                Dim LoteMaximo As Double = 0
                Dim LoteMinimo As Double = 0 ' Material.LoteMinimo
                Dim valRedondeo As Double = 0 ' Material.Redondeo
                If MaterialDetalle.Tipo = "1206" Then
                    'LoteMinimo = Redondeo((MaterialDetalle.LoteMinimo * CantidadKgX1000UN) / CantidadBase, 0)
                    'Los Convertimos a KG
                    If MaterialDetalle.Redondeo = 0 Then
                        valRedondeo = MaterialDetalle.Redondeo
                    Else
                        valRedondeo = Redondeo((MaterialDetalle.Redondeo * CantidadKgX1000UN) / CantidadBase, 0)
                    End If
                    If MaterialDetalle.LoteMinimo = 0 Then
                        LoteMinimo = MaterialDetalle.LoteMinimo
                    Else
                        LoteMinimo = Redondeo((MaterialDetalle.LoteMinimo * CantidadKgX1000UN) / CantidadBase, 0)
                    End If
                    If MaterialDetalle.LoteMaximo = 0 Then
                        LoteMaximo = MaterialDetalle.LoteMaximo
                    Else
                        LoteMaximo = Redondeo((MaterialDetalle.LoteMaximo * CantidadKgX1000UN) / CantidadBase, 0)
                    End If

                Else
                    LoteMinimo = MaterialDetalle.LoteMinimo
                    valRedondeo = MaterialDetalle.Redondeo
                    LoteMaximo = MaterialDetalle.LoteMaximo
                End If

                'Dim EsRedondeo As Boolean = False
                Dim ValorRdoTanque = calculaValorRedondeoTanque(LoteMinimo, LoteMaximo, valRedondeo, NuevaFabricacionKg)
                'If ValorRdoTanque = 0 Then
                '    ValorRdoTanque = miPullSystem.ValorRdoTanque
                'End If
                Dim num = Redondeo((NuevaFabricacionKg / ValorRdoTanque), 0)
                Dim resultKg = ValorRdoTanque * num

                'Convertimos a Unidades la Nueva Fabricacion en KG
                UNFabricacion = (CType(resultKg, Double) * CantidadBase) / CantidadKgX1000UN
            End If





            Return UNFabricacion
        Catch ex As Exception
            DatosProduccion.FnlogApp("ERROR Carga de Fabricaciones Contra Stock : " & ex.Message)
            Return UNFabricacion

        End Try
    End Function

    Private Function calculaValorRedondeoTanque(ByVal LoteMinimo As Double, ByVal LoteMaximo As Double, ByVal Redondeo As Double, ByVal nuevaFabricacionKg As Integer) As Double
        Dim ValorRedondeoTanque As Double = 0
        Try

            'Dim LoteMinimo = Material.LoteMinimo
            'Dim Redondeo = Material.Redondeo

            If LoteMinimo > 0 Then
                If LoteMinimo >= nuevaFabricacionKg Then
                    ValorRedondeoTanque = LoteMinimo
                Else
                    If Redondeo > 0 Then
                        ValorRedondeoTanque = LoteMinimo
                        While ValorRedondeoTanque < nuevaFabricacionKg
                            ValorRedondeoTanque = ValorRedondeoTanque + Redondeo
                        End While
                        If LoteMaximo > 0 Then
                            If LoteMaximo <= ValorRedondeoTanque Then
                                ValorRedondeoTanque = LoteMaximo
                            End If
                        End If

                    Else
                        If nuevaFabricacionKg > LoteMinimo Then
                            ValorRedondeoTanque = nuevaFabricacionKg

                        End If
                    End If
                End If
            Else
                If LoteMaximo > 0 Then
                    If LoteMaximo >= nuevaFabricacionKg Then
                        ValorRedondeoTanque = LoteMaximo
                    End If
                Else
                    ValorRedondeoTanque = nuevaFabricacionKg
                End If

            End If

            Return ValorRedondeoTanque
        Catch ex As Exception
            DatosProduccion.FnlogApp("ERROR Carga de Fabricaciones Contra Stock : " & ex.Message)
            Return ValorRedondeoTanque

        End Try
    End Function

    Private Function Redondeo(ByVal Numero As Double, ByVal Decimales As Int32) As Double

        Redondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
        If Redondeo = 0 Then
            Redondeo = 1
        End If
        Return Redondeo
    End Function

    Private Sub CargarPedidosVenta()
        Try
            DatosProduccion.FnlogApp("GPP by CL")
            DatosProduccion.FnlogApp("Consulta Datos GPP......")
            DatosProduccion.FnlogApp("Creation Date: 2024-03-15..............................")
            DatosProduccion.FnlogApp("Last update Date: 2024-03-15..........................")
            DatosProduccion.FnlogApp("V1.0.0")

            DatosProduccion.FnlogApp("........................INICIO..........................")

            'bCargado = False
            Dim dFechaInicial As String = sDameFechaCorta(Date.Now.AddMonths(-2))
            Dim dFechaFinal As String = sDameFechaCorta(Date.Now.AddMonths(4))


            Dim miListaPedidosVenta As List(Of PedidosVenta)

            DatosProduccion.FnlogApp("Consultamos la BAPI ZDAMEPEDIDOSVENTA_V2")
            miListaPedidosVenta = DatosSAPConexion.DatosSAP.DamePedidosVentaV2(Centro:="12",
                                                                                 FechaInicio:=dFechaInicial,
                                                                                 FechaFin:=dFechaFinal)

            DatosProduccion.FnlogApp("Regresa : " & miListaPedidosVenta.Count & " Registros")
            Dim l = miListaPedidosVenta.Select(Function(s) s.Almacen).Distinct().ToList()
            miListaMaterialesPullSystem = DameMaterialesPullSystem()





            If miListaPedidosVenta.Count > 0 Then


                DatosProduccion.FnlogApp("Quitamos Pedidos concluidos y que aparecen en PullSystem")
                ' filtro mostrar en pedidos de venta 
                miListaPedidosVenta = (From datos In miListaPedidosVenta.AsParallel Where datos.MaterialDetalle.MostrarEnPedidosVenta = True
                                       Select datos).ToList


                ' filto para determinar si el material existe el pullsystem
                ' consultar los Materiales que existen el pullsystem y quitarlos de la lista de Pedidos de Venta.

                'If chkMatPullSystem.Checked = False Then
                'miListaPedidosVenta = (From datos In miListaPedidosVenta Where Not miListaMaterialesPullSystem.Exists(Function(p) p.Codigo.Trim = datos.Material.Trim)).ToList()
                'End If


                'If chkOcultarEntregados.Checked = True Then
                miListaPedidosVenta = (From datos In miListaPedidosVenta.AsParallel Where datos.StatusGLobal <> Estatus_Pedido_VentaDescripcion.Concluido
                                       Select datos).ToList
                'End If

                Dim t = miListaPedidosVenta.Where(Function(w) w.Material = "70904555").ToList()

                DatosProduccion.FnlogApp("Comienza el reparto del Stock")
                'Se repartira el stock de acuerdo al material comun y a la fecha de entrega mas proxima
                miListaPedidosVenta = RepartirStock(miListaPedidosVenta)

                DatosProduccion.EliminarPedidosTemporal()

                DatosProduccion.InsertarPedidosTMP(miListaPedidosVenta)

                DatosProduccion.EliminarPedidos(dFechaInicial, dFechaFinal)

                DatosProduccion.FnlogApp("Guardamos : " & miListaPedidosVenta.Count & " Registros en la tabla PedidosVentaOFFLINE")
                DatosProduccion.InsertarPedidos(miListaPedidosVenta)

                DatosProduccion.FnlogApp("........................FIN..........................")

                'miListaPedidosVentaGLOBAL = miListaPedidosVenta

                'Asignar Nombres Puesto de trabajo
                'For Each Pedido In miListaPedidosVenta
                '    'If Pedido.MaterialDetalle.Codigo = "70902803" Then
                '    '    Dim g = 8
                '    'End If


                '    Dim ListaNombres = New List(Of String)

                '    For Each hoja In Pedido.MaterialDetalle.HojasDeRuta
                '        ListaNombres.AddRange(hoja.PuestosTrabajoHojaRuta.Select(Function(s) s.Nombre.Trim))
                '    Next
                '    If ListaNombres.Distinct().Count = 1 Then
                '        Pedido.NombrePuestoTrabajo = Pedido.MaterialDetalle.HojasDeRuta(0).PuestosTrabajoHojaRuta(0).Nombre.Trim()
                '    Else
                '        Pedido.NombrePuestoTrabajo = Pedido.MaterialDetalle.HojasRutaDefecto.PuestosTrabajoHojaRuta(0).Nombre.Trim()
                '    End If

                '    Dim misTurnos As New List(Of Calendario)
                '    'If NombresPuestoTrabajo.Count = 1 Then
                '    '    misTurnos = DameTurnosMaquina(Pedido.FechaPrevista.AddDays(-(Pedido.MaterialDetalle.DiasFabPropia + Pedido.MaterialDetalle.DiasPP) * 5), NombresPuestoTrabajo(0).CodigoPuestoTrabajo)
                '    'Else
                '    Dim Cod = Pedido.MaterialDetalle.HojasRutaDefecto.PuestosTrabajoHojaRuta(0).CodigoPuestoTrabajo
                '    misTurnos = DameTurnosMaquina(Pedido.FechaPrevista.AddDays(-(Pedido.MaterialDetalle.DiasFabPropia + Pedido.MaterialDetalle.DiasPP) * 5), Cod)
                '    'End If

                '    Dim turnosARestar = (Pedido.MaterialDetalle.DiasFabPropia + Pedido.MaterialDetalle.DiasPP) * 3
                '    'Obtenemos en que turno esta la la Fecha Prevista
                '    Dim IdTurno = misTurnos.Where(Function(w) Pedido.FechaPrevista >= w.InicioTurno).ToList().Select(Function(s) s.Id).Max() - turnosARestar
                '    If IdTurno > 0 Then
                '        Pedido.FechaPlan = misTurnos.Where(Function(w) w.Id = IdTurno).FirstOrDefault().InicioTurno
                '    Else
                '        Pedido.FechaPlan = misTurnos.FirstOrDefault().InicioTurno
                '    End If


                'Next


                'dgvPedidosVentas.DataSource = miListaPedidosVenta

                'For index = 0 To gvPedidosVentas.RowCount
                '    Dim puesto = CType(gvPedidosVentas.GetRowCellValue(index, "NombrePuestoTrabajo"), String)
                '    gvPedidosVentas.SetRowCellValue(index, "PuestoTrabajo", puesto)
                'Next


            Else
                ' MDIPrincipal.closeSplashForm()
                'sMensajeUsuario = "No existen datos"
                'XtraMessageBox.Show(sMensajeUsuario, "ADMINISTRADOR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If


        Catch ex As Exception
            'bCargado = False
            DatosProduccion.FnlogApp("ERROR Carga de Pedidos de Venta : " & ex.Message)
        Finally
            'MDIPrincipal.closeSplashForm()
            'bCargado = True


        End Try
    End Sub

    Public Function DepurarMateriales(ListaPullsystemLocal As List(Of BeanPullSystem)) As List(Of BeanPullSystem)
        Dim ListaPullsystemDepurada As New List(Of BeanPullSystem)
        For Each miPullSystem In ListaPullsystemLocal
            Dim Material As New Material(miPullSystem.Codigo)
            If Material.Tipo.Trim = "1206" Then
                ListaPullsystemDepurada.Add(miPullSystem)
            End If
        Next
        Return ListaPullsystemDepurada
    End Function
    Private Function RepartirStock(miListaPedidosVenta As List(Of PedidosVenta)) As List(Of PedidosVenta)

        Try


            Dim NewmListaPedidosVenta As List(Of PedidosVenta) = New List(Of PedidosVenta)
            'Obtenemos todos los materiales
            Dim ListaMateriales = miListaPedidosVenta.Select(Function(s) s.Material).Distinct().ToList()
            'Recorremos los materiales
            For Each Material As String In ListaMateriales
                'Obtenemos los elementos de ese Material 
                Dim ElementosXMaterial = miListaPedidosVenta.Where(Function(w) w.Material = Material).OrderBy(Function(o) o.FechaPrevista).ThenBy(Function(o) o.NumPedido).ToList()
                'Obtenemos el Stock a Repartir
                Dim Stock = ElementosXMaterial.FirstOrDefault().StockPedidosVenta
                If Stock > 0 Then

                    'Recorremos los elementos donde se repartira el stock X Material
                    For index = 0 To ElementosXMaterial.Count - 1
                        Dim elemento = ElementosXMaterial(index)
                        If Stock > 0 Then

                            'If index = 0 Then
                            'Si es el primer elemento se pone el total del stock
                            elemento.NuevoStockAPedidoVenta = Stock


                            'Descontamos el stock que fue asignado
                            Stock = Stock - elemento.UnidadesPtes
                            'Else

                            'End If
                            'elemento.Kilos
                        Else
                            elemento.NuevoStockAPedidoVenta = 0
                        End If
                        'Correccion de calculo de cantidada Pendiente (KilosPtes) 05-09-2023
                        If elemento.UnidadesPtes >= elemento.NuevoStockAPedidoVenta Then
                            If elemento.Material = "70903201" Then
                                Dim g = 5
                            End If
                            elemento.KilosPtes = elemento.UnidadesPtes - elemento.NuevoStockAPedidoVenta

                        Else
                            elemento.KilosPtes = 0
                        End If

                        'Obtenemos los kg por base 1000 (ejemplo, si son 50 Kg, es decir 1000 unidades pesan 50 Kg)
                        Dim Kg As Double = 0
                        For Each materialA In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
                            If materialA.UnidadMedida.ToUpper().Trim() = "KG" Then
                                Kg = materialA.Cantidad
                                Exit For
                            End If
                        Next
                        If Kg = 0 Then
                            For Each materialA In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
                                For Each materialB In materialA.Material.CabLista_Material.MaterialesLista
                                    If materialB.UnidadMedida.ToUpper().Trim() = "KG" Then
                                        Kg = materialB.Cantidad
                                        Exit For
                                    End If

                                Next
                                If Kg > 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        elemento = asignarPuestoTrabajoyFechaPlan(elemento)
                        'Tenemos las Unidades Pendientes, agregaremos los kg Pendientes
                        elemento.KgPtes = elemento.KilosPtes * (Kg / 1000)

                        Dim sCodigoMaterialFab As String = ""
                        If elemento.MaterialDetalle.ProductoFormula.Count > 0 Then
                            sCodigoMaterialFab = elemento.MaterialDetalle.ProductoFormula(0).Codigo.Trim
                        End If
                        If sCodigoMaterialFab = "" Then
                            For Each semi As ListaMaterial In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
                                If semi.Material.Tipo = "1215" Then
                                    Dim MaterialDetalleSemielaborado = New Material(semi.Material.Codigo)
                                    If MaterialDetalleSemielaborado.ProductoFormula.Count > 0 Then
                                        sCodigoMaterialFab = MaterialDetalleSemielaborado.ProductoFormula(0).Codigo.Trim
                                    End If

                                    'Dim MaterialDetalleSemielaboradoV2 = New Material(semi.CodigoMaterial)
                                    'sCodigoMaterialFab = MaterialDetalleSemielaboradoV2.ProductoFormula(0).Codigo.Trim
                                End If
                            Next
                        End If
                        elemento.CodigoMaterialFab = sCodigoMaterialFab

                        NewmListaPedidosVenta.Add(elemento)

                    Next
                Else
                    'Si el Stock es 0 Agrega todos los elementos del Material y sigue al siguiente material
                    'No hay nada que repartir
                    'Se calcula la cantidad Pendiente y los kg Pendientes para los elementos sin Stock
                    'Recorremos los elementos donde se repartira el stock X Material
                    Dim NewmListaPedidosVentaSinStock As List(Of PedidosVenta) = New List(Of PedidosVenta)
                    For index = 0 To ElementosXMaterial.Count - 1
                        Dim elemento = ElementosXMaterial(index)
                        'Correccion de calculo de cantidada Pendiente (KilosPtes) 05-09-2023
                        If elemento.UnidadesPtes >= elemento.NuevoStockAPedidoVenta Then
                            If elemento.Material = "70903201" Then
                                Dim g = 5
                            End If
                            elemento.KilosPtes = elemento.UnidadesPtes - elemento.NuevoStockAPedidoVenta

                        Else
                            elemento.KilosPtes = 0
                        End If
                        'Obtenemos los kg por base 1000 (ejemplo, si son 50 Kg, es decir 1000 unidades pesan 50 Kg)
                        Dim Kg As Double = 0
                        For Each materialA In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
                            If materialA.UnidadMedida.ToUpper().Trim() = "KG" Then
                                Kg = materialA.Cantidad
                                Exit For
                            End If
                        Next
                        If Kg = 0 Then
                            For Each materialA In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
                                For Each materialB In materialA.Material.CabLista_Material.MaterialesLista
                                    If materialB.UnidadMedida.ToUpper().Trim() = "KG" Then
                                        Kg = materialB.Cantidad
                                        Exit For
                                    End If

                                Next
                                If Kg > 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        elemento = asignarPuestoTrabajoyFechaPlan(elemento)
                        'Tenemos las Unidades Pendientes, agregaremos los kg Pendientes
                        elemento.KgPtes = elemento.KilosPtes * (Kg / 1000)

                        Dim sCodigoMaterialFab As String = ""
                        If elemento.MaterialDetalle.ProductoFormula.Count > 0 Then
                            sCodigoMaterialFab = elemento.MaterialDetalle.ProductoFormula(0).Codigo.Trim
                        End If
                        If sCodigoMaterialFab = "" Then
                            For Each semi As ListaMaterial In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
                                If semi.Material.Tipo = "1215" Then
                                    Dim MaterialDetalleSemielaborado = New Material(semi.Material.Codigo)
                                    If MaterialDetalleSemielaborado.ProductoFormula.Count > 0 Then
                                        sCodigoMaterialFab = MaterialDetalleSemielaborado.ProductoFormula(0).Codigo.Trim
                                    End If
                                    'Dim MaterialDetalleSemielaboradoV2 = New Material(semi.CodigoMaterial)
                                    'sCodigoMaterialFab = MaterialDetalleSemielaboradoV2.ProductoFormula(0).Codigo.Trim
                                End If
                            Next
                        End If
                        elemento.CodigoMaterialFab = sCodigoMaterialFab

                        NewmListaPedidosVentaSinStock.Add(elemento)
                    Next


                    NewmListaPedidosVenta.AddRange(NewmListaPedidosVentaSinStock)
                    Continue For
                End If


            Next
            Return NewmListaPedidosVenta
        Catch ex As NullReferenceException
            'XtraMessageBox.Show("ERROR " & ex.Message & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", "ADMINISTRADOR",
            '                      MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            'XtraMessageBox.Show("ERROR " & ex.Message & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", "ADMINISTRADOR",
            '             MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function


    'Private Function RepartirStock(miListaPedidosVenta As List(Of PedidosVenta)) As List(Of PedidosVenta)
    '    Dim NewmListaPedidosVenta As List(Of PedidosVenta) = New List(Of PedidosVenta)
    '    'Obtenemos todos los materiales
    '    Dim ListaMateriales = miListaPedidosVenta.Select(Function(s) s.Material).Distinct().ToList()
    '    'Recorremos los materiales
    '    For Each Material As String In ListaMateriales
    '        'Obtenemos los elementos de ese Material 
    '        Dim ElementosXMaterial = miListaPedidosVenta.Where(Function(w) w.Material = Material).OrderBy(Function(o) o.FechaPrevista).ThenBy(Function(o) o.NumPedido).ToList()
    '        'Obtenemos el Stock a Repartir
    '        Dim Stock = ElementosXMaterial.FirstOrDefault().StockPedidosVenta
    '        If Stock > 0 Then

    '            'Recorremos los elementos donde se repartira el stock X Material
    '            For index = 0 To ElementosXMaterial.Count - 1
    '                Dim elemento = ElementosXMaterial(index)
    '                If Stock > 0 Then

    '                    'If index = 0 Then
    '                    'Si es el primer elemento se pone el total del stock
    '                    elemento.NuevoStockAPedidoVenta = Stock


    '                    'Descontamos el stock que fue asignado
    '                    Stock = Stock - elemento.UnidadesPtes
    '                    'Else

    '                    'End If
    '                    'elemento.Kilos
    '                Else
    '                    elemento.NuevoStockAPedidoVenta = 0
    '                End If
    '                'Correccion de calculo de cantidada Pendiente (KilosPtes) 05-09-2023
    '                If elemento.UnidadesPtes >= elemento.NuevoStockAPedidoVenta Then
    '                    If elemento.Material = "70903201" Then
    '                        Dim g = 5
    '                    End If
    '                    elemento.KilosPtes = elemento.UnidadesPtes - elemento.NuevoStockAPedidoVenta

    '                Else
    '                    elemento.KilosPtes = 0
    '                End If

    '                'Obtenemos los kg por base 1000 (ejemplo, si son 50 Kg, es decir 1000 unidades pesan 50 Kg)
    '                Dim Kg As Double = 0
    '                For Each materialA In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
    '                    If materialA.UnidadMedida.ToUpper().Trim() = "KG" Then
    '                        Kg = materialA.Cantidad
    '                        Exit For
    '                    End If
    '                Next
    '                If Kg = 0 Then
    '                    For Each materialA In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
    '                        For Each materialB In materialA.Material.CabLista_Material.MaterialesLista
    '                            If materialB.UnidadMedida.ToUpper().Trim() = "KG" Then
    '                                Kg = materialB.Cantidad
    '                                Exit For
    '                            End If

    '                        Next
    '                        If Kg > 0 Then
    '                            Exit For
    '                        End If
    '                    Next
    '                End If
    '                elemento = asignarPuestoTrabajoyFechaPlan(elemento)
    '                'Tenemos las Unidades Pendientes, agregaremos los kg Pendientes
    '                elemento.KgPtes = elemento.KilosPtes * (Kg / 1000)
    '                NewmListaPedidosVenta.Add(elemento)

    '            Next
    '        Else
    '            'Si el Stock es 0 Agrega todos los elementos del Material y sigue al siguiente material
    '            'No hay nada que repartir
    '            'Se calcula la cantidad Pendiente y los kg Pendientes para los elementos sin Stock
    '            'Recorremos los elementos donde se repartira el stock X Material
    '            Dim NewmListaPedidosVentaSinStock As List(Of PedidosVenta) = New List(Of PedidosVenta)
    '            For index = 0 To ElementosXMaterial.Count - 1
    '                Dim elemento = ElementosXMaterial(index)
    '                'Correccion de calculo de cantidada Pendiente (KilosPtes) 05-09-2023
    '                If elemento.UnidadesPtes >= elemento.NuevoStockAPedidoVenta Then
    '                    If elemento.Material = "70903201" Then
    '                        Dim g = 5
    '                    End If
    '                    elemento.KilosPtes = elemento.UnidadesPtes - elemento.NuevoStockAPedidoVenta

    '                Else
    '                    elemento.KilosPtes = 0
    '                End If
    '                'Obtenemos los kg por base 1000 (ejemplo, si son 50 Kg, es decir 1000 unidades pesan 50 Kg)
    '                Dim Kg As Double = 0
    '                For Each materialA In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
    '                    If materialA.UnidadMedida.ToUpper().Trim() = "KG" Then
    '                        Kg = materialA.Cantidad
    '                        Exit For
    '                    End If
    '                Next
    '                If Kg = 0 Then
    '                    For Each materialA In elemento.MaterialDetalle.CabLista_Material.MaterialesLista
    '                        For Each materialB In materialA.Material.CabLista_Material.MaterialesLista
    '                            If materialB.UnidadMedida.ToUpper().Trim() = "KG" Then
    '                                Kg = materialB.Cantidad
    '                                Exit For
    '                            End If

    '                        Next
    '                        If Kg > 0 Then
    '                            Exit For
    '                        End If
    '                    Next
    '                End If
    '                elemento = asignarPuestoTrabajoyFechaPlan(elemento)
    '                'Tenemos las Unidades Pendientes, agregaremos los kg Pendientes
    '                elemento.KgPtes = elemento.KilosPtes * (Kg / 1000)
    '                NewmListaPedidosVentaSinStock.Add(elemento)
    '            Next


    '            NewmListaPedidosVenta.AddRange(NewmListaPedidosVentaSinStock)
    '            Continue For
    '        End If


    '    Next
    '    Return NewmListaPedidosVenta
    'End Function

    Private Function asignarPuestoTrabajoyFechaPlan(pedido As PedidosVenta) As PedidosVenta
        Dim ListaNombres = New List(Of String)

        For Each hoja In pedido.MaterialDetalle.HojasDeRuta
            ListaNombres.AddRange(hoja.PuestosTrabajoHojaRuta.Select(Function(s) s.Nombre.Trim))
        Next

        If ListaNombres.Count > 0 Then
            If ListaNombres.Distinct().Count = 1 Then
                pedido.NombrePuestoTrabajo = pedido.MaterialDetalle.HojasDeRuta(0).PuestosTrabajoHojaRuta(0).Nombre.Trim()
            Else
                pedido.NombrePuestoTrabajo = pedido.MaterialDetalle.HojasRutaDefecto.PuestosTrabajoHojaRuta(0).Nombre.Trim()
            End If

            Dim misTurnos As New List(Of Calendario)
            'If NombresPuestoTrabajo.Count = 1 Then
            '    misTurnos = DameTurnosMaquina(Pedido.FechaPrevista.AddDays(-(Pedido.MaterialDetalle.DiasFabPropia + Pedido.MaterialDetalle.DiasPP) * 5), NombresPuestoTrabajo(0).CodigoPuestoTrabajo)
            'Else
            Dim Cod = pedido.MaterialDetalle.HojasRutaDefecto.PuestosTrabajoHojaRuta(0).CodigoPuestoTrabajo
            misTurnos = DameTurnosMaquina(pedido.FechaPrevista.AddDays(-(pedido.MaterialDetalle.DiasFabPropia + pedido.MaterialDetalle.DiasPP) * 5), Cod)
            'End If

            Dim turnosARestar = (pedido.MaterialDetalle.DiasFabPropia + pedido.MaterialDetalle.DiasPP) * 3
            'Obtenemos en que turno esta la la Fecha Prevista
            Dim IdTurno = misTurnos.Where(Function(w) pedido.FechaPrevista >= w.InicioTurno).ToList().Select(Function(s) s.Id).Max() - turnosARestar
            If IdTurno > 0 Then
                pedido.FechaPlan = misTurnos.Where(Function(w) w.Id = IdTurno).FirstOrDefault().InicioTurno
            Else
                pedido.FechaPlan = misTurnos.FirstOrDefault().InicioTurno
            End If
        End If



        Return pedido
    End Function
End Module
