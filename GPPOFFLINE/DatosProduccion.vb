

Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP
Imports System.Data.SqlClient
Imports System.Configuration

Public Module DatosProduccion

    Public Function FnlogApp(ByVal sMsg As String) As String
        Try
            Dim Ruta As String = ConfigurationManager.AppSettings("rutalog").ToString()
            Dim oSW As System.IO.StreamWriter = New System.IO.StreamWriter(Ruta & "\Log_" & DateTime.Now.Date.ToString("yyyy-MM-dd") & ".txt", True)
            Dim scomando As String = String.Empty
            oSW.WriteLine(DateTime.Now & " || Evento: " & sMsg)
            oSW.Flush()
            oSW.Close()
            Return Ruta & "\Log_" & DateTime.Now.Date.ToString("yyyy-MM-dd") & ".txt"
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function DameMaterialesPullSystem() As List(Of Material)
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = "SELECT * FROM Materiales M with(nolock)  WHERE  M.maFecIniPS <> '1900-01-01' and M.maFecFinPS <> '1900-01-01'"

            DameMaterialesPullSystem = New List(Of Material)

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DameMaterialesPullSystem.Add(New Material(Codigo:=CStr(NoNull(miRegistro.Item("maCod"), "A")),
                                                        Tipo:=CStr(NoNull(miRegistro.Item("maTipoMat"), "A")),
                                                        Grupo:=CStr(NoNull(miRegistro.Item("maGrupoArt"), "A")),
                                                        UnidadMedida:=CStr(NoNull(miRegistro.Item("maUMBase"), "A")),
                                                        Nombre:=CStr(NoNull(miRegistro.Item("maNombre"), "A")),
                                                        Lista_Mat:=CStr(NoNull(miRegistro.Item("maListaMaterial"), "A")),
                                                        Familia_Envasado:=CByte(NoNull(miRegistro.Item("maFamiliaEnvasado"), "D")),
                                                        Fecha_IniPS:=CDate(NoNull(miRegistro.Item("maFecIniPS"), "DT")),
                                                        Fecha_FinPS:=CDate(NoNull(miRegistro.Item("maFecFinPS"), "DT")),
                                                        Dias_PP:=CByte(NoNull(miRegistro.Item("maDiasPP"), "D")),
                                                        Stock_MaxPS:=CInt(NoNull(miRegistro.Item("maStokMaxPS"), "D")),
                                                        Stock_MinPS:=CInt(NoNull(miRegistro.Item("maStokMinPS"), "D")),
                                                        Activo:=CBool(NoNull(miRegistro.Item("maActivo"), "D")),
                                                        Lote_Minimo:=CInt(NoNull(miRegistro.Item("maLoteMin"), "D")),
                                                        Lote_Maximo:=CInt(NoNull(miRegistro.Item("maLoteMax"), "D")),
                                                        Lote_Fijo:=CInt(NoNull(miRegistro.Item("maLoteFijo"), "D")),
                                                        Redondeo_Lote:=CInt(NoNull(miRegistro.Item("maRedondeo"), "D")),
                                                        Tipo_TamañoLote:=CStr(NoNull(miRegistro.Item("maTipoTamLote"), "A")),
                                                        Dias_FabPropia:=CInt(NoNull(miRegistro.Item("maDiasFabPropia"), "D")),
                                                        Grupo_HojaRuta:=CStr(NoNull(miRegistro.Item("maGrupoHR"), "A")),
                                                        Contador_HojaRuta:=CStr(NoNull(miRegistro.Item("maContHR"), "A")),
                                                        Grupo_Compra:=CStr(NoNull(miRegistro.Item("maGrupoCompra"), "A")),
                                                        Mostrar_Informes:=CBool(NoNull(miRegistro.Item("mnMostrarInformes"), "D")),
                                                        Unidades_Pack:=CInt(NoNull(miRegistro.Item("maUnidadesPACK"), "D")),
                                                        UnidadesPorPalet:=CInt(NoNull(miRegistro.Item("maUnidadesPalet"), "D")),
                                                        MesesLoteCarga:=CInt(NoNull(miRegistro.Item("maMesesLoteCarga"), "D"))))
                Next
            End If
        Catch ex As Exception
            DameMaterialesPullSystem = New List(Of Material)
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Function DameTurnosMaquina(ByVal FechaInicio As Date,
                                         ByVal CodPuestotrabajo As Integer) As List(Of Calendario)
        Try
            Dim sSql As String = "SELECT * FROM CalendarioProduccion " &
                                 "WHERE caPuestoTrabajo=" & CodPuestotrabajo &
                                 " And ( caFinTurnoM >= '" & FechaInicio & "' " &
                                 "or cafinTurnoT >= '" & FechaInicio & "' " &
                                 "or cafinTurnoN >= '" & FechaInicio & "' ) " &
                                 "ORDER BY caInicioTurnoM "


            Dim DTDatos As New DataTable
            Dim DameCalendarioPorPuestoTrabajo = New List(Of PuestosTrabajoDias)

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows

                    DameCalendarioPorPuestoTrabajo.Add(New PuestosTrabajoDias(CodPuestotrabajo:=CInt(NoNull(miRegistro.Item("caPuestoTrabajo"), "D")),
                                                                Fecha:=CDate(NoNull(miRegistro.Item("caFecha"), "DT")),
                                                                Dias:=CDate(NoNull(miRegistro.Item("caFecha"), "DT")).Day,
                                                                InicioM:=CDate(NoNull(miRegistro.Item("caInicioTurnoM"), "DT")), 'CStr(NoNull(miRegistro.Item("clInicioMañana"), "A")),
                                                                InicioT:=CDate(NoNull(miRegistro.Item("caInicioTurnoT"), "DT")),
                                                                InicioN:=CDate(NoNull(miRegistro.Item("caInicioTurnoN"), "DT")),
                                                                FinM:=CDate(NoNull(miRegistro.Item("caFinTurnoM"), "DT")),
                                                                FinT:=CDate(NoNull(miRegistro.Item("caFinTurnoT"), "DT")),
                                                                FinN:=CDate(NoNull(miRegistro.Item("caFinTurnoN"), "DT")),
                                                                TurnoM:=CBool(NoNull(miRegistro.Item("caTurnoM"), "D")),
                                                                TurnoT:=CBool(NoNull(miRegistro.Item("caTurnoT"), "D")),
                                                                TurnoN:=CBool(NoNull(miRegistro.Item("caTurnoN"), "D")),
                                                                Operarios:=CInt(NoNull(miRegistro.Item("caNumOperarios"), "D")),
                                                                OperariosM:=CInt(NoNull(miRegistro.Item("caNumOperariosM"), "D")),
                                                                OperariosT:=CInt(NoNull(miRegistro.Item("caNumOperariosT"), "D")),
                                                                OperariosN:=CInt(NoNull(miRegistro.Item("caNumOperariosN"), "D")),
                                                                Observaciones:=CStr(NoNull(miRegistro.Item("caComentarios"), "A")),
                                                                HoraInicioM:=New TimeSpan(CDate(NoNull(miRegistro.Item("caInicioTurnoM"), "DT")).Hour, CDate(NoNull(miRegistro.Item("caInicioTurnoM"), "DT")).Minute, CDate(NoNull(miRegistro.Item("caInicioTurnoM"), "DT")).Second),
                                                                HoraFinM:=New TimeSpan(CDate(NoNull(miRegistro.Item("caFinTurnoM"), "DT")).Hour, CDate(NoNull(miRegistro.Item("caFinTurnoM"), "DT")).Minute, CDate(NoNull(miRegistro.Item("caFinTurnoM"), "DT")).Second),
                                                                HoraInicioT:=New TimeSpan(CDate(NoNull(miRegistro.Item("caInicioTurnoT"), "DT")).Hour, CDate(NoNull(miRegistro.Item("caInicioTurnoT"), "DT")).Minute, CDate(NoNull(miRegistro.Item("caInicioTurnoT"), "DT")).Second),
                                                                HoraFinT:=New TimeSpan(CDate(NoNull(miRegistro.Item("caFinTurnoT"), "DT")).Hour, CDate(NoNull(miRegistro.Item("caFinTurnoT"), "DT")).Minute, CDate(NoNull(miRegistro.Item("caFinTurnoT"), "DT")).Second),
                                                                HoraInicioN:=New TimeSpan(CDate(NoNull(miRegistro.Item("caInicioTurnoN"), "DT")).Hour, CDate(NoNull(miRegistro.Item("caInicioTurnoN"), "DT")).Minute, CDate(NoNull(miRegistro.Item("caInicioTurnoN"), "DT")).Second),
                                                                HoraFinN:=New TimeSpan(CDate(NoNull(miRegistro.Item("caFinTurnoN"), "DT")).Hour, CDate(NoNull(miRegistro.Item("caFinTurnoN"), "DT")).Minute, CDate(NoNull(miRegistro.Item("caFinTurnoN"), "DT")).Second)))

                Next
            End If


            DameTurnosMaquina = New List(Of Calendario)
            Dim ID As Integer = 0

            If DameCalendarioPorPuestoTrabajo.Count > 0 Then
                For Each miRegistro In DameCalendarioPorPuestoTrabajo

                    If miRegistro.TurnoM Then
                        ID += 1
                        Dim minutos As Long = DateDiff(DateInterval.Minute, miRegistro.InicioM, miRegistro.FinM)
                        DameTurnosMaquina.Add(New Calendario(
                                                              Codigo:=ID,
                                                             Cod_PuestoTrabajo:=miRegistro.CodPuestotrabajo,
                                                             Fecha_turno:=miRegistro.Fecha,
                                                             Turno:=CChar(NoNull("M", "A")),
                                                             Fec_Inicio:=CDate(NoNull(miRegistro.InicioM, "DT")),
                                                             Fec_Fin:=CDate(NoNull(miRegistro.FinM, "DT")),
                                                             Num_Operarios:=CInt(NoNull(miRegistro.Operarios, "D")),
                                                             Minutos_Turno:=CInt(minutos))
                                                             )
                    End If
                    If miRegistro.TurnoT Then
                        ID += 1
                        Dim minutos As Long = DateDiff(DateInterval.Minute, miRegistro.InicioT, miRegistro.FinT)
                        DameTurnosMaquina.Add(New Calendario(
                                                              Codigo:=ID,
                                                             Cod_PuestoTrabajo:=miRegistro.CodPuestotrabajo,
                                                             Fecha_turno:=miRegistro.Fecha,
                                                             Turno:=CChar(NoNull("T", "A")),
                                                             Fec_Inicio:=CDate(NoNull(miRegistro.InicioT, "DT")),
                                                             Fec_Fin:=CDate(NoNull(miRegistro.FinT, "DT")),
                                                             Num_Operarios:=CInt(NoNull(miRegistro.Operarios, "D")),
                                                             Minutos_Turno:=CInt(minutos))
                                                             )
                    End If
                    If miRegistro.TurnoN Then
                        ID += 1
                        Dim minutos As Long = DateDiff(DateInterval.Minute, miRegistro.InicioN, miRegistro.FinN)
                        DameTurnosMaquina.Add(New Calendario(
                                                              Codigo:=ID,
                                                             Cod_PuestoTrabajo:=miRegistro.CodPuestotrabajo,
                                                             Fecha_turno:=miRegistro.Fecha,
                                                             Turno:=CChar(NoNull("N", "A")),
                                                             Fec_Inicio:=CDate(NoNull(miRegistro.InicioN, "DT")),
                                                             Fec_Fin:=CDate(NoNull(miRegistro.FinN, "DT")),
                                                             Num_Operarios:=CInt(NoNull(miRegistro.Operarios, "D")),
                                                             Minutos_Turno:=CInt(minutos))
                                                             )
                    End If

                Next
            Else
                DameTurnosMaquina = New List(Of Calendario)
            End If

        Catch ex As Exception
            DameTurnosMaquina = New List(Of Calendario)
            Throw New Exception(ex.Message)
        End Try
    End Function


    Public Function DameHoraFin(ByVal SegundosFab As Long,
                                ByVal FechaInicio As Date,
                                ByVal CodPuestoTrabajo As Integer,
                                Optional ByRef dtProduccion As DataTable = Nothing,
                                Optional NombrePuestoTrabajo As String = "",
                                Optional CodMaterial As String = "",
                                Optional NombreMaterial As String = "",
                                Optional Operarios As Integer = 0,
                                Optional CantidadRestante As Integer = 0,
                                Optional IdFabricacion As Integer = 0,
                                Optional Unidad As String = "",
                                Optional incluirTurnos As Boolean = True) As Date
        Try
            Dim dFecIni As Date = FechaInicio
            Dim dFecTurno As Date = FechaInicio
            Dim dFecIniDia As Date
            Dim dFecFinDia As Date
            Dim dFechaPrevFin As Date
            Dim iSegundosPtes As Long = SegundosFab
            Dim misTurnos As New List(Of Calendario)

            misTurnos = DameTurnosMaquina(FechaInicio, CodPuestoTrabajo)
            misTurnos = misTurnos.Where(Function(w) w.FinTurno >= FechaInicio).ToList()

            Dim i As Integer = 0

            If misTurnos.Count > 0 And incluirTurnos = True Then
                Do While i < misTurnos.Count
                    dFecIniDia = misTurnos(i).InicioTurno
                    dFecFinDia = misTurnos(i).FinTurno

                    If dFecIni < dFecIniDia Then
                        dFecIni = dFecIniDia
                    End If

                    'Saco la fecha prevista de finalización
                    dFechaPrevFin = DateAdd(DateInterval.Second, iSegundosPtes, dFecIni)

                    If dFechaPrevFin >= dFecIniDia And dFechaPrevFin <= dFecFinDia Then
                        If Not dtProduccion Is Nothing Then

                            dtProduccion.Rows.Add(dFecIni,
                                              dFechaPrevFin,
                                              NombrePuestoTrabajo,
                                              CodMaterial,
                                              NombreMaterial.Trim(),
                                              IIf(misTurnos(i).Operarios = 0, Operarios, misTurnos(i).Operarios),
                                              misTurnos(i).Turno,
                                              misTurnos(i).FechaTurno,
                                              DatePart(DateInterval.WeekOfYear, misTurnos(i).FechaTurno),
                                              CantidadRestante,
                                              IdFabricacion,
                                              Unidad)
                        End If

                        Return dFechaPrevFin
                    Else
                        If dFecIniDia > dFechaPrevFin Then
                            'Si el turno empieza mas tarde de la hora inicio prevista
                            dFecIni = dFecIniDia
                        Else
                            If dFechaPrevFin > dFecFinDia Then
                                'La fecha previsto fin es posterior a la fecha del turno
                                iSegundosPtes = CInt(iSegundosPtes - (DateDiff(DateInterval.Second, dFecIni, dFecFinDia)))
                            End If
                        End If
                    End If

                    '' ************* VICENTE, Aqui he generado un nuevo parcial - VER SI SE PUEDE GUARDAR EN EL Datatable del DATASET ****************
                    If Not dtProduccion Is Nothing Then
                        dtProduccion.Rows.Add(dFecIni,
                                              dFecFinDia,'dFechaPrevFin,
                                              NombrePuestoTrabajo,
                                              CodMaterial,
                                              NombreMaterial.Trim(),
                                              IIf(misTurnos(i).Operarios = 0, Operarios, misTurnos(i).Operarios),
                                              misTurnos(i).Turno,
                                              misTurnos(i).FechaTurno,
                                              DatePart(DateInterval.WeekOfYear, misTurnos(i).FechaTurno),
                                              CantidadRestante,
                                              IdFabricacion,
                                              Unidad)
                    End If

                    i += 1
                Loop

                'Si llego aqui es que no hay mas turnos
                Return DateAdd(DateInterval.Second, iSegundosPtes, dFecIni)

            Else
                If Not dtProduccion Is Nothing Then
                    dtProduccion.Rows.Add(dFecIni,
                                          DateAdd(DateInterval.Second, iSegundosPtes, dFecIni),
                                          NombrePuestoTrabajo,
                                          CodMaterial,
                                          NombreMaterial.Trim(),
                                          Operarios,
                                          TurnoManana,
                                          dFecIni,
                                          DatePart(DateInterval.WeekOfYear, dFecIni),
                                          CantidadRestante,
                                          IdFabricacion,
                                          Unidad)
                End If

                Return DateAdd(DateInterval.Second, iSegundosPtes, dFecIni)
            End If

        Catch ex As Exception
            DameHoraFin = ConstantesGPP.FechaGlobal
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function

    Public Function Dame_Minutos_Tiempo_Fabricacion(Cantidad As Integer,
                                               Hoja_de_Ruta As HojaRuta,
                                               Puestotrabajo As Integer,
                                               IncluirPreparacion As Boolean, GrupoHojaRuta As String, ContadorHojaRuta As String) As Integer
        Try
            Dim iMinutos As Integer = 0

            For Each miOper In Hoja_de_Ruta.OperacHojaRutaLista
                If Puestotrabajo = miOper.CodigoPuestoDeTrabajo AndAlso miOper.CantidadBase <> 0 AndAlso miOper.GrupoSAP.Trim = GrupoHojaRuta AndAlso miOper.ContGrupoSAP.Trim = ContadorHojaRuta Then
                    iMinutos += CInt(miOper.MinutosLimpieza) +
                                CInt(Cantidad * miOper.MinutosMaquina / miOper.CantidadBase) +
                                CInt(IIf(IncluirPreparacion = True, miOper.MinutosPreparacion, 0))
                End If
            Next

            Return iMinutos

        Catch ex As Exception
            Dame_Minutos_Tiempo_Fabricacion = 0
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try


    End Function

    Public Function DameListaPuestosTrabajoMaterial(ByVal Material As String) As List(Of PuestosTrabajo)
        Dim puestoTrabajo As New List(Of PuestosTrabajo)
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = "SELECT DISTINCT " &
                                 "ISNULL(PT.ptNombre,'') AS PuestosTrabajo , PT.ptCod as Codigo  " &
                                 "From Materiales MA " &
                                 "INNER Join OperacionesHojaRuta OHR ON MA.maGrupoHR = OHR.opGrupoSAP And MA.maContHR=OHR.opContGrupoSAP  " &
                                 "INNER Join PuestosTrabajo PT ON PT.ptCod=OHR.opPuestoTrabajo Or (OHR.opClaveControl='ZPE4' AND PT.ptNombre='EXTERNO')  " &
                                 "WHERE MA.maTipoMat =" & TipoMaterial.ProdTerminado & " And MA.maCod ='" & Material & "'"


            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                If DTDatos.Rows.Count > 0 Then
                    For index = 0 To DTDatos.Rows.Count - 1
                        Dim pue As New PuestosTrabajo
                        pue.Nombre = UTrim(DTDatos.Rows(index).Item("PuestosTrabajo"))
                        pue.CodigoPuestoTrabajo = CInt(NoNull(DTDatos.Rows(index).Item("Codigo"), "N"))
                        puestoTrabajo.Add(pue)
                        'listaPuestosTrabajo.Add(UTrim(DTDatosPuestos.Rows(index).Item("PuestosTrabajo")))
                    Next

                Else
                    puestoTrabajo = New List(Of PuestosTrabajo)
                End If

            End If
            Return puestoTrabajo
        Catch ex As Exception
            puestoTrabajo = New List(Of PuestosTrabajo)
            Throw New Exception(ex.Message & " - -" & "()", ex)
        End Try
    End Function
    Public Function DameFabricacionExistente(ByVal NumPedSap As String,
                                             ByVal PosPedSap As Integer,
                                             ByVal Estado_Fabricacion As EstadoFabricacion,
                                             ByVal CodPuestoTrabajo As Integer,
                                             ByVal IdEnvio As Integer,
                                             Optional ByVal OrdenarOrdenFabrcacion As Boolean = True) As List(Of Fabricaciones)
        Try
            Dim DTDatos As New DataTable
            Dim sWhere As String = ""
            Dim sSql As String = "SELECT * " &
                                 "FROM Fabricaciones " &
                                 "WHERE 1=1 "


            If NumPedSap.Trim.Length <> 0 Then
                sSql &= " AND opNumPedSAP = '" & NumPedSap & "' and " & " opPosPedSAP=" & PosPedSap
            End If

            If IdEnvio <> 0 Then
                sSql &= " AND opIdEnvio = " & IdEnvio
            End If

            If CodPuestoTrabajo <> 0 Then
                sSql &= " AND opPuestoTrabajo=" & CodPuestoTrabajo
            End If

            If Estado_Fabricacion <> EstadoFabricacion.Ninguna Then
                sSql &= " AND opEnMarcha=" & Estado_Fabricacion
            End If

            If OrdenarOrdenFabrcacion = True Then
                sSql &= " ORDER BY openMarcha desc,opOrdenMaq"
            End If

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                DameFabricacionExistente = (From elemento In DTDatos
                                            Select New Fabricaciones(CInt(NoNull(elemento.Item("opIdFab"), "D")),
                                                        CInt(NoNull(elemento.Item("opPuestoTrabajo"), "D")),
                                                        CByte(NoNull(elemento.Item("opEnmarcha"), "D")),
                                                        CInt(NoNull(elemento.Item("opOrdenMaq"), "D")),
                                                        CStr(NoNull(elemento.Item("opMaterial"), "A")),
                                                        CInt(NoNull(elemento.Item("opCantidadPlanif"), "D")),
                                                        CInt(NoNull(elemento.Item("opCantidadFab"), "D")),
                                                        CDate(NoNull(elemento.Item("opTfecIni"), "DT")),
                                                        CChar(NoNull(elemento.Item("opTurno"), "A")),
                                                        CDate(NoNull(elemento.Item("opFechaIni"), "DT")),
                                                        CDate(NoNull(elemento.Item("opFechaFin"), "DT")),
                                                        CInt(NoNull(elemento.Item("opOrdenFabSAP"), "D")),
                                                        CInt(NoNull(elemento.Item("opOrdenEnvSAP"), "D")),
                                                        CDate(NoNull(elemento.Item("opFechaPrevFin"), "DT")),
                                                        CStr(NoNull(elemento.Item("opListaMaterial"), "A")),
                                                        CStr(NoNull(elemento.Item("opGrupoHR"), "A")),
                                                        CStr(NoNull(elemento.Item("opContHR"), "A")),
                                                        CInt(NoNull(elemento.Item("opSigPtoTrabajo"), "D")),
                                                        CInt(NoNull(elemento.Item("opCantidadFabSAP"), "D")),
                                                        CStr(NoNull(elemento.Item("opNumeroLoteSAP"), "A")),
                                                        CShort(NoNull(elemento.Item("opEquipo"), "D")),
                                                        CStr(NoNull(elemento.Item("opNumPedSAP"), "A")),
                                                        CInt(NoNull(elemento.Item("opPosPedSAP"), "D")),
                                                        CStr(NoNull(elemento.Item("opFormato"), "A")),
                                                        CStr(NoNull(elemento.Item("opMaterialPadre"), "A")),
                                                        CInt(NoNull(elemento.Item("opCantidadPlanifPadre"), "D")),
                                                        CInt(NoNull(elemento.Item("opCantidadFabRechazada"), "D")),
                                                        CInt(NoNull(elemento.Item("opCantidadFabBuenas"), "D")),
                                                        CInt(NoNull(elemento.Item("opCantidadReprocTap"), "D")),
                                                        CInt(NoNull(elemento.Item("opCantidadReprocEti"), "D")),
                                                        CInt(NoNull(elemento.Item("opMinutosFabObj"), "D")),
                                                        CInt(NoNull(elemento.Item("opMinutosFabReal"), "D")),
                                                        CStr(NoNull(elemento.Item("opNombreMaterial"), "A")),
                                                        CInt(NoNull(elemento.Item("opIdEnvio"), "D")))).ToList

            Else
                DameFabricacionExistente = New List(Of Fabricaciones)
            End If
        Catch ex As Exception
            DameFabricacionExistente = New List(Of Fabricaciones)
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function EliminarPedidosTemporal() As Boolean
        Try
            Dim sSql As String = " DELETE FROM PedidosVentaOFFLINETMP "

            EliminarPedidosTemporal = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarPedidosTemporal = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function
    Public Function EliminarPedidos(ByVal FechaInicio As String, ByVal FechaFinal As String) As Boolean
        Try
            Dim sSql As String = " DELETE FROM PedidosVentaOFFLINE " '&
            '" WHERE pvFechaPrevista >= '" & FechaInicio & "' and pvFechaPrevista <= '" & FechaFinal & "'"

            EliminarPedidos = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarPedidos = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function InsertarPedidosTMP(ByVal ListaPedidos As List(Of PedidosVenta)) As Boolean
        Dim insertar As Boolean = True
        Try

            For Each miPedido In ListaPedidos
                Dim sSql As String = "INSERT INTO [dbo].[PedidosVentaOFFLINETMP]
                                           ([pvFecha]
                                           ,[pvFechaPrevista]
                                           ,[pvFechaReal]
                                           ,[pvCodClienteSolic]
                                           ,[pvCodClienteDest]
                                           ,[pvMaterial]
                                           ,[pvGrupo]
                                           ,[pvKilos]
                                           ,[pvUnidad]
                                           ,[pvKilosPtes]
                                           ,[pvUnidadesPtes]
                                           ,[pvKgPtes]
                                           ,[pvKilosEnt]                                           
                                           ,[pvLineaPedido]
                                           ,[pvClaseEntrega]
                                           ,[pvTipoPosicion]
                                           ,[pvTipoEnvio]
                                           ,[pvCentro]
                                           ,[pvAlmacen]
                                           ,[pvPedido]
                                           ,[pvOrdenTransporte]
                                           ,[pvEstadoOrdenTpte]
                                           ,[pvEntregaPendiente]
                                           ,[pvNombreMaterial]
                                           ,[pvNombreCliente]
                                           ,[pvStockActual]
                                           ,[pvNuevoStockActual]
                                           ,[pvNuevoStockAPedidoVenta]
                                           ,[pvStatusGLobal]
                                           ,[pvStatusEntrega]
                                           ,[pvNombrePuestoTrabajo]
                                           ,[pvFechaPlan],[pvCodigoMaterialFab])
                                     VALUES ('" &
                                            miPedido.Fecha & "','" &
                                            miPedido.FechaPrevista & "','" &
                                             miPedido.FechaReal & "'," &
                                            miPedido.CodClienteSolic & "," &
                                            miPedido.CodClienteDest & ",'" &
                                            miPedido.Material & "','" &
                                            miPedido.Grupo & "'," &
                                            miPedido.Kilos & ",'" &
                                            miPedido.Unidad & "'," &
                                            miPedido.KilosPtes & "," &
                                            miPedido.UnidadesPtes & ",'" &
                                            miPedido.KgPtes & "'," &
                                            miPedido.KilosEnt & ",'" &
                                            miPedido.LineaPedido & "','" &
                                            miPedido.ClaseEntrega & "','" &
                                            miPedido.TipoPosicion & "','" &
                                            miPedido.TipoEnvio & "','" &
                                            miPedido.Centro & "','" &
                                            miPedido.Almacen & "','" &
                                            miPedido.Pedido & "','" &
                                            miPedido.OrdenTransporte & "','" &
                                            miPedido.EstadoOrdenTpte & "','" &
                                            miPedido.EntregaPendiente & "','" &
                                            miPedido.NombreMaterial & "','" &
                                            miPedido.NombreCliente & "'," &
                                            miPedido.StockActual & "," &
                                            miPedido.NuevoStockActual & "," &
                                             miPedido.NuevoStockAPedidoVenta & ",'" &
                                            miPedido.StatusGLobal & "','" &
                                            miPedido.StatusEntrega & "','" &
                                            miPedido.NombrePuestoTrabajo & "','" &
                                            miPedido.FechaPlan & "','" &
                                            miPedido.CodigoMaterialFab & "')"


                insertar = Datos.CGPL.EjecutarConsulta(sSql)

                'If Insertar Then
                '    Me.bCreado = insertar
                '    'Datos.GuardarLog(TipoLogDescripcion.Alta & " Materiales ", Me.sCodigo)
                'End If
            Next

            Return insertar

        Catch ex As Exception
            insertar = False
            Throw New Exception(ex.Message & "() ", ex)
        End Try
    End Function

    Public Function InsertarPedidos(ByVal ListaPedidos As List(Of PedidosVenta)) As Boolean
        Dim insertar As Boolean = True
        Try

            'For Each miPedido In ListaPedidos
            Dim sSql As String = "insert into PedidosVentaOFFLINE select * from PedidosVentaOFFLINETMP"


            insertar = Datos.CGPL.EjecutarConsulta(sSql)

            'If Insertar Then
            '    Me.bCreado = insertar
            '    'Datos.GuardarLog(TipoLogDescripcion.Alta & " Materiales ", Me.sCodigo)
            'End If
            'Next

            Return insertar

        Catch ex As Exception
            insertar = False
            Throw New Exception(ex.Message & "() ", ex)
        End Try
    End Function

    Public Function EliminarForecastTemporal() As Boolean
        Try

            Dim sSql As String = " DELETE FROM FabricacionesContraStockOFFLINETMP "

            EliminarForecastTemporal = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarForecastTemporal = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function InsertarForecastTMP(ByVal ListaForecast As List(Of BeanPullSystem)) As Boolean
        Dim insertar As Boolean = True
        Try

            For Each miForecast In ListaForecast
                Dim sSql As String = "INSERT INTO [dbo].[FabricacionesContraStockOFFLINETMP]
                                       ([Codigo]
                                       ,[Nombre]
                                       ,[Formato]
                                       ,[GrupoPlanif]
                                       ,[Cantidad]
                                       ,[DiasPP]                                                                      
                                       ,[StockActual]
                                       ,[FabricacionesPendientes]
                                       ,[SituacionActual]
                                       ,[UnidadesFabricar]
                                       ,[KgNuevaFabricacion]
                                       ,[ValorRdoTanque]
                                       ,[NuevaFabricacion]
                                       ,[FechaFin]
                                       ,[ValorRedondeo]
                                       ,[MesPS]
                                       ,[AnioPS]                                                                      
                                       ,[CodPuestoTrabajo]
                                       ,[NombrePuestoTrabajo]                                      
                                       ,[CantidadBaseFormula]
                                       ,[CantidadBaseMaterias]                                     
                                       ,[FechaFinFanPendientes]
                                       ,[CantidadPendientePedVentas]
                                       ,[Creado]
                                       ,[PuestoTrabajo]
                                       ,[FechaRotura])
                                 VALUES
                                       ('" &
                                        miForecast.Codigo & "','" &
                                        miForecast.Nombre & "','" &
                                        miForecast.Formato & "','" &
                                        miForecast.GrupoPlanif & "'," &
                                       miForecast.Cantidad & "," &
                                       miForecast.DiasPP & "," &
                                        miForecast.StockActual & "," &
                                       miForecast.FabricacionPendiente & "," &
                                        miForecast.SituacionActual & "," &
                                       miForecast.UnidadesFabricar & "," &
                                       miForecast.KgNuevaFabricacion & "," &
                                       miForecast.ValorRdoTanque & "," &
                                       miForecast.NuevaFabricacion & ",'" &
                                       miForecast.FechaFin & "'," &
                                       miForecast.ValorRedondeo & "," &
                                       miForecast.MesPS & "," &
                                       miForecast.AnioPS & ",'" &
                                       miForecast.CodPuestoTrabajo & "','" &
                                       miForecast.NombrePuestoTrabajo & "'," &
                                       miForecast.CantidadBaseFormula & "," &
                                       miForecast.CantidadBaseMaterias & ",'" &
                                       miForecast.FechaFinFanPendientes & "'," &
                                       miForecast.CantidadPendientePedVentas & ",'" &
                                       miForecast.Creado & "','" &
                                       miForecast.PuestoTrabajo & "','" &
                                       miForecast.FechaRotura & "'" & ")"


                insertar = Datos.CGPL.EjecutarConsulta(sSql)

                'If Insertar Then
                '    Me.bCreado = insertar
                '    'Datos.GuardarLog(TipoLogDescripcion.Alta & " Materiales ", Me.sCodigo)
                'End If
            Next

            Return insertar

        Catch ex As Exception
            insertar = False
            Throw New Exception(ex.Message & "() ", ex)
        End Try
    End Function

    Public Function EliminarForecast() As Boolean
        Try
            Dim sSql As String = " DELETE FROM FabricacionesContraStockOFFLINE " '&
            '" WHERE pvFechaPrevista >= '" & FechaInicio & "' and pvFechaPrevista <='" & FechaFinal & "'"

            EliminarForecast = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarForecast = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function InsertarForecast() As Boolean
        Dim insertar As Boolean = True
        Try

            'For Each miPedido In ListaPedidos
            Dim sSql As String = "insert into FabricacionesContraStockOFFLINE select * from FabricacionesContraStockOFFLINETMP"


            insertar = Datos.CGPL.EjecutarConsulta(sSql)

            'If Insertar Then
            '    Me.bCreado = insertar
            '    'Datos.GuardarLog(TipoLogDescripcion.Alta & " Materiales ", Me.sCodigo)
            'End If
            'Next

            Return insertar

        Catch ex As Exception
            insertar = False
            Throw New Exception(ex.Message & "() ", ex)
        End Try
    End Function

    Public Function EliminarPullSystemTemporal() As Boolean
        Try

            Dim sSql As String = " DELETE FROM PullSystemOFFLINETMP "

            EliminarPullSystemTemporal = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarPullSystemTemporal = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function EliminarPullSystemTemporalAgrupado() As Boolean
        Try

            Dim sSql As String = " DELETE FROM PullSystemOFFLINEAgrupadoTMP "

            EliminarPullSystemTemporalAgrupado = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarPullSystemTemporalAgrupado = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function InsertarPullSystemTMP(ByVal ListaForecast As List(Of PullSystem), ByVal DiasControl As Integer, ByVal AñoActualizacion As Integer, ByVal MesActualizacion As Integer, ByVal Version As String) As Boolean
        Dim insertar As Boolean = True
        Try

            For Each miForecast In ListaForecast
                Dim sSql As String = "INSERT INTO [dbo].[PullSystemOFFLINETMP]
                                               ([CodigoMaterial]
                                               ,[Mes]
                                               ,[Año]
                                               ,[Cantidad]
                                               ,[diasControl]
                                               ,[StockActual]
                                               ,[FechaRotura]
                                               ,[FechaRoturaEntradas]
                                               ,[AñoActualizacion]
                                               ,[MesActualizacion]
                                               ,[StockBloqueado]
                                               ,[Estatus]
                                               ,[FechaCorta]
                                               ,[Necesidad]
                                               ,[Version])    
                                 VALUES
                                       ('" &
                                        miForecast.CodigoMaterial & "'," &
                                        miForecast.Mes & "," &
                                        miForecast.Año & "," &
                                        miForecast.Cantidad & "," &
                                        DiasControl & "," &
                                        miForecast.Stock & ",'" &
                                        miForecast.FechaRotura & "','" &
                                        miForecast.FechaRoturaEntradas & "'," &
                                        AñoActualizacion & "," &
                                        MesActualizacion & ", " &
                                        miForecast.StockBloqueado & ",'" &
                                        miForecast.Estatus & "','" &
                                        miForecast.FechaCorta & "','" &
                                        miForecast.Necesidad & "','" &
                                        Version & "')"


                insertar = Datos.CGPL.EjecutarConsulta(sSql)

                'If Insertar Then
                '    Me.bCreado = insertar
                '    'Datos.GuardarLog(TipoLogDescripcion.Alta & " Materiales ", Me.sCodigo)
                'End If
            Next

            Return insertar

        Catch ex As Exception
            insertar = False
            Throw New Exception(ex.Message & "() ", ex)
        End Try
    End Function

    Public Function InsertarPullSystemTMPAgrupado(ByVal ListaForecast As List(Of BeanPullSystem), ByVal DiasControl As Integer, ByVal AñoActualizacion As Integer, ByVal MesActualizacion As Integer, ByVal Version As String) As Boolean
        Dim insertar As Boolean = True
        Try

            For Each miForecast In ListaForecast
                Dim sSql As String = "INSERT INTO [dbo].[PullSystemOFFLINEAgrupadoTMP]
                                       ([Codigo]
                                       ,[Nombre]
                                       ,[NombreV2]
                                       ,[Formato]
                                       ,[GrupoPlanif]
                                       ,[Cantidad]
                                       ,[DiasPP]
                                       ,[CalculoStockMinimo]
                                       ,[StockMaxPS]
                                       ,[StockActual]
                                       ,[FabricacionPendiente]
                                       ,[FechaFinFanPendientes]
                                       ,[SituacionActual]
                                       ,[UnidadesFabricar]
                                       ,[KgNuevaFabricacion]
                                       ,[ValorRdoTanque]
                                       ,[NuevaFabricacion]
                                       ,[FechaFin]
                                       ,[ValorRedondeo]
                                       ,[MesPS]
                                       ,[AnioPS]
                                       ,[FechaRotura]
                                       ,[StockBloqueado]
                                       ,[Estatus]
                                       ,[FechaPS]
                                       ,[StockCritico]
                                       ,[CodPuestoTrabajo]
                                       ,[NombrePuestoTrabajo]
                                       ,[PuestoTrabajo]
                                       ,[CodigoMaterialFab]
                                       ,[NombresPuestoTrabajo]
                                       ,[CantidadPendientePedVentas]
                                       ,[CantidadBaseFormula]
                                       ,[CantidadBaseMaterias]
                                       ,[Necesidad]
                                       ,[FechaCorta]
                                       ,[AñoActualizacion]
                                       ,[MesActualizacion]
                                       ,[Version])  
                                 VALUES
                                   ('" &
                                        miForecast.Codigo & "','" &
                                        miForecast.Nombre & "','" &
                                   miForecast.NombreV2 & "','" &
                                   miForecast.Formato & "','" &
                                   miForecast.GrupoPlanif & "'," &
                                   miForecast.Cantidad & "," &
                                   miForecast.DiasPP & "," &
                                   miForecast.CalculoStockMinimo & "," &
                                   miForecast.StockMaxPS & "," &
                                   miForecast.StockActual & "," &
                                   miForecast.FabricacionPendiente & ",'" &
                                   miForecast.FechaFinFanPendientes & "'," &
                                   miForecast.SituacionActual & "," &
                                   miForecast.UnidadesFabricar & "," &
                                   miForecast.KgNuevaFabricacion & "," &
                                   miForecast.ValorRdoTanque & "," &
                                   miForecast.NuevaFabricacion & ",'" &
                                   miForecast.FechaFin & "'," &
                                   miForecast.ValorRedondeo & "," &
                                   miForecast.MesPS & "," &
                                   miForecast.AnioPS & ",'" &
                                   miForecast.FechaRotura & "'," &
                                   miForecast.StockBloqueado & ",'" &
                                   miForecast.Estatus & "','" &
                                   miForecast.FechaPS & "','" &
                                   miForecast.StockCritico & "'," &
                                   miForecast.CodPuestoTrabajo & ",'" &
                                   miForecast.NombrePuestoTrabajo & "','" &
                                   miForecast.PuestoTrabajo & "','" &
                                   miForecast.CodigoMaterialFab & "','" &
                                   DamePuestosConcatenados(miForecast.NombresPuestoTrabajo) & "'," &
                                   miForecast.CantidadPendientePedVentas & "," &
                                   miForecast.CantidadBaseFormula & "," &
                                   miForecast.CantidadBaseMaterias & ",'" &
                                   miForecast.Necesidad & "','" &
                                   miForecast.FechaCorta & "'," &
                                   AñoActualizacion & "," &
                                   MesActualizacion & ",'" &
                                   Version & "')"


                insertar = Datos.CGPL.EjecutarConsulta(sSql)

            Next

            Return insertar

        Catch ex As Exception
            insertar = False
            Throw New Exception(ex.Message & "() ", ex)
        End Try
    End Function

    Private Function DamePuestosConcatenados(nombresPuestoTrabajo As List(Of PuestosTrabajo)) As String
        Dim PuestosConcatenados = ""
        Try
            Dim PuestosItems = nombresPuestoTrabajo.Select(Function(s) s.CodigoPuestoTrabajo.ToString()).Distinct().ToArray()
            PuestosConcatenados = Join(PuestosItems, "|")
            Return PuestosConcatenados
        Catch ex As Exception
            Return PuestosConcatenados
        End Try
    End Function

    Public Function Dame_Minutos_Tiempo_Preparacion(Hoja_de_Ruta As HojaRuta,
                                               Puestotrabajo As Integer,
                                               IncluirPreparacion As Boolean, GrupoHojaRuta As String, ContadorHojaRuta As String) As Integer
        Try
            Dim iMinutos As Integer = 0

            For Each miOper In Hoja_de_Ruta.OperacHojaRutaLista
                If Puestotrabajo = miOper.CodigoPuestoDeTrabajo AndAlso miOper.CantidadBase <> 0 AndAlso miOper.GrupoSAP.Trim = GrupoHojaRuta AndAlso miOper.ContGrupoSAP.Trim = ContadorHojaRuta Then
                    'iMinutos += CInt(miOper.MinutosLimpieza) +
                    iMinutos += CInt(IIf(IncluirPreparacion = True, miOper.MinutosPreparacion, 0))
                End If
            Next

            Return iMinutos

        Catch ex As Exception
            Dame_Minutos_Tiempo_Preparacion = 0
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try


    End Function

    Public Function Dame_Minutos_Tiempo_FabricacionEnMarcha(Cantidad As Integer,
                                               Hoja_de_Ruta As HojaRuta,
                                               Puestotrabajo As Integer,
                                               IncluirPreparacion As Boolean, GrupoHojaRuta As String, ContadorHojaRuta As String) As Integer
        Try
            Dim iMinutos As Integer = 0
            If Cantidad > 0 Then
                For Each miOper In Hoja_de_Ruta.OperacHojaRutaLista
                    If Puestotrabajo = miOper.CodigoPuestoDeTrabajo AndAlso miOper.CantidadBase <> 0 AndAlso miOper.GrupoSAP.Trim = GrupoHojaRuta AndAlso miOper.ContGrupoSAP.Trim = ContadorHojaRuta Then
                        iMinutos += CInt(Cantidad * miOper.MinutosMaquina / miOper.CantidadBase)
                    End If
                Next
            Else
                iMinutos = 0
            End If


            Return iMinutos

        Catch ex As Exception
            Dame_Minutos_Tiempo_FabricacionEnMarcha = 0
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try


    End Function

    Public Function DamePuestosTrabajo(ByVal Optional TipoPuestoTrabajo As String = "", Optional ByVal Recursos As Boolean = True,
                                       Optional ByVal activo As Boolean = True, Optional ByVal todo As Boolean = False,
                                       Optional ByVal sCodigoCentroProd As String = "") As List(Of PuestosTrabajo)
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = " SELECT DISTINCT puestostrabajo.* " &
                                 " FROM puestostrabajo WITH(NOLOCK) "
            '" WHERE ptCod IN " &
            '"(SELECT DISTINCT opPuestoTrabajo " &
            '" FROM OperacionesHojaRuta with(nolock) "


            If TipoPuestoTrabajo.Trim.Length <> 0 Then
                sSql &= " WHERE ptTipo in(" & TipoPuestoTrabajo & ")"
            End If

            If Recursos = False Then
                sSql &= " AND ptRecurso =0"
            End If

            If Not String.IsNullOrEmpty(sCodigoCentroProd) Then
                sSql &= " AND ptCentroProd =" & sCodigoCentroProd
            End If

            If todo = False Then
                If activo = False Then
                    sSql &= " AND  ptActivo =0 ORDER BY PTORDEN ASC"
                Else
                    sSql &= " AND  ptActivo =1 ORDER BY PTORDEN ASC"
                End If
            End If


            DamePuestosTrabajo = New List(Of PuestosTrabajo)

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DamePuestosTrabajo.Add(New PuestosTrabajo(_CodigoPuestoTrabajo:=CInt(miRegistro.Item("ptCod")),
                                                                      _Nombre:=CStr(miRegistro.Item("ptNombre")),
                                                                      _Centro:=CStr(miRegistro.Item("ptCentro")),
                                                                      _Operarios:=CInt(miRegistro.Item("ptOperarios")),
                                                                      _Tipo:=CStr(miRegistro.Item("ptTipo")),
                                                                      _AreaProduccion:=CStr(miRegistro.Item("ptAreaProd")),
                                                                      _Activo:=CBool(NoNull(miRegistro.Item("ptActivo"), "A")),
                                                                       _Recurso:=CBool(NoNull(miRegistro.Item("ptRecurso"), "N")),
                                                                      _CambioPedido:=CBool(miRegistro.Item("ptCambiarPedido")),
                                                                      _Orden:=CInt(miRegistro.Item("ptOrden")),
                                                                      _VelocidadMax:=CInt(miRegistro.Item("ptVelMax")),
                                                                      _VelocidadActual:=CInt(miRegistro.Item("ptVelActual")),
                                                                      _CodCentroProd:=CInt(miRegistro.Item("ptCentroProd")),
                                                                      _CodProveedor:=CStr(miRegistro.Item("ptProveedor")),
                                                                      _EsCentroExterno:=CBool(miRegistro.Item("ptEsCoperativaExt")),
                                                                      _PuedeRealizarTraspasos:=CBool(miRegistro.Item("ptPuedeRealizarTraspaso"))))
                Next
            End If

        Catch ex As Exception
            DamePuestosTrabajo = New List(Of PuestosTrabajo)
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function


    Public Function EliminarPullSystem(ByVal Año As Integer, ByVal Mes As Integer, ByVal Version As String) As Boolean
        Try
            Dim sSql As String = " DELETE FROM PullSystemOFFLINE " &
            " WHERE AñoActualizacion = " & Año & " and MesActualizacion = " & Mes & " and Version = '" & Version & "'"

            EliminarPullSystem = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarPullSystem = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function EliminarPullSystemAgrupado(ByVal Año As Integer, ByVal Mes As Integer, ByVal Version As String) As Boolean
        Try
            Dim sSql As String = " DELETE FROM PullSystemOFFLINEAgrupado " &
            " WHERE AñoActualizacion = " & Año & " and MesActualizacion = " & Mes & " and Version = '" & Version & "'"

            EliminarPullSystemAgrupado = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarPullSystemAgrupado = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function EliminarPullSystemPasados(ByVal Version As String) As Boolean

        Try
            Dim sSql As String = " DELETE FROM PullSystemOFFLINE " &
            " WHERE Version LIKE '%" & Version & "%' or Version is null"

            EliminarPullSystemPasados = Datos.CGPL.EjecutarConsulta(sSql)

            sSql = " DELETE FROM PullSystemOFFLINEAgrupado " &
            " WHERE  Version LIKE '%" & Version & "%' or Version is null"

            EliminarPullSystemPasados = Datos.CGPL.EjecutarConsulta(sSql)

            'If EliminarPedidos Then
            '    Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Pedidos ", Me.sCodigo)
            'End If
        Catch ex As Exception
            EliminarPullSystemPasados = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function
    Public Function InsertarPullSystem() As Boolean
        Dim insertar As Boolean = True
        Try

            'For Each miPedido In ListaPedidos
            Dim sSql As String = "insert into PullSystemOFFLINE select * from PullSystemOFFLINETMP"


            insertar = Datos.CGPL.EjecutarConsulta(sSql)

            'If Insertar Then
            '    Me.bCreado = insertar
            '    'Datos.GuardarLog(TipoLogDescripcion.Alta & " Materiales ", Me.sCodigo)
            'End If
            'Next

            Return insertar

        Catch ex As Exception
            insertar = False
            Throw New Exception(ex.Message & "() ", ex)
        End Try
    End Function

    Public Function InsertarPullSystemAgrupado() As Boolean
        Dim insertar As Boolean = True
        Try

            'For Each miPedido In ListaPedidos
            Dim sSql As String = "insert into PullSystemOFFLINEAgrupado select * from PullSystemOFFLINEAgrupadoTMP"


            insertar = Datos.CGPL.EjecutarConsulta(sSql)

            'If Insertar Then
            '    Me.bCreado = insertar
            '    'Datos.GuardarLog(TipoLogDescripcion.Alta & " Materiales ", Me.sCodigo)
            'End If
            'Next

            Return insertar

        Catch ex As Exception
            insertar = False
            Throw New Exception(ex.Message & "() ", ex)
        End Try
    End Function

    Public Function ValidacionPullsystem(ByVal codMaterial As String, ByVal FechaPrevistaFin As Date) As Boolean
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = "SELECT * " &
                                 "From Materiales ma " &
                                 "WHERE ma.maCod ='" & codMaterial & "' and ma.maFecIniPS <>'1900-01-01' and ma.maFecFinPS <> '1900-01-01' and " &
                                 "not '" & FechaPrevistaFin & "'>= ma.maFecIniPS and  not '" & FechaPrevistaFin & "' <= ma.maFecFinPS "

            Dim ResultadoValidacion As Boolean = False
            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                If DTDatos.Rows.Count > 0 Then
                    ResultadoValidacion = True
                Else
                    ResultadoValidacion = False
                End If

            End If
            Return ResultadoValidacion
        Catch ex As Exception
            ValidacionPullsystem = False
            Throw New Exception(ex.Message & " - -" & "()", ex)
        End Try
    End Function


    'Public Function RegistrarEjecucion()
    '    Try

    '        Using DB = New SqlConnection(cnn)
    '            DB.Execute(sql:="sp_RegistraEjecucionTarea", param:=New With {Key
    '                .tlIdTarea = "CargaInfoGPS", Key
    '                .tlPeriodo = 6, Key
    '                .tlUltimaEjecucion = DateTime.Now
    '            }, commandType:=CommandType.StoredProcedure)
    '        End Using

    '    Catch ex As Exception
    '        Throw New Exception("Error al ejecutar sp_RegistraEjecucionTarea, detalle: " & vbLf & ex.Message, ex)
    '    End Try
    'End Function

    'Public Function RegistrarEjecucion(ByVal [Error] As String)
    '    Try

    '        Using DB = New SqlConnection(cnn)
    '            DB.Execute(sql:="sp_RegistraEjecucionTarea", param:=New With {Key
    '                .tlIdTarea = "CargaInfoGPS", Key
    '                .tlPeriodo = 6, Key
    '                .tlUltimaEjecucion = DateTime.Now, Key
    '                .tlError = [Error]
    '            }, commandType:=CommandType.StoredProcedure)
    '        'Using (Dim DB = New SqlConnection(cnn))
    '        '    {
    '        '        DB.Execute(Sql:  "sp_RegistraEjecucionTarea", param: New { tlIdTarea = "CargaInfoGPS", tlPeriodo = 6, tlUltimaEjecucion = DateTime.Now, tlError = Error },
    '        '                        commandType: CommandType.StoredProcedure);
    '        '    }
    '        End Using

    '    Catch ex As Exception
    '        Throw New Exception("Error al ejecutar sp_RegistraEjecucionTarea, detalle: " & vbLf & ex.Message, ex)
    '    End Try
    'End Function

End Module
