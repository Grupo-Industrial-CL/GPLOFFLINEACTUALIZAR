
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Module DatosPullSystem


    Public Function DamelistaDiasLaborales() As List(Of DiasLaborales)

        Try

            DamelistaDiasLaborales = New List(Of DiasLaborales)

            For iMes = 1 To 12

                DamelistaDiasLaborales.Add(New DiasLaborales(iMes))

            Next

        Catch ex As Exception

            DamelistaDiasLaborales = New List(Of DiasLaborales)
            Throw New Exception(ex.Message & " - -" & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try


    End Function

    Public Function DameAnalisisCapacidad(ByVal EstadosFabricacion As String, ByVal TipoPuestoTrabajo As String) As DataTable
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = "SELECT " &
                                 "FB.opFechaPrevFin as FechaPreFin, " &
                                "PT.ptNombre as LineaPT , " &
                                "sum(Case  " &
                                "	when  FB.opCantidadPlanif-FB.opCantidadFab >0 and  OHR.opCantidadBase <>0  " &
                                "         then (((FB.opCantidadPlanif-FB.opCantidadFab)* OHR.opMinutos)/OHR.opCantidadBase)+ " &
                                "		  OHR.opMinutosPrep+OHR.opMinutosLimpieza   " &
                                "	else 0  " &
                                "end )as HorasMaquina,  " &
                                "OHR.opOperarios  as Operarios  " &
                                "FROM Fabricaciones FB    " &
                                "INNER JOIN OperacionesHojaRuta OHR ON FB.opGrupoHR=OHR.opGrupoSAP And FB.opContHR=OHR.opContGrupoSAP  " &
                                "INNER JOIN PuestosTrabajo PT ON PT.ptCod=OHR.opPuestoTrabajo AND PT.ptTipo=" & TipoPuestoTrabajo &
                                "WHERE  opEnmarcha IN (" & EstadosFabricacion & ")   " &
                                "Group by FB.opFechaPrevFin,PT.ptNombre,OHR.opOperarios   " &
                                "order by  FB.opFechaPrevFin asc  "

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Return DTDatos
            Else
                Return New DataTable
            End If
        Catch ex As Exception
            Return New DataTable
            Throw New Exception(ex.Message & " - -" & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function DameFabricacionesAgrupadas(ByVal FechaIni_FinFab As Date,
                                      ByVal FechaFin_FinFab As Date,
                                      ByVal EstadosFabricacion As String,
                                      ByVal OrdenarResultado As Boolean) As List(Of Fabricaciones)
        Try
            Dim DTDatos As New DataTable
            Dim sWhere As String = ""
            Dim sSql As String = ""

            sSql = "SELECT * FROM Fabricaciones "


            sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opFechaFin BETWEEN '" & FechaIni_FinFab & "' AND '" & FechaFin_FinFab & "'"

            If EstadosFabricacion <> "" Then
                sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opEnmarcha IN (" & EstadosFabricacion & ") "
            Else
                sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opEnmarcha NOT IN (" & EstadoFabricacion.Anulada & ") "
            End If

            If sWhere <> "" Then
                sSql &= " WHERE " & sWhere
            End If

            If OrdenarResultado = True Then
                sSql &= " ORDER BY openMarcha desc,opFechaIni,opOrdenMaq"
            End If

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                DameFabricacionesAgrupadas = (From elemento In DTDatos
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
                DameFabricacionesAgrupadas = New List(Of Fabricaciones)
            End If

        Catch ex As Exception
            DameFabricacionesAgrupadas = New List(Of Fabricaciones)
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function DameFabricaciones(Optional ByVal EstadosFabricacion As String = "",
                                     Optional ByVal FechaInicio As Date = FechaGlobal,
                                     Optional ByVal FechaFin As Date = FechaGlobal,
                                     Optional ByVal CodMaterial As String = "",
                                     Optional ByVal FechaPrevistaFIN As Date = FechaGlobal,
                                     Optional ByVal FechaTurno As Date = FechaGlobal,
                                     Optional ByVal FechaTurnoFin As Date = FechaGlobal,
                                     Optional ByVal OrdenarOrdenFabrcacion As Boolean = False,
                                     Optional ByVal PuestoTrabajo As Integer = 0,
                                     Optional ByVal bTopOne As Boolean = False,
                                     Optional ByVal Turno As String = "",
                                     Optional ByVal bControlProduccion As Boolean = False,
                                     Optional ByVal bPlanifFuturo As Boolean = False,
                                     Optional ByVal codigoEquipo As Integer = 0,
                                     Optional ByVal bResumenFabricacion As Boolean = False,
                                      Optional ByVal iOrdenEnvasado As Integer = 0,
                                      Optional ByVal VariosPuestoTrabajo As String = "") As List(Of Fabricaciones)
        Try
            Dim DTDatos As New DataTable
            Dim sWhere As String = ""
            Dim sSql As String = ""
            If bTopOne = True Then
                sSql = "SELECT TOP 1 * " &
                                 "FROM Fabricaciones "
            Else
                sSql = "SELECT * " &
                                 "FROM Fabricaciones "
            End If

            If bResumenFabricacion = True Then

                sWhere &= " Convert(Date, opTFecIni,120)  between '" & FechaInicio & "' AND  '" & FechaFin & "' AND  opPuestoTrabajo = " & PuestoTrabajo
                sWhere &= " AND  opEnmarcha IN (" & EstadosFabricacion & ") "

            Else
                If FechaInicio <> FechaGlobal Then
                    sWhere &= " opFechaIni >= '" & FechaInicio & "' "
                End If

                If FechaPrevistaFIN <> FechaGlobal Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opFechaPrevFin <= '" & FechaPrevistaFIN & "' "
                End If

                If FechaFin <> FechaGlobal Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opFechaFin <= '" & FechaFin & "' "
                End If

                If PuestoTrabajo <> 0 Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opPuestoTrabajo = " & PuestoTrabajo
                End If

                If VariosPuestoTrabajo <> "" Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opPuestoTrabajo IN (" & VariosPuestoTrabajo & ")"
                End If

                If Turno <> "" Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opTurno = '" & Turno & "' "
                End If

                If FechaTurno <> FechaGlobal AndAlso
                   FechaTurnoFin <> FechaGlobal Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opTFecIni BETWEEN '" & FechaTurno & "' AND '" & FechaTurnoFin & "'"
                ElseIf FechaTurno <> FechaGlobal Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opTFecIni = '" & FechaTurno & "' "
                End If

                If CodMaterial <> "" Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) &
                                " opMaterialPadre = '" & Trim(CodMaterial) & "' "
                End If

                If iOrdenEnvasado <> 0 Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opOrdenEnvSAP = " & iOrdenEnvasado
                End If

                If bControlProduccion = True Then
                    'sWhere &= " AND (opFechaFin=CONVERT(DATE,GETDATE(),120) OR opFechaFin='1900-01-01') "
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & "  (opEnMarcha IN (" & EstadoFabricacion.EnMarcha & "," & EstadoFabricacion.PteFabricar & ") OR " &
                                       "(opEnMarcha = " & EstadoFabricacion.Finalizada & " AND CONVERT(DATE, opFechaFin) ='" & Now.Date.ToString("yyyy-MM-dd") & "'))"
                ElseIf bPlanifFuturo = True Then
                    'sWhere &= " AND (opFechaFin=CONVERT(DATE,GETDATE(),120) OR opFechaFin='1900-01-01') "
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & "  (opEnMarcha IN (" & EstadoFabricacion.EnMarcha & "," & EstadoFabricacion.PteFabricar & "," & EstadoFabricacion.PlanFuturo & ") OR " &
                                       "(opEnMarcha = " & EstadoFabricacion.Finalizada & " AND CONVERT(DATE, opFechaFin) ='" & Now.Date.ToString("yyyy-MM-dd") & "'))"
                ElseIf EstadosFabricacion <> "" Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opEnmarcha IN (" & EstadosFabricacion & ") "
                Else
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opEnmarcha NOT IN (" & EstadoFabricacion.Anulada & ") "
                End If

                If codigoEquipo > 0 Then
                    sWhere &= "and opEquipo = " & codigoEquipo & " "
                End If
            End If

            If sWhere <> "" Then
                sSql &= " WHERE " & sWhere
            End If

            If OrdenarOrdenFabrcacion = True Then
                sSql &= " ORDER BY openMarcha desc,opFechaIni,opOrdenMaq"
            Else
                sSql &= " ORDER BY opPuestoTrabajo, opFechaIni"
            End If

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then





                DameFabricaciones = (From elemento In DTDatos
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
                                                            CInt(NoNull(elemento.Item("opEquipo"), "D")),
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
                                                            CInt(NoNull(elemento.Item("opIdEnvio"), "D")),
                                                            CStr(NoNull(elemento.Item("opEsLiberada"), "A")),
                                                            CStr(NoNull(elemento.Item("opMensajeError"), "A")),
                                                            CStr(NoNull(elemento.Item("opTieneError"), "A")),
                                                            CStr(NoNull(elemento.Item("opFaltanteOrdenEnvasadoSAP"), "A")),
                                                            CStr(NoNull(elemento.Item("opFaltanteOrdenFabricacionSAP"), "A")),
                                                            CStr(NoNull(elemento.Item("opComentarioUsuario"), "A")),
                                                            CStr(NoNull(elemento.Item("opProcedencia"), "A")),
                                                            CStr(NoNull(elemento.Item("opCodGranel"), "A")),
                                                            CDate(NoNull(elemento.Item("opFechaFuturoIni"), "DT")),
                                                            CDate(NoNull(elemento.Item("opFechaFuturoFin"), "DT")),
                                                            CStr(NoNull(elemento.Item("opEsLiberadaEnvasado"), "A")),
                                                            CInt(NoNull(elemento.Item("opUnidadesPorCaja"), "D")))).ToList
            Else
                DameFabricaciones = New List(Of Fabricaciones)
            End If
        Catch ex As Exception
            DameFabricaciones = New List(Of Fabricaciones)
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function



    Public Function DameFabricacionesOld(Optional ByVal EstadosFabricacion As String = "",
                                     Optional ByVal FechaInicio As Date = FechaGlobal,
                                     Optional ByVal FechaFin As Date = FechaGlobal,
                                     Optional ByVal CodMaterial As String = "",
                                     Optional ByVal FechaPrevistaFIN As Date = FechaGlobal,
                                     Optional ByVal FechaTurno As Date = FechaGlobal,
                                     Optional ByVal FechaTurnoFin As Date = FechaGlobal,
                                     Optional ByVal OrdenarOrdenFabrcacion As Boolean = False,
                                     Optional ByVal PuestoTrabajo As Integer = 0,
                                     Optional ByVal bTopOne As Boolean = False,
                                     Optional ByVal Turno As String = "",
                                     Optional ByVal bControlProduccion As Boolean = False,
                                     Optional ByVal codigoEquipo As Integer = 0,
                                     Optional ByVal bResumenFabricacion As Boolean = False,
                                      Optional ByVal iOrdenEnvasado As Integer = 0,
                                      Optional ByVal VariosPuestoTrabajo As String = "") As List(Of Fabricaciones)
        Try
            Dim DTDatos As New DataTable
            Dim sWhere As String = ""
            Dim sSql As String = ""
            If bTopOne = True Then
                sSql = "SELECT TOP 1 * " &
                                 "FROM Fabricaciones "
            Else
                sSql = "SELECT * " &
                                 "FROM Fabricaciones "
            End If

            If bResumenFabricacion = True Then

                sWhere &= " Convert(Date, opTFecIni,120)  between '" & FechaInicio & "' AND  '" & FechaFin & "' AND  opPuestoTrabajo = " & PuestoTrabajo
                sWhere &= " AND  opEnmarcha IN (" & EstadosFabricacion & ") "

            Else
                If FechaInicio <> FechaGlobal Then
                    sWhere &= " opFechaIni >= '" & FechaInicio & "' "
                End If

                If FechaPrevistaFIN <> FechaGlobal Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opFechaPrevFin <= '" & FechaPrevistaFIN & "' "
                End If

                If FechaFin <> FechaGlobal Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opFechaFin <= '" & FechaFin & "' "
                End If

                If PuestoTrabajo <> 0 Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opPuestoTrabajo = " & PuestoTrabajo
                End If

                If VariosPuestoTrabajo <> "" Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opPuestoTrabajo IN (" & VariosPuestoTrabajo & ")"
                End If

                If Turno <> "" Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opTurno = '" & Turno & "' "
                End If

                If FechaTurno <> FechaGlobal AndAlso
                   FechaTurnoFin <> FechaGlobal Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opTFecIni BETWEEN '" & FechaTurno & "' AND '" & FechaTurnoFin & "'"
                ElseIf FechaTurno <> FechaGlobal Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opTFecIni = '" & FechaTurno & "' "
                End If

                If CodMaterial <> "" Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) &
                                " opMaterialPadre = '" & Trim(CodMaterial) & "' "
                End If

                If iOrdenEnvasado <> 0 Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opOrdenEnvSAP = " & iOrdenEnvasado
                End If

                If bControlProduccion = True Then
                    'sWhere &= " AND (opFechaFin=CONVERT(DATE,GETDATE(),120) OR opFechaFin='1900-01-01') "
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & "  (opEnMarcha IN (" & EstadoFabricacion.EnMarcha & "," & EstadoFabricacion.PteFabricar & ") OR " &
                                       "(opEnMarcha = " & EstadoFabricacion.Finalizada & " AND CONVERT(DATE, opFechaFin) ='" & Now.Date.ToString("yyyy-MM-dd") & "'))"
                ElseIf EstadosFabricacion <> "" Then
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opEnmarcha IN (" & EstadosFabricacion & ") "
                Else
                    sWhere &= CStr(IIf(sWhere <> "", " AND ", "")) & " opEnmarcha NOT IN (" & EstadoFabricacion.Anulada & ") "
                End If

                If codigoEquipo > 0 Then
                    sWhere &= "and opEquipo = " & codigoEquipo & " "
                End If
            End If

            If sWhere <> "" Then
                sSql &= " WHERE " & sWhere
            End If

            If OrdenarOrdenFabrcacion = True Then
                sSql &= " ORDER BY openMarcha desc,opFechaIni,opOrdenMaq"
            Else
                sSql &= " ORDER BY opPuestoTrabajo, opFechaIni"
            End If

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                DameFabricacionesOld = (From elemento In DTDatos
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
                                                            CInt(NoNull(elemento.Item("opIdEnvio"), "D")),
                                                            CStr(NoNull(elemento.Item("opEsLiberada"), "A")),
                                                            CStr(NoNull(elemento.Item("opMensajeError"), "A")),
                                                            CStr(NoNull(elemento.Item("opTieneError"), "A")))).ToList
            Else
                DameFabricacionesOld = New List(Of Fabricaciones)
            End If
        Catch ex As Exception
            DameFabricacionesOld = New List(Of Fabricaciones)
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarOrdenEnvasadoSAP_CantidadFabricada(ByVal CodigoFabricacion As String, ByVal CantidadFabricadaSAP As Integer) As Boolean
        Dim Modificar = False
        Try
            Dim sSql As String = "UPDATE Fabricaciones " &
                                " Set  opCantidadFabricadaSAPV2 = " & CantidadFabricadaSAP & " " &
                                 " WHERE opIdFab=" & CodigoFabricacion

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarOrdenEnvasadoSAP_CantidadFabricadaGeneral(ByVal CodigoFabricacion As String, ByVal CantidadFabricadaSAP As Integer) As Boolean
        Dim Modificar = False
        Try
            Dim sSql As String = "UPDATE Fabricaciones " &
                                " Set  opCantidadFabricadaSAPGeneral = " & CantidadFabricadaSAP & " " &
                                 " WHERE opIdFab=" & CodigoFabricacion

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarOrdenFabricacionSAP_CantidadFabricada(ByVal CodigoFabricacion As String, ByVal CantidadFabricadaSAP As Integer) As Boolean
        Dim Modificar = False
        Try
            Dim sSql As String = "UPDATE Fabricaciones " &
                                " Set  opCantidadFabSAP = " & CantidadFabricadaSAP & " " &
                                 " WHERE opIdFab=" & CodigoFabricacion

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarOrdenFabricacionSAP_HIGIENECOSMETICA(ByVal CodigoFabricacion As String, ByVal HIGIENECOSMETICA As String) As Boolean
        Dim Modificar = False
        Try
            Dim sSql As String = "UPDATE Fabricaciones " &
                                " Set  opHIGIENECOSMETICASAP = '" & HIGIENECOSMETICA & "' " &
                                 " WHERE opIdFab=" & CodigoFabricacion

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarOrdenEnvasadoSAP(ByVal CodigoFabricacion As String, ByVal FaltanteLista As String, ByVal esLiberada As String, ByVal esNumorden As String) As Boolean
        Dim Modificar = False
        Try
            Dim esLiberadaLocal As String = "0"
            If esLiberada.ToUpper.Contains("LIB") Then
                esLiberadaLocal = "1"
            Else
                esLiberadaLocal = "0"
            End If

            'Dim sSql As String = "UPDATE Fabricaciones " &
            '                     " Set  opFaltanteOrdenEnvasadoSAP = '" & FaltanteLista.Trim() & "', opEsLiberadaEnvasado =  '" & esLiberadaLocal & "' " &
            '                     " WHERE opIdFab=" & CodigoFabricacion

            'Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If esNumorden.Trim() <> "" Then
                Dim sSql = "UPDATE Fabricaciones " &
                                 " Set  opFaltanteOrdenEnvasadoSAP = '" & FaltanteLista.Trim() & "', opEsLiberadaEnvasado =  '" & esLiberadaLocal & "' " &
                                 " WHERE opOrdenEnvSAP=" & esNumorden

                Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            End If

            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarOrdenFabricacionSAP(ByVal CodigoFabricacion As String, ByVal FaltanteLista As String, ByVal esLiberada As String, ByVal esNumorden As String) As Boolean
        Dim Modificar = False
        Try
            Dim esLiberadaLocal As String = "0"
            If esLiberada.ToUpper.Contains("LIB") Then
                esLiberadaLocal = "1"
            Else
                esLiberadaLocal = "0"
            End If

            'Dim sSql As String = "UPDATE Fabricaciones " &
            '                     " Set  opFaltanteOrdenFabricacionSAP = '" & FaltanteLista.Trim() & "', opEsLiberada =  '" & esLiberadaLocal & "' " &
            '                     " WHERE opIdFab=" & CodigoFabricacion

            'Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If esNumorden.Trim() <> "" Then
                Dim sSql = "UPDATE Fabricaciones " &
                               " Set  opFaltanteOrdenFabricacionSAP = '" & FaltanteLista.Trim() & "', opEsLiberada =  '" & esLiberadaLocal & "' " &
                               " WHERE opOrdenFabSAP=" & esNumorden

                Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            End If



            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarFechaInicioFinVerificacionSAP(ByVal FechaInicial As Date, ByVal FechaFinal As Date) As Boolean
        Dim Modificar = False
        Try
            Dim FechaInicialLocal As String = FechaInicial.ToString("yyyy-MM-dd HH:mm:ss")
            Dim FechaFinalLocal As String = FechaFinal.ToString("yyyy-MM-dd HH:mm:ss")

            Dim sSql As String = "UPDATE Constantes " &
                                 " Set  cxsValor = '" & FechaInicialLocal.Trim() & "' " &
                                 " WHERE cxsID= 1"

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            sSql = "UPDATE Constantes " &
                                 " Set  cxsValor = '" & FechaFinalLocal.Trim() & "' " &
                                 " WHERE cxsID= 2"

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarFechasPrevYTiempo(ByVal CodigoFabricacion As String, ByVal FechaInicioPrev As Date, ByVal FechaFinPrev As Date, ByVal Tiempo As String) As Boolean
        Dim Modificar = False
        Try
            Dim sSql As String = "UPDATE Fabricaciones " &
                                " Set  opFechaInicioPrevista = '" & FechaInicioPrev & "', opFechaFinPrevista = '" & FechaFinPrev & "', opTiempoPrevisto = '" & Tiempo & "'" &
                                 " WHERE opIdFab=" & CodigoFabricacion

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function ModificarVariables(ByVal CodigoFabricacion As String, ByVal listaRecursos As String, ByVal CantidadEnvasadaRegistro As Integer, ByVal CantidadPlanificadaKg As Double, ByVal StockGranelLocal As Integer, ByVal MaterialPadreV As String) As Boolean
        Dim Modificar = False
        Try

            'listaRecursos:=listaRecursos, CantidadEnvasadaRegistro:=CantidadEnvasadaRegistro,
            '                                       CantidadPlanificadaKg:=CantidadPlanificadaKg, StockGranelLocal:=StockGranelLocal,
            '                                       MaterialPadreV:=MaterialPadreV

            Dim sSql As String = "UPDATE Fabricaciones " &
                                " Set  oplistaRecursos = '" & listaRecursos & "', opCantidadEnvasadaRegistro = " & CantidadEnvasadaRegistro & ", opCantidadPlanificadaKg = " & CantidadPlanificadaKg & ", opStockGranelLocal = " & StockGranelLocal & ", opMaterialPadreV = '" & MaterialPadreV & "' " &
                                 " WHERE opIdFab=" & CodigoFabricacion

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            Return Modificar

            'If Modificar Then
            '    Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            'End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    'Public Function DameRegistrosPullSystem(ByVal FechaInicio As Date,
    '                                        ByVal FechaFin As Date,
    '                                        ByVal Material As String,
    '                                        ByVal TipoMaterial As String,
    '                                        ByVal bSoloConDatos As Boolean,
    '                                        ByVal iDiasControl As Integer,
    '                                        Optional ByVal codPuestoTrabajo As Integer = 0) As List(Of PullSystem)
    '    Try
    '        Dim DTDatos As New DataTable
    '        Dim sWhere As String = " WHERE CAST(CAST(1 AS varchar) + '-' + CAST(fcmes AS varchar) + '-' + CAST(fcanio AS varchar) as datetime) " &
    '                               " BETWEEN '" & sDameFechaCorta(FechaInicio) & "' AND '" & sDameFechaCorta(FechaFin) & "'"

    '        If Material.Trim.Length > 0 Then
    '            sWhere = sWhere & " AND upper(rtrim(fcMaterial)) = '" & UTrim(Material) & "'"
    '        End If

    '        If bSoloConDatos = True Then
    '            sWhere &= " AND fcCantidad>0"
    '        Else
    '            sWhere &= " Or (FC.fcMes Is null And FC.fcAnio Is null And MA.maTipoMat=" & TipoMaterial & " AND maActivo='true')"
    '        End If

    '        Dim sSql As String = "Select " &
    '                             "MA.maCod As fcMaterial, " &
    '                             "FC.fcMes, " &
    '                             "FC.fcAnio, " &
    '                             "isnull(FC.fcCantidad,0) fcCantidad " &
    '                             " FROM Materiales   MA With(nolock) " &
    '                             " left join  ForeCastVentas FC On MA.maCod = Fc.fcMaterial " &
    '                             sWhere

    '        DameRegistrosPullSystem = New List(Of PullSystem)

    '        If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
    '            For Each miRegistro As DataRow In DTDatos.Rows
    '                DameRegistrosPullSystem.Add(New PullSystem(sCodigoMaterial:=CStr(NoNull(miRegistro.Item("fcMaterial"), "A")),
    '                                                           iMes:=CInt(NoNull(miRegistro.Item("fcMes"), "D")),
    '                                                           iAño:=CInt(NoNull(miRegistro.Item("fcAnio"), "D")),
    '                                                           iCantidad:=CInt(NoNull(miRegistro.Item("fcCantidad"), "D")),
    '                                                           idiasControl:=iDiasControl,
    '                                                           StockActual:=0))
    '            Next
    '        End If

    '    Catch ex As Exception
    '        DameRegistrosPullSystem = New List(Of PullSystem)
    '        Throw New Exception(ex.Message & " - -" & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
    '    End Try
    'End Function
    'Public Function DameRegistrosPullSystemSAP(ByVal FechaInicio As Date,
    '                                    ByVal FechaFin As Date,
    '                                    ByVal Material As String,
    '                                    ByVal TipoMaterial As String,
    '                                    ByVal bSoloConDatos As Boolean,
    '                                    ByVal iDiasControl As Integer,
    '                                    Optional ByVal codPuestoTrabajo As Integer = 0) As List(Of PullSystem)
    '    Try

    '        Dim listaForeCast As List(Of BeanForeCastSap)

    '        listaForeCast = DatosSAPConexion.DatosSAP.DameForeCastSAP(Centro:="12",
    '                                                                  Mes:=FechaInicio.Month,
    '                                                                  Anio:=FechaInicio.Year
    '                                                                  )


    '        DameRegistrosPullSystemSAP = New List(Of PullSystem)

    '        If listaForeCast.Count > 0 Then
    '            For Each miRegistro In listaForeCast
    '                DameRegistrosPullSystemSAP.Add(New PullSystem(sCodigoMaterial:=miRegistro.CodMaterial,
    '                                                          iMes:=miRegistro.Mes,
    '                                                          iAño:=miRegistro.Anio,
    '                                                          iCantidad:=miRegistro.CantidadPlanificada,
    '                                                          idiasControl:=iDiasControl,
    '                                                          StockActual:=miRegistro.Stock))
    '            Next

    '        End If


    '    Catch ex As Exception
    '        DameRegistrosPullSystemSAP = New List(Of PullSystem)
    '        Throw New Exception(ex.Message & " - -" & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
    '    End Try
    'End Function

    Public Function DameRegistrosPullSystemSAP2(ByVal MesPS As Integer,
                                       ByVal AnioPS As Integer,
                                       ByVal iDiasControl As Integer) As List(Of PullSystem)
        Try

            Dim listaForeCast As List(Of BeanForeCastSap)

            listaForeCast = DatosSAPConexion.DatosSAP.DameForeCastSAP(Centro:="12",
                                                                        Mes:=MesPS,
                                                                      Anio:=AnioPS)

            DameRegistrosPullSystemSAP2 = New List(Of PullSystem)

            If listaForeCast.Count > 0 Then
                For Each miRegistro In listaForeCast
                    'If miRegistro.CodMaterial.Trim = "70903396" Then
                    '    Dim h = 5
                    'End If
                    DameRegistrosPullSystemSAP2.Add(New PullSystem(sCodigoMaterial:=miRegistro.CodMaterial,
                                                              iMes:=miRegistro.Mes,
                                                              iAño:=miRegistro.Anio,
                                                              iCantidad:=miRegistro.CantidadPlanificada,
                                                              idiasControl:=iDiasControl,
                                                              StockActual:=miRegistro.Stock,
                                                              FechaRotura:=miRegistro.FechaRotura,
                                                              FechaRoturaEntradas:=miRegistro.FechaRoturaEntradas,
                                                              miStockBloqueado:=miRegistro.StockBloqueado,
                                                              miEstatus:=miRegistro.Estatus,
                                                            FechaCorta:=miRegistro.Fecha,
                                                              Necesidad:=miRegistro.Necesidad))
                Next
            End If
        Catch ex As Exception
            DameRegistrosPullSystemSAP2 = New List(Of PullSystem)
            Throw New Exception(ex.Message & " - -" & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    'Public Function DamePedidosVentaOFFLINE(ByVal FechaInicio As Date,
    '                                  ByVal FechaFin As Date
    '                                ) As List(Of PedidosVenta)
    '    Dim miListaPedido = New List(Of PedidosVenta)
    '    Try
    '        Dim DTDatos As New DataTable
    '        Dim sWhere As String = ""
    '        Dim sSql As String = ""

    '        sSql = "SELECT * " &
    '                             "FROM PedidosVentaOFFLINE " &
    '         " where pvFechaPrevista >= '" & FechaInicio & "' and  pvFechaPrevista <= '" & FechaFin & "' "





    '        If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
    '            For Each elemento In DTDatos.Rows
    '                Dim miPedido = New PedidosVenta
    '                miPedido.Fecha = CDate(NoNull(elemento.Item("pvFecha"), "DT"))
    '                miPedido.FechaPrevista = CDate(NoNull(elemento.Item("pvFechaPrevista"), "DT"))
    '                'miPedido.FechaReal = CDate(NoNull(elemento.Item("pvFechaReal"), "F"))
    '                miPedido.CodClienteSolic = CInt(NoNull(elemento.Item("pvCodClienteSolic"), "D"))
    '                miPedido.CodClienteDest = CInt(NoNull(elemento.Item("pvCodClienteDest"), "D"))
    '                miPedido.Material = CStr(NoNull(elemento.Item("pvMaterial"), "A"))
    '                miPedido.Grupo = CStr(NoNull(elemento.Item("pvGrupo"), "A"))
    '                miPedido.Kilos = CInt(NoNull(elemento.Item("pvKilos"), "D"))
    '                miPedido.Unidad = CStr(NoNull(elemento.Item("pvUnidad"), "A"))
    '                miPedido.KilosPtes = CInt(NoNull(elemento.Item("pvKilosPtes"), "D"))
    '                miPedido.UnidadesPtes = CInt(NoNull(elemento.Item("pvUnidadesPtes"), "D"))
    '                miPedido.KgPtes = CDbl(NoNull(elemento.Item("pvKgPtes"), "D"))
    '                miPedido.KilosEnt = CInt(NoNull(elemento.Item("pvKilosEnt"), "D"))
    '                miPedido.LineaPedido = CStr(NoNull(elemento.Item("pvLineaPedido"), "A"))
    '                miPedido.ClaseEntrega = CStr(NoNull(elemento.Item("pvClaseEntrega"), "A"))
    '                miPedido.TipoPosicion = CStr(NoNull(elemento.Item("pvTipoPosicion"), "A"))
    '                miPedido.TipoEnvio = CStr(NoNull(elemento.Item("pvTipoEnvio"), "A"))
    '                miPedido.Centro = CStr(NoNull(elemento.Item("pvCentro"), "A"))
    '                miPedido.Almacen = CStr(NoNull(elemento.Item("pvAlmacen"), "A"))
    '                miPedido.Pedido = CStr(NoNull(elemento.Item("pvPedido"), "A"))
    '                miPedido.OrdenTransporte = CStr(NoNull(elemento.Item("pvOrdenTransporte"), "A"))
    '                miPedido.EstadoOrdenTpte = CStr(NoNull(elemento.Item("pvEstadoOrdenTpte"), "A"))
    '                miPedido.EntregaPendiente = CByte(elemento.Item("pvEntregaPendiente"))
    '                miPedido.NombreMaterial = CStr(NoNull(elemento.Item("pvNombreMaterial"), "A"))
    '                miPedido.NombreCliente = CStr(NoNull(elemento.Item("pvNombreCliente"), "A"))
    '                miPedido.StockActual = CInt(NoNull(elemento.Item("pvStockActual"), "D"))
    '                miPedido.NuevoStockActual = CInt(NoNull(elemento.Item("pvNuevoStockActual"), "D"))
    '                miPedido.NuevoStockAPedidoVenta = CInt(NoNull(elemento.Item("pvNuevoStockAPedidoVenta"), "D"))
    '                miPedido.StatusGLobal = CStr(NoNull(elemento.Item("pvStatusGLobal"), "A"))
    '                miPedido.StatusEntrega = CStr(NoNull(elemento.Item("pvStatusEntrega"), "A"))
    '                miPedido.NombrePuestoTrabajo = CStr(NoNull(elemento.Item("pvNombrePuestoTrabajo"), "A"))
    '                miPedido.FechaPlan = CDate(NoNull(elemento.Item("pvFechaPlan"), "DT"))
    '                miListaPedido.Add(miPedido)
    '            Next

    '        Else
    '            miListaPedido = New List(Of PedidosVenta)
    '        End If
    '        Return miListaPedido
    '    Catch ex As Exception
    '        DamePedidosVentaOFFLINE = New List(Of PedidosVenta)
    '        Throw New Exception(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
    '    End Try
    'End Function


End Module
