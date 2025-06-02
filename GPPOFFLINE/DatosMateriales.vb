
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Module DatosMateriales

    Public Function DamelistaOperacionesHojaRuta(GrupoHojaRuta As String,
                                                 ContGrupo As String,
                                                 CodigoPuestoTrabajo As Integer) As List(Of OperacionesHojaRuta)
        Try
            Dim sWhere As String = ""
            Dim DTDatos As New DataTable

            If GrupoHojaRuta.Trim.Length <> 0 Then
                sWhere = sWhere & " WHERE UPPER(rtrim(opGrupoSAP)) = '" & GrupoHojaRuta.ToString.Trim &
                                  "' AND UPPER(rtrim(opContGrupoSAP))='" & ContGrupo.Trim & "'"
            End If

            If CodigoPuestoTrabajo <> 0 Then
                sWhere = sWhere & " WHERE UPPER(rtrim(opPuestoTrabajo)) = '" & UTrim(CodigoPuestoTrabajo.ToString) & "'"
            End If

            Dim sSql As String = " SELECT *" &
                                  " FROM OperacionesHojaRuta with(nolock) " &
                                  sWhere

            DamelistaOperacionesHojaRuta = New List(Of OperacionesHojaRuta)


            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DamelistaOperacionesHojaRuta.Add(New OperacionesHojaRuta(GrupoHojaRuta:=CStr(miRegistro.Item("opGrupoSAP")),
                                                                             Nodo_SAP:=CInt(miRegistro.Item("opNodoSAP")),
                                                                             ContOper_SAP:=CInt(miRegistro.Item("opContOperSAP")),
                                                                             NumOperacion_SAP:=CStr(miRegistro.Item("opNumOper")),
                                                                             sUnidadMedida:=miRegistro.Item("opUnidadMedida").ToString.Trim,
                                                                             sNombre:=miRegistro.Item("opNombre").ToString.Trim,
                                                                             iCantidadBase:=CInt(miRegistro.Item("opCantidadBase")),
                                                                             iOperarios:=CInt(miRegistro.Item("opOperarios")),
                                                                             iPuestoDeTrabajo:=CInt(miRegistro.Item("opPuestoTrabajo")),
                                                                             dMinutosMaq:=CDec(miRegistro.Item("opMinutos")),
                                                                             ContGrupo_SAP:=CStr(miRegistro.Item("opContGrupoSAP")),
                                                                             ClaveControl_SAP:=CStr(miRegistro.Item("opClaveControl")),
                                                                             dMinutosLimp:=CDec(miRegistro.Item("opMinutosLimpieza")),
                                                                             dMinutosPrep:=CDec(miRegistro.Item("opMinutosPrep")),
                                                                             sCodProveedor:=CStr(miRegistro.Item("opProveedor"))))
                Next
            End If

        Catch ex As Exception
            DamelistaOperacionesHojaRuta = New List(Of OperacionesHojaRuta)
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function

    Public Function DamePuestosTrabajoHojaRuta(GrupoHojaRuta As String,
                                               ContHojaRuta As String,
                                               TipoPuestoTrabajo As String) As List(Of PuestosTrabajo)
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = " SELECT DISTINCT puestostrabajo.* " &
                                 " FROM puestostrabajo " &
                                 " WHERE ptCod IN " &
                                 "(SELECT DISTINCT opPuestoTrabajo " &
                                 " FROM OperacionesHojaRuta with(nolock) " &
                                 " WHERE opGrupoSAP = '" & UTrim(GrupoHojaRuta.ToString) &
                                 "' AND opContGrupoSAP='" & UTrim(ContHojaRuta.ToString) & "')"

            If TipoPuestoTrabajo.Trim.Length <> 0 Then
                sSql &= " AND ptTipo='" & TipoPuestoTrabajo & "'"
            End If

            DamePuestosTrabajoHojaRuta = New List(Of PuestosTrabajo)

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DamePuestosTrabajoHojaRuta.Add(New PuestosTrabajo(_CodigoPuestoTrabajo:=CInt(miRegistro.Item("ptCod")),
                                                                      _Nombre:=CStr(miRegistro.Item("ptNombre")),
                                                                      _Centro:=CStr(miRegistro.Item("ptCentro")),
                                                                      _Operarios:=CInt(miRegistro.Item("ptOperarios")),
                                                                      _Tipo:=CStr(miRegistro.Item("ptTipo")),
                                                                      _AreaProduccion:=CStr(miRegistro.Item("ptAreaProd")),
                                                                      _Activo:=CBool(CStr(miRegistro.Item("ptActivo"))),
                                                                      _Recurso:=CBool(CStr(NoNull(miRegistro.Item("ptRecurso"), "N"))),
                                                                      _CambioPedido:=CBool(CStr(miRegistro.Item("ptCambiarPedido"))),
                                                                      _Orden:=CInt(miRegistro.Item("ptOrden")),
                                                                      _VelocidadMax:=CInt(miRegistro.Item("ptVelMax")),
                                                                      _VelocidadActual:=CInt(miRegistro.Item("ptVelActual")),
                                                                      _CodCentroProd:=CInt(miRegistro.Item("ptCentroProd")),
                                                                      _CodProveedor:=CStr(miRegistro.Item("ptProveedor")),
                                                                       _EsCentroExterno:=CBool(CStr(miRegistro.Item("ptEsCoperativaExt"))),
                                                                      _PuedeRealizarTraspasos:=CBool(CStr(miRegistro.Item("ptPuedeRealizarTraspaso")))))
                Next
            End If

        Catch ex As Exception
            DamePuestosTrabajoHojaRuta = New List(Of PuestosTrabajo)
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function

    Public Function DameGrupoCompras() As List(Of GrupoCompras)
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = "select * from  GrupoCompras with(nolock)"

            DameGrupoCompras = New List(Of GrupoCompras)
            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DameGrupoCompras.Add(New GrupoCompras(sCodigo:=CStr(miRegistro.Item("gcCod")),
                                                          sNombre:=CStr(miRegistro.Item("gcNombre"))))
                Next
            End If

        Catch ex As Exception
            DameGrupoCompras = New List(Of GrupoCompras)
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function

    Public Function DameHojasRutaMaterial(CodigoMaterial As String,
                                          Optional bQuitarBorradas As Boolean = True) As List(Of HojaRuta)
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = " SELECT hojaRuta.* " &
                                 " FROM HojaRutaxMaterial INNER JOIN hojaRuta on hrGrupo=hmGrupoHR AND hrContGrupo=hmContGrupoHR" &
                                 " WHERE hmMaterial='" & CodigoMaterial.Trim & "'"

            If bQuitarBorradas = True Then
                sSql &= " AND hrBorrada='false'"
            End If

            DameHojasRutaMaterial = New List(Of HojaRuta)

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DameHojasRutaMaterial.Add(New HojaRuta(sGrupo:=CStr(miRegistro.Item("hrGrupo")),
                                                           sContGrupo:=CStr(miRegistro.Item("hrContGrupo")),
                                                           sUnidadMedida:=CStr(miRegistro.Item("hrUnidadMedida")),
                                                           sNombre:=CStr(miRegistro.Item("hrNombre")),
                                                           sCentro:=CStr(miRegistro.Item("hrCentro")),
                                                           sTipoHR:=CStr(miRegistro.Item("hrTipoHR")),
                                                           bBorrada:=CBool(CStr(miRegistro.Item("hrBorrada"))),
                                                           sFormato:=CStr(miRegistro.Item("hrFormato"))))
                Next
            End If

        Catch ex As Exception
            DameHojasRutaMaterial = New List(Of HojaRuta)
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function


    Public Function DameProducto(CodigoMaterial As String, ByVal tipoMaterial As String) As List(Of Material)
        Try
            Dim materialPadre As Material = New Material(Codigo:=CodigoMaterial)
            Dim misMateriales As New List(Of Material)
            If materialPadre IsNot Nothing Then
                If materialPadre.Creado = True Then
                    'funcion recursiva
                    consultaRecursivaMaterial(materialPadre, misMateriales, tipoMaterial)
                End If

            End If
            Return misMateriales
        Catch ex As Exception
            DameProducto = Nothing
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function

    Private Sub consultaRecursivaMaterial(ByVal miMaterialPadre As Material, ByRef listafabricaciones As List(Of Material), ByVal tipoMaterial As String)
        Try
            For Each elemento In miMaterialPadre.CabLista_Material.MaterialesLista
                If elemento.Material.Tipo = tipoMaterial Then
                    listafabricaciones.Add(elemento.Material)
                    If tipoMaterial <> ConstantesGPP.TipoMaterial.Fabricaciones Then
                        consultaRecursivaMaterial(elemento.Material, listafabricaciones, tipoMaterial)
                    End If

                Else
                    If tipoMaterial <> ConstantesGPP.TipoMaterial.Fabricaciones Then
                        consultaRecursivaMaterial(elemento.Material, listafabricaciones, tipoMaterial)
                    End If

                End If
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub




    'Public Function DameHojaRutaxMaterial(sGrupoHR As String) As List(Of HojaRutaxMaterial)
    '    Try
    '        Dim DTDatos As New DataTable
    '        Dim sSql As String = " SELECT * " &
    '                             " FROM HojaRutaxMaterial with(nolock) " &
    '                             " WHERE hmGrupoHR='" & sGrupoHR.Trim & "'"

    '        DameHojaRutaxMaterial = New List(Of HojaRutaxMaterial)

    '        If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
    '            For Each miRegistro As DataRow In DTDatos.Rows
    '                DameHojaRutaxMaterial.Add(New HojaRutaxMaterial(sCodigoMaterial:=CStr(miRegistro.Item("hmMaterial")),
    '                                                       sGrupoHR:=CStr(miRegistro.Item("hmGrupoHR")),
    '                                                       sContadorGrupoHR:=CStr(miRegistro.Item("hmContGrupoHR"))))
    '            Next
    '        End If

    '    Catch ex As Exception
    '        DameHojaRutaxMaterial = New List(Of HojaRutaxMaterial)
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & MethodBase.GetCurrentMethod().DeclaringType.Name & "." & MethodInfo.GetCurrentMethod.Name & "()", ex)
    '    End Try
    'End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Tipo"></param>
    ''' <param name="Mostrar_Activos">1-Activos; 2-Inactivos; 0-Todos</param>
    ''' <returns></returns>
    Public Function DameMateriales(Tipo As String,
                                   Optional Mostrar_Activos As Integer = 0) As List(Of Material)
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = " SELECT * " &
                                 " FROM Materiales "

            Select Case Mostrar_Activos
                Case 1 'Activos
                    sSql &= " WHERE maActivo='true'"

                Case 2 'Mostrar Inactivos
                    sSql &= " WHERE maActivo='false'"

                Case Else 'TODOS
                    sSql &= " WHERE 1=1"
            End Select

            If Tipo.ToString <> ConstantesGPP.TipoMaterial.NINGUNO Then
                sSql &= " AND UPPER(rtrim(maTipoMat)) in (" & UTrim(Tipo.ToString) & ")"
            End If

            DameMateriales = New List(Of Material)

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DameMateriales.Add(New Material(Codigo:=miRegistro.Item("maCod").ToString.Trim,
                                                    Tipo:=miRegistro.Item("maTipoMat").ToString.Trim,
                                                    Grupo:=miRegistro.Item("maGrupoArt").ToString.Trim,
                                                    UnidadMedida:=miRegistro.Item("maUMBase").ToString.Trim,
                                                    Nombre:=miRegistro.Item("maNombre").ToString.Trim.ToUpper,
                                                    Lista_Mat:=miRegistro.Item("maListaMaterial").ToString.Trim,
                                                    Familia_Envasado:=CByte(miRegistro.Item("maFamiliaEnvasado")),
                                                    Fecha_IniPS:=CDate(miRegistro.Item("maFecIniPS")),
                                                    Fecha_FinPS:=CDate(miRegistro.Item("maFecFinPS")),
                                                    Dias_PP:=CByte(miRegistro.Item("maDiasPP")),
                                                    Stock_MaxPS:=CInt(miRegistro.Item("maStokMaxPS")),
                                                    Stock_MinPS:=CInt(miRegistro.Item("maStokMinPS")),
                                                    Activo:=CBool(miRegistro.Item("maActivo")),
                                                    Lote_Minimo:=CInt(miRegistro.Item("maLoteMin").ToString.Trim),
                                                    Lote_Maximo:=CInt(miRegistro.Item("maLoteMax").ToString.Trim),
                                                    Lote_Fijo:=CInt(miRegistro.Item("maLoteFijo").ToString.Trim),
                                                    Redondeo_Lote:=CInt(miRegistro.Item("maRedondeo").ToString.Trim),
                                                    Dias_FabPropia:=CInt(miRegistro.Item("maDiasFabPropia").ToString.Trim),
                                                    Tipo_TamañoLote:=CStr(miRegistro.Item("maTipoTamLote").ToString.Trim),
                                                    Grupo_HojaRuta:=CStr(miRegistro.Item("maGrupoHR").ToString.Trim),
                                                    Contador_HojaRuta:=CStr(miRegistro.Item("maContHR").ToString.Trim),
                                                    Grupo_Compra:=CStr(miRegistro.Item("maGrupoCompra").ToString.Trim),
                                                    Mostrar_Informes:=CBool(miRegistro.Item("mnMostrarInformes")),
                                                    Unidades_Pack:=CInt(miRegistro.Item("maUnidadesPACK")),
                                                    UnidadesPorPalet:=CInt(miRegistro.Item("maUnidadesPalet")),
                                                    MesesLoteCarga:=CInt(miRegistro.Item("maMesesLoteCarga"))))
                Next
            End If

        Catch ex As Exception
            DameMateriales = New List(Of Material)
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function

    Public Function DameListasMaterial(ListaMaterial As String) As List(Of ListaMaterial)
        Try
            Dim sWhere As String = ""
            Dim sINNERJoin As String = ""

            Dim DTDatos As New DataTable

            If ListaMaterial.Trim.Length > 0 Then
                sWhere = sWhere & " AND UPPER(RTRIM(dlLista)) = '" & UTrim(ListaMaterial) & "'"
            End If

            Dim sSql As String = " SELECT * " &
                                 " FROM ListaMateriales " &
                                 " WHERE dlNodo > 0" &
                                 sWhere

            DameListasMaterial = New List(Of ListaMaterial)
            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DameListasMaterial.Add(New ListaMaterial(UTrim(miRegistro.Item("dlLista")),
                                                              CInt(NoNull(miRegistro.Item("dlNodo"), "D")),
                                                              UTrim(miRegistro.Item("dlMaterial")),
                                                              UTrim(miRegistro.Item("dlPosicion")),
                                                              CDbl(NoNull(miRegistro.Item("dlCantidad"), "D")),
                                                              UTrim(miRegistro.Item("dlUM")),
                                                              UTrim(miRegistro.Item("dlTipoPos")),
                                                             CDbl(NoNull(miRegistro.Item("dlPorcMerma"), "D"))))

                Next
            End If

        Catch ex As Exception
            DameListasMaterial = New List(Of ListaMaterial)
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function

    Public Function DameMaterialSeleccionado(MACOD As String,
                                  Optional Mostrar_Activos As Integer = 0) As List(Of Material)
        Try
            Dim DTDatos As New DataTable
            Dim sSql As String = " SELECT * " &
                                 " FROM Materiales "

            Select Case Mostrar_Activos
                Case 1 'Activos
                    sSql &= " WHERE maActivo='true'"

                Case 2 'Mostrar Inactivos
                    sSql &= " WHERE maActivo='false'"

                Case Else 'TODOS
                    sSql &= " WHERE 1=1"
            End Select

            If MACOD <> "" Then
                sSql &= " AND MACOD = '" & MACOD & "'  "
            End If

            DameMaterialSeleccionado = New List(Of Material)

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                For Each miRegistro As DataRow In DTDatos.Rows
                    DameMaterialSeleccionado.Add(New Material(Codigo:=miRegistro.Item("maCod").ToString.Trim,
                                                    Tipo:=miRegistro.Item("maTipoMat").ToString.Trim,
                                                    Grupo:=miRegistro.Item("maGrupoArt").ToString.Trim,
                                                    UnidadMedida:=miRegistro.Item("maUMBase").ToString.Trim,
                                                    Nombre:=miRegistro.Item("maNombre").ToString.Trim.ToUpper,
                                                    Lista_Mat:=miRegistro.Item("maListaMaterial").ToString.Trim,
                                                    Familia_Envasado:=CByte(miRegistro.Item("maFamiliaEnvasado")),
                                                    Fecha_IniPS:=CDate(miRegistro.Item("maFecIniPS")),
                                                    Fecha_FinPS:=CDate(miRegistro.Item("maFecFinPS")),
                                                    Dias_PP:=CByte(miRegistro.Item("maDiasPP")),
                                                    Stock_MaxPS:=CInt(miRegistro.Item("maStokMaxPS")),
                                                    Stock_MinPS:=CInt(miRegistro.Item("maStokMinPS")),
                                                    Activo:=CBool(miRegistro.Item("maActivo")),
                                                    Lote_Minimo:=CInt(miRegistro.Item("maLoteMin").ToString.Trim),
                                                    Lote_Maximo:=CInt(miRegistro.Item("maLoteMax").ToString.Trim),
                                                    Lote_Fijo:=CInt(miRegistro.Item("maLoteFijo").ToString.Trim),
                                                    Redondeo_Lote:=CInt(miRegistro.Item("maRedondeo").ToString.Trim),
                                                    Dias_FabPropia:=CInt(miRegistro.Item("maDiasFabPropia").ToString.Trim),
                                                    Tipo_TamañoLote:=CStr(miRegistro.Item("maTipoTamLote").ToString.Trim),
                                                    Grupo_HojaRuta:=CStr(miRegistro.Item("maGrupoHR").ToString.Trim),
                                                    Contador_HojaRuta:=CStr(miRegistro.Item("maContHR").ToString.Trim),
                                                    Grupo_Compra:=CStr(miRegistro.Item("maGrupoCompra").ToString.Trim),
                                                    Mostrar_Informes:=CBool(miRegistro.Item("mnMostrarInformes")),
                                                    Unidades_Pack:=CInt(miRegistro.Item("maUnidadesPACK")),
                                                    UnidadesPorPalet:=CInt(miRegistro.Item("maUnidadesPalet")),
                                                    MesesLoteCarga:=CInt(miRegistro.Item("maMesesLoteCarga"))))
                Next
            End If

        Catch ex As Exception
            DameMaterialSeleccionado = New List(Of Material)
            Throw New Exception(ex.Message & " -- " & "()", ex)
        End Try
    End Function

End Module
