
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
'Imports NegocioGPP.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP
'Imports AccesoDatos
'Imports Utilidades.Util
'Imports Utilidades
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Configuration


Public Class DatosGPP

#Region "Atributos"
    Private Shared misDatos As DatosGPP
    Private BaseDatosGPL As BaseDatos
    Private BaseDatosPerseida As BaseDatos
    Private Shared oConnection As Object
    'Private miSistema As Sistema
    Public Property ConfiguracionSap As String = ConstantesGPP.ConfiguracionSAP.Desarrollo
    Private miUsuario As New Usuario
    'Private miAlmacenBigBags As AlmacenBigBags

#End Region

#Region "Constructores"
    ''' <summary>
    ''' Constructor vacio.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
        Try
            Conectar()

        Catch ex As Exception
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub
#End Region

#Region "Propiedades"

    Public Sub ForzarReinicio()
        Try
            DatosGPP.misDatos = Nothing
        Catch ex As Exception
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Public Property Usuario() As Usuario
        Get
            Usuario = miUsuario
        End Get
        Set(ByVal value As Usuario)
            miUsuario = value
        End Set
    End Property

    Public Property CGPL() As BaseDatos
        Get
            Try
                If Me.BaseDatosGPL.EstadoConexion <> ConnectionState.Closed AndAlso Me.BaseDatosGPL.EstadoConexion <> ConnectionState.Broken Then
                    CGPL = BaseDatosGPL
                Else
                    If Me.BaseDatosGPL.EstadoConexion <> ConnectionState.Connecting Then
                        'reconectar
                        BaseDatosGPL = New BaseDatos(ConfigurationManager.AppSettings.Get("SISTEMA_ESCOGIDO"))
                        BaseDatosGPL.Conectar()
                    End If
                    CGPL = BaseDatosGPL
                End If

            Catch ex As Exception
                CGPL = BaseDatosGPL
                Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & MethodInfo.GetCurrentMethod.Name & "() ", ex)
            End Try
        End Get
        Set(ByVal value As BaseDatos)
            BaseDatosGPL = value
        End Set
    End Property

    Public Property CGPLPerseida() As BaseDatos
        Get
            Try
                If Me.BaseDatosPerseida.EstadoConexion <> ConnectionState.Closed AndAlso Me.BaseDatosPerseida.EstadoConexion <> ConnectionState.Broken Then
                    CGPLPerseida = BaseDatosPerseida
                Else
                    If Me.BaseDatosPerseida.EstadoConexion <> ConnectionState.Connecting Then
                        'reconectar
                        BaseDatosPerseida = New BaseDatos("PRODUCCIONPERSEIDA")
                        BaseDatosPerseida.Conectar()
                    End If
                    CGPLPerseida = BaseDatosPerseida
                End If

            Catch ex As Exception
                CGPLPerseida = BaseDatosPerseida
                'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & MethodInfo.GetCurrentMethod.Name & "() ", ex)
            End Try
        End Get
        Set(ByVal value As BaseDatos)
            BaseDatosPerseida = value
        End Set
    End Property

    ''' <summary>
    ''' Instancia unica de la clase DatosMantenimiento.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property Datos() As DatosGPP
        Get
            If misDatos Is Nothing Then
                misDatos = New DatosGPP
            End If
            Return misDatos
        End Get
    End Property

    'Public Property Sistema As Sistema
    '    Get
    '        Return Me.miSistema
    '    End Get
    '    Set(value As Sistema)
    '        Me.miSistema = value
    '    End Set
    'End Property

    Public Function Conectar() As Boolean
        Try
            Me.BaseDatosGPL = New BaseDatos(ConfigurationManager.AppSettings.Get("SISTEMA_ESCOGIDO"))

            Me.ConfiguracionSap = CStr(IIf(ConfigurationManager.AppSettings.Get("SISTEMA_ESCOGIDO") = "PRODUCCION",
                                           ConfigurationManager.AppSettings.Get("SISTEMA_ESCOGIDO"),
                                           ConfigurationManager.AppSettings.Get("SAP_ESCOGIDO")))

            Me.BaseDatosGPL.Conectar()
            Dim v = ConfigurationManager.AppSettings.Get("PRODUCCIONPERSEIDA")
            Me.BaseDatosPerseida = New BaseDatos("PRODUCCIONPERSEIDA")
            Me.BaseDatosPerseida.Conectar()
            Conectar = True
        Catch ex As Exception
            Conectar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function ConectarPerseida() As Boolean
        Try
            Me.BaseDatosPerseida = New BaseDatos(ConfigurationManager.AppSettings.Get("PRODUCCIONPERSEIDA"))

            'Me.ConfiguracionSap = CStr(IIf(ConfigurationManager.AppSettings.Get("SISTEMA_ESCOGIDO") = "PRODUCCION",
            '                               ConfigurationManager.AppSettings.Get("SISTEMA_ESCOGIDO"),
            '                               ConfigurationManager.AppSettings.Get("SAP_ESCOGIDO")))

            Me.BaseDatosPerseida.Conectar()
            ConectarPerseida = True
        Catch ex As Exception
            ConectarPerseida = False
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    'Public Property AlmacenBigBags As AlmacenBigBags
    '    Get
    '        Return Me.miAlmacenBigBags
    '    End Get
    '    Set(value As AlmacenBigBags)
    '        Me.miAlmacenBigBags = value
    '    End Set
    'End Property

    Public Function ConexionAbierta() As Boolean
        Try
            ConexionAbierta = False
            If misDatos.CGPL.EstadoConexion = ConnectionState.Open Then
                ConexionAbierta = True
            End If
        Catch ex As Exception
            ConexionAbierta = False
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    ''' <summary>
    ''' Desconecta la Base de datos.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DesconectarBBDD()
        Try
            If CGPL.EstadoConexion = ConnectionState.Open Then
                CGPL.Desconectar()
            End If

            MyBase.Finalize()
        Catch ex As Exception
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Function GuardarLog(DescripcionOperacion As String,
                                 CodigoAfectado As String,
                                 Optional CodigoCliente As Integer = 0) As Boolean
        Try
            'Dim miLog As New Log(0,
            '                    Datos.Usuario.CodigoSociedadActual,
            '                    Now,
            '                    DescripcionOperacion,
            '                    Datos.Usuario.Codigo,
            '                    Datos.Usuario.OperacionActual.Codigo,
            '                    0,
            '                    CodigoAfectado,
            '                    CodigoCliente)

            'GuardarLog = miLog.Insertar

        Catch ex As Exception
            GuardarLog = False
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    ''' <summary>
    ''' Devuelve las operaciones del programa en funcion de los parametros pasados
    ''' </summary>
    ''' <param name="CodigoPerfil">Perfil del usuario. Si es = 0, no lo tiene en cuenta</param>
    ''' <param name="OrdenarNombre">Si es TRUE, entonces ordena las operaciones por el nombre, sino lo hara por el codigo de la operacion</param>
    ''' <returns></returns>
    ''' <remarks></remarks>






    ''' <summary>
    ''' Funcion que nos devuelve los favoritos de un Usuario
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>




    ''' <summary>
    ''' Devuelve las tareas de la agenda en funcion de las opciones pasadas
    ''' </summary>
    ''' <param name="CodigoUsuario">Si el codiog es = 0, no se tiene en cuenta</param>
    ''' <param name="Perfiles">Lista de perfiles. Si no se pasan perfiles, no se tiene en cuenta</param>
    ''' <param name="Fecha">Fecha de las tareas a mostrar, si igual a FECHAGLOBAL no se tiene en cuenta</param>
    ''' <param name="TipoConsulta">Tipo de consulta. 1-> Personal:2->Grupo;3->SinAsignar; 0->Ninguna de la anteriores</param>
    ''' <param name="SoloPendientes">Solo muestran las tareas pendientes de la agenda</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    ''' <summary>
    ''' Esta función se usa para cargar combos con los datos de los usuarios
    ''' </summary>
    ''' <returns>Lista de usuarios, solo con código y nombre</returns>
    ''' <remarks></remarks>


    Public Function CodigoPINRepetido(CodUsuario As Integer,
                                      CodigoPIN As String) As Boolean
        Try
            Dim DTDatos As New DataTable

            Dim sSql As String = "SELECT * " &
                                 " FROM usuarios " &
                                 " WHERE  usPin='" & CodigoPIN.Trim &
                                 "' AND usCodigo<>" & CodUsuario

            CodigoPINRepetido = Datos.CGPL.DameDatosDT(sSql, DTDatos)

        Catch ex As Exception
            CodigoPINRepetido = False
        End Try
    End Function







    ''' <summary>
    ''' Devuelve los registros del fore cast de ventas de un mes en concreto teniendo en cuentas los registros info.
    ''' </summary>
    ''' <param name="Mes"></param>
    ''' <param name="Año"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public Function DameRegistrosInfoForeCastVentas(ByVal Mes As Integer,
    '                                                ByVal Año As Integer,
    '                                                ByVal CodigoCliente As Integer,
    '                                                ByVal CodigoMaterial As String,
    '                                                ByVal CodigoGrupo As String,
    '                                                ByVal CodigoPresentacion As Integer,
    '                                                ByVal SoloConToneladas As Boolean,
    '                                                ByVal SoloTipoPreciosActivos As Boolean,
    '                                                ByVal SoloAceptados As Boolean) As List(Of ForeCastVentas)
    '    Try
    '        Dim sWhere As String = " WHERE clCod > 0 "

    '        If CodigoCliente > 0 Then
    '            sWhere = sWhere & " AND clCod = " & CodigoCliente
    '        End If

    '        If CodigoMaterial.Trim.Length > 0 Then
    '            sWhere = sWhere & " AND upper(rtrim(maCod))= '" & UTrim(CodigoMaterial) & "'"
    '        End If

    '        If CodigoGrupo.Trim.Length > 0 Then
    '            sWhere = sWhere & " AND upper(rtrim(maGrupoArt)) = '" & UTrim(CodigoGrupo) & "'"
    '        End If

    '        If CodigoPresentacion > 0 Then
    '            sWhere = sWhere & " AND maTipoEmbalaje = " & CodigoPresentacion
    '        End If

    '        If SoloConToneladas Then
    '            sWhere = sWhere & " AND fcTon > 0 "
    '        End If

    '        If SoloTipoPreciosActivos Then
    '            sWhere = sWhere & " AND tpActivo = 'TRUE'"
    '        End If

    '        If SoloAceptados Then
    '            sWhere = sWhere & " AND fcAceptado = 'TRUE'"
    '        End If

    '        Dim sSql As String = " SELECT clCod,clNombre,maCod,maNombre,maGrupoArt,clPais,maTipoEmbalaje,fcTon,fcPorcSem1,fcPorcSem2,fcPorcSem3,fcPorcSem4,fcPorcSem5," &
    '                             " tpNombre, case  isnull(fcPrecio,0) when 0 then riPrecioTN else fcPrecio end as Precio,fcaceptado," &
    '                             " case isnull(fcTipoPrecio,0) when 0 then riTipoPrecio else fctipoprecio end  as TipoPrecio," &
    '                             " ((tpPrecioParaxileno*tpCteParaxileno)+(tpPrecioMEG*tpCteMEG))+tpCteBRM as PrecioBRM,isnull(tpBRM,'false') as tpBRM " &
    '                             " FROM registroinfoventas " &
    '                             " INNER join Clientes on ricliente = clCod  " &
    '                             " INNER join materiales on rimaterial = maCod " &
    '                             " LEFT JOIN TipoPrecio ON tpCodigo = riTipoPrecio " &
    '                             " LEFT JOIN ForeCastVentas ON riCliente= fcCliente AND riMaterial = fcMaterial AND fcTipoPrecio = riTipoPrecio " &
    '                             " AND fcMes = " & Mes & " AND fcAnio = " & Año &
    '                             sWhere &
    '                             " UNION ALL" &
    '                             " SELECT clCod,clNombre,maCod,maNombre,maGrupoArt,clPais,maTipoEmbalaje,fcTon,fcPorcSem1,fcPorcSem2,fcPorcSem3,fcPorcSem4,fcPorcSem5,tpNombre," &
    '                             " case  isnull(fcPrecio,0) when 0 then riPrecioTN else fcPrecio end As Precio,fcaceptado, " &
    '                             " case isnull(fcTipoPrecio,0) when 0 then riTipoPrecio else fctipoprecio end  As TipoPrecio, " &
    '                             " ((tpPrecioParaxileno * tpCteParaxileno) + (tpPrecioMEG * tpCteMEG))+tpCteBRM As PrecioBRM, " &
    '                             " isnull(tpBRM,'false') as tpBRM " &
    '                             " From ForeCastVentas " &
    '                             " INNER Join Clientes on fcCliente = clcod  " &
    '                             " INNER Join Materiales on fcMaterial = maCod " &
    '                             " LEFT Join tipoPrecio on fcTipoPrecio = tpCodigo " &
    '                             " LEFT Join RegistroInfoVentas ON fcCliente = riCliente And fcMaterial = riMaterial And fcTipoPrecio = riTipoPrecio " &
    '                             sWhere & " And fcMes = " & Mes & " AND  fcAnio = " & Año & " AND riCliente IS NULL "

    '        Dim DTDatos As New DataTable

    '        DameRegistrosInfoForeCastVentas = New List(Of ForeCastVentas)

    '        If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
    '            For Each miRegistro As DataRow In DTDatos.Rows
    '                DameRegistrosInfoForeCastVentas.Add(New ForeCastVentas(CInt(NoNull(miRegistro.Item("clCod"), "D")),
    '                                                                UTrim(miRegistro.Item("clNombre")),
    '                                                                UTrim(miRegistro.Item("maCod")),
    '                                                                UTrim(miRegistro.Item("maNombre")),
    '                                                                Mes,
    '                                                                Año,
    '                                                                CInt(NoNull(miRegistro.Item("fcTon"), "D")),
    '                                                                CShort(NoNull(miRegistro.Item("fcPorcSem1"), "D")),
    '                                                                CShort(NoNull(miRegistro.Item("fcPorcSem2"), "D")),
    '                                                                CShort(NoNull(miRegistro.Item("fcPorcSem3"), "D")),
    '                                                                CShort(NoNull(miRegistro.Item("fcPorcSem4"), "D")),
    '                                                                CShort(NoNull(miRegistro.Item("fcPorcSem5"), "D")),
    '                                                                UTrim(miRegistro.Item("clPais")),
    '                                                                UTrim(miRegistro.Item("maGrupoArt")),
    '                                                                CType(CInt(NoNull(miRegistro.Item("maTipoEmbalaje"), "D")), ConstantesGPP.TipoEmbalaje),
    '                                                                CDbl(IIf(CBool(NoNull(miRegistro.Item("tpBRM"), "D")) = True, CDbl(NoNull(miRegistro.Item("PrecioBRM"), "D")), CDbl(NoNull(miRegistro.Item("Precio"), "D")))),
    '                                                                CInt(NoNull(miRegistro.Item("TipoPrecio"), "D")),
    '                                                                CBool(NoNull(miRegistro.Item("fcAceptado"), "D"))) With {.NombreTipoPrecio = UTrim(miRegistro.Item("tpNombre"))})
    '            Next
    '        End If

    '    Catch ex As Exception
    '        DameRegistrosInfoForeCastVentas = New List(Of ForeCastVentas)
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & MethodBase.GetCurrentMethod().DeclaringType.Name & "." & MethodInfo.GetCurrentMethod.Name & "()", ex)
    '    End Try
    'End Function

    ''' <summary>
    ''' Devuelve los registros del forecast de ventas reales, es decir sin tener en cuenta los registros info
    ''' </summary>
    ''' <param name="FechaInicio"></param>
    ''' <param name="FechaFin"></param>
    ''' <param name="Material"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>




    'Public Function CrearTablaTemporalNecesidades(RegistrosForeCast As List(Of ForeCastVentas)) As Boolean
    '    Try

    '        'misRegistrosForeCast = Datos.DameRegistrosForeCastVentas(FechaInicio, FechaFin, "", 0, Material.Codigo, "", Constantes.TipoEmbalaje.Desconocido)
    '        CrearTablaTemporalNecesidades = False
    '        If RegistrosForeCast.Count > 0 Then

    '            Dim sSql As String = " DELETE FROM TemporalNecesidades " &
    '                                 " WHERE tmUsuario = " & Datos.Usuario.Codigo

    '            Datos.CGPL.EjecutarConsulta(sSql)

    '            For Each miRegistro As ForeCastVentas In RegistrosForeCast
    '                If miRegistro.Cantidad > 0 Then
    '                    For Each miLista As ListaMaterial In miRegistro.Material.ListasMateriales
    '                        'Pasamos el numero de toneladas a kilos
    '                        GuardarRegistroTemporalNecesidades(miLista, miRegistro.Cantidad * 1000, miRegistro.CodigoMaterial)
    '                    Next
    '                End If
    '            Next

    '            CrearTablaTemporalNecesidades = True
    '        End If
    '    Catch ex As Exception
    '        CrearTablaTemporalNecesidades = False
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & MethodBase.GetCurrentMethod().DeclaringType.Name & "." & MethodInfo.GetCurrentMethod.Name & "()", ex)
    '    End Try
    'End Function

    'Public Sub GuardarRegistroTemporalNecesidades(Lista As ListaMaterial,
    '                                              CantidadNecesaria As Double,
    '                                              CodigoMaterialPadre As String)
    '    Try
    '        Dim sSQL As String = ""
    '        Dim dCantidad As Double = 0

    '        dCantidad = Math.Round(Lista.Cantidad * CantidadNecesaria / Lista.Cabecera.CantidadBase, 10)

    '        If Lista.Material.ListasMateriales.Count > 0 Then
    '            For Each miLista As ListaMaterial In Lista.Material.ListasMateriales
    '                GuardarRegistroTemporalNecesidades(miLista, dCantidad, Lista.CodigoMaterial)
    '            Next
    '        End If

    '        sSQL = "INSERT INTO TemporalNecesidades(tmUsuario,tmMaterial,tmCantidadNecesaria,tmUNBase,tmMaterialPadre) " &
    '               " VALUES(" & Datos.Usuario.Codigo & ",'" &
    '                             UTrim(Lista.CodigoMaterial) & "'," &
    '                             PuntoComa(Math.Round(dCantidad, 3)) & ",'" &
    '                             UTrim(Lista.UnidadMedida) & "','" &
    '                             CodigoMaterialPadre & "')"

    '        Datos.CGPL.EjecutarConsulta(sSQL)

    '    Catch ex As Exception
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & MethodBase.GetCurrentMethod().DeclaringType.Name & "." & MethodInfo.GetCurrentMethod.Name & "()", ex)
    '    End Try
    'End Sub

    ''' <summary>
    ''' Devuelve los registros del temporal de necesidades con la cantidad acumulada
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>


    Public Enum TipoAgrupacion
        PorMaterial
        PorGrupo
        PorCliente
    End Enum

    Public Enum SubAgrupacion
        PorGrupo
        PorMaterial
        Sin_Agrupar
    End Enum

#End Region

#Region "Métodos Auxiliares"

    'Public Shared Function ValidarLote(ByVal descLote As StructureLote,
    '                                   ByVal ZZCALIDAD As String,
    '                                   ByVal ZZVISCOSIDAD As String,
    '                                   ByVal ZZCALADI As String,
    '                                   Optional ByVal HayParticionesLote As Boolean = False) As Boolean
    '    Try
    '        'Si hay partición de lote no compruebo ya que solo puedo cargar ese.
    '        If HayParticionesLote Then Return True

    '        Dim zzCalidad_ As String = CStr(IIf(ZZCALIDAD = "", "00", ZZCALIDAD))
    '        'Viscosidad ???:
    '        Dim zzViscosidad_ As String = CStr(IIf(ZZVISCOSIDAD = "", "00", ZZVISCOSIDAD))
    '        'Instalación: 1,2 y 3
    '        Dim zzInstalacion As String = zzViscosidad_(0)
    '        'Motivo: 0 a 9
    '        Dim zzMotivo As String = zzViscosidad_(1)

    '        'Causa: 00 a 99
    '        Dim zzCaladi_ As String = CStr(IIf(ZZCALADI = "", "00", ZZCALADI))

    '        If zzCalidad_ <> descLote.Calidad Then
    '            Return False
    '        ElseIf zzInstalacion <> "0" AndAlso zzInstalacion <> descLote.Instalacion Then
    '            Return False
    '        ElseIf zzMotivo <> "0" AndAlso zzMotivo <> descLote.Motivo Then
    '            Return False
    '        ElseIf zzCaladi_ <> "00" AndAlso zzCaladi_ <> descLote.Causa Then
    '            Return False
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        Throw New NegocioDatosExcepction(ex.Message & " - " & System.Reflection.MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    Shared Function Pasar_Segundos_a_Horas(Segundos As Long,
                                           Optional bMostrarSegundos As Boolean = False) As String
        Try
            Dim iMinutos As Long = 0
            Dim iHoras As Long = 0
            Dim iSegundos As Long = 0

            iHoras = Segundos \ SegundosHoras
            iMinutos = (Segundos Mod SegundosHoras) \ 60
            iSegundos = (iSegundos Mod SegundosHoras) Mod 60

            Pasar_Segundos_a_Horas = CStr(iHoras & ":" & Format(iMinutos, "00"))

            If bMostrarSegundos = True Then
                Pasar_Segundos_a_Horas = CStr(Pasar_Segundos_a_Horas & ":" & Format(iSegundos, "00"))
            End If

        Catch ex As Exception
            Pasar_Segundos_a_Horas = "00:00:00"
        End Try
    End Function


    'Public Shared Function DameEntregasProvisionales(ByVal FechaInicio As Date,
    '                                                 ByVal FechaFin As Date,
    '                                                 Optional ByVal Estado As Integer = EstadoEntregaPrevista.Desconocido,
    '                                                 Optional ByVal Articulo As String = "",
    '                                                 Optional ByVal TipoEmbalaje As Integer = 0) As List(Of EntregasPrevistas)
    '    Try
    '        Dim DTDatos As New DataTable
    '        Dim sSql As String = "SELECT * " &
    '                             "FROM EntregasProvisionales " &
    '                             "WHERE epFechaPrev BETWEEN '" & FechaInicio & "' AND '" & FechaFin & "' "

    '        If Estado <> EstadoEntregaPrevista.Desconocido Then
    '            sSql &= " AND epEstado=" & Estado
    '        End If

    '        If Articulo <> "" Then
    '            sSql &= " AND epGrupoArt='" & Articulo.Trim & "'"
    '        End If

    '        If TipoEmbalaje <> 0 Then
    '            sSql &= " AND opTipoEnvasado=" & TipoEmbalaje
    '        End If


    '        If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
    '            DameEntregasProvisionales = (From elemento In DTDatos
    '                                         Select New EntregasPrevistas(CStr(NoNull(elemento.Item("epGrupoArt"), "A")).Trim,
    '                                                                      CInt(NoNull(elemento.Item("epToneladas"), "D")),
    '                                                                      CDate(NoNull(elemento.Item("epFechaPrev"), "DT")),
    '                                                                      CType(NoNull(elemento.Item("epEstado"), "D"), EstadoEntregaPrevista),
    '                                                                      CType(NoNull(elemento.Item("epTipoEnvasado"), "D"), TipoEmbalaje))).ToList
    '        Else
    '            DameEntregasProvisionales = New List(Of EntregasPrevistas)
    '        End If

    '    Catch ex As Exception
    '        DameEntregasProvisionales = New List(Of EntregasPrevistas)
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name &
    '                                                        "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
    '    End Try
    'End Function

    Public Function DamePalabrasIdioma() As Dictionary(Of String, String)
        Try
            Dim DTDatos As New DataTable

            Dim Ssql As String = "SELECT idContenido,[" & System.Threading.Thread.CurrentThread.CurrentUICulture.ToString.Trim & "]" &
                                 " FROM Idiomas "
            DamePalabrasIdioma = New Dictionary(Of String, String)

            If Datos.CGPL.DameDatosDT(Ssql, DTDatos) Then
                DamePalabrasIdioma = (From elemento In DTDatos
                                      Select New With {.Codigo = UTrim((elemento.Item("idContenido"))).Trim,
                                                      .Valor = CStr(NoNull(elemento.Item(System.Threading.Thread.CurrentThread.CurrentUICulture.Name), "A")).Trim}).ToDictionary(Function(p) p.Codigo, Function(p) p.Valor)

            End If

        Catch ex As Exception
            DamePalabrasIdioma = New Dictionary(Of String, String)
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & MethodInfo.GetCurrentMethod.DeclaringType.Name & "." & MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

#End Region

End Class
