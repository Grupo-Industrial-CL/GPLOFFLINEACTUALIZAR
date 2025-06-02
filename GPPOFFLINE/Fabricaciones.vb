Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP


Public Class Fabricaciones

#Region "Atributos"
    Private miMaterial As Material
    'Private miEquipo As Equipos
    Private miPuestoTrabajo As PuestosTrabajo
    Private miHojaRuta As HojaRuta
    'Private miParada As Paradas
    'Private miListaParadas As List(Of Paradas)
    'Private miEnvio As EnviosProveedor

    Private bCreado As Boolean

#End Region
#Region "Constructores"
    Private Sub InicializarVariables()
        Try
            CodigoFabricacion = 0
            CodigoPuestoTrabajo = 0
            EnMarcha = 0
            OrdenMaq = 0
            CodigoMaterial = String.Empty
            CantidadPlanificada = 0
            CantidadFabricada = 0
            TfechaIni = Nothing
            Turno = New Char()
            FechaInicio = Nothing
            FechaFin = Nothing
            OrdenFabSAP = 0
            OrdenEnvSAP = 0
            FechaPreFin = Nothing
            CodigoListaMaterial = String.Empty
            GrupoHojaRuta = String.Empty
            ContadorHojaRuta = String.Empty
            SigPtoTrabajo = 0
            CantidadFabSAP = 0
            NumeroLoteSAP = String.Empty
            CodEquipo = 0
            PedidoSAP = ""
            PosPedidoSAP = 0
            Formato = String.Empty
            'miMaterial = New Material
            'miEquipo = New Equipos
            'miPuestoTrabajo = New PuestosTrabajo
            'miHojaRuta = New HojaRuta
            MaterialPadre = String.Empty
            CantidadPlanificadaPadre = 0
            CantidadFabRechazada = 0
            CantidadFabBuenas = 0
            CantidadReprocTap = 0
            CantidadReprocEti = 0
            MinutosFabObj = 0
            MinutosFabReal = 0
            NombreMaterial = ""
            IdEnvioProveedor = 0

            Me.bCreado = False
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal iCodigoFabricacion As Integer,
                  ByVal iCodigoPuestoTrabajo As Integer,
                  ByVal iEnMarcha As Byte,
                  ByVal iOrdernMaq As Integer,
                  ByVal sCodigoMaterial As String,
                  ByVal iCantidadPlanificada As Integer,
                  ByVal iCantidadFabricada As Integer,
                  ByVal dTfechaIni As Date,
                  ByVal cTurno As Char,
                  ByVal dFechaInicio As Date,
                  ByVal dFechaFin As Date,
                  ByVal iOrdenFabSAP As Integer,
                  ByVal iOrdernEnvSAP As Integer,
                  ByVal dFechaPreFin As Date,
                  ByVal sCodigoListaMaterial As String,
                  ByVal sGrupoHojaRuta As String,
                  ByVal sContadorHojaRuta As String,
                  ByVal iSigPtoTrabajo As Integer,
                  ByVal iCantFabSAP As Integer,
                  ByVal sNumeroLoteSAP As String,
                  ByVal iCodigoEquipo As Integer,
                  ByVal Pedido_SAP As String,
                  ByVal PosPedisoSAP As Integer,
                  ByVal sFormato As String,
                  ByVal sMaterialPadre As String,
                  ByVal iCantidadPlanificadaPadre As Integer,
                  ByVal iCantidadFabRechazadas As Integer,
                  ByVal iCantidadFabBuenas As Integer,
                  ByVal Cantidad_ReprocTap As Integer,
                  ByVal Cantidad_ReprocEti As Integer,
                  ByVal Minutos_FabObj As Integer,
                  ByVal Minutos_FabReal As Integer,
                  ByVal Nombre_Material As String,
                  ByVal IdEnvio_proveedor As Integer,
                   ByVal iEsLiberada As String,
                   ByVal iMensajeError As String,
                   ByVal itieneError As String,
                   ByVal iFaltanteOrdenEnvasadoSAP As String,
                   ByVal iFaltanteOrdenFabricacionSAP As String,
                   ByVal iopComentarioUsuario As String,
                   ByVal iProcedencia As String,
                   ByVal iCodGranel As String,
                   ByVal dFechaFuturoInicio As Date,
                   ByVal dFechaFuturoFin As Date,
                   ByVal iEsLiberadaEnvasado As String,
                   ByVal iUnidadesPorCaja As Integer)
        Try
            InicializarVariables()


            Dim sSql As String = ""
            Dim iFechaRotura = ConstantesGPP.FechaGlobal
            sSql = "select TOP 1 FechaRotura from PullSystemOFFLINE " &
                    " where CodigoMaterial = '" & sMaterialPadre.Trim() & "' and Version is not null order by Año, Mes desc "
            Dim DTDatos As New DataTable
            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Dim elemento = DTDatos.Rows(0)
                iFechaRotura = CDate(NoNull(elemento.Item("FechaRotura"), "DT"))
            End If


            CodigoFabricacion = iCodigoFabricacion
            CodigoPuestoTrabajo = iCodigoPuestoTrabajo
            EnMarcha = iEnMarcha
            OrdenMaq = iOrdernMaq
            CodigoMaterial = sCodigoMaterial
            CantidadPlanificada = iCantidadPlanificada
            CantidadFabricada = iCantidadFabricada
            TfechaIni = dTfechaIni
            Turno = cTurno
            FechaInicio = dFechaInicio
            FechaFin = dFechaFin
            OrdenFabSAP = iOrdenFabSAP
            OrdenEnvSAP = iOrdernEnvSAP
            FechaPreFin = dFechaPreFin
            CodigoListaMaterial = sCodigoListaMaterial
            GrupoHojaRuta = sGrupoHojaRuta
            ContadorHojaRuta = sContadorHojaRuta
            SigPtoTrabajo = iSigPtoTrabajo
            CantidadFabSAP = iCantFabSAP
            NumeroLoteSAP = sNumeroLoteSAP
            CodEquipo = CShort(iCodigoEquipo)
            PedidoSAP = Pedido_SAP
            PosPedidoSAP = PosPedisoSAP
            Formato = sFormato
            MaterialPadre = sMaterialPadre
            CantidadPlanificadaPadre = iCantidadPlanificadaPadre
            CantidadFabRechazada = iCantidadFabRechazadas
            CantidadFabBuenas = iCantidadFabBuenas
            CantidadReprocTap = Cantidad_ReprocTap
            CantidadReprocEti = Cantidad_ReprocEti
            MinutosFabObj = Minutos_FabObj
            MinutosFabReal = Minutos_FabReal
            NombreMaterial = Nombre_Material
            IdEnvioProveedor = IdEnvio_proveedor

            EsLiberadaEnvasado = iEsLiberadaEnvasado
            EsLiberada = iEsLiberada
            MensajeError = iMensajeError
            tieneError = itieneError
            FechaRotura = iFechaRotura

            FaltanteOrdenEnvasadoSAP = iFaltanteOrdenEnvasadoSAP
            FaltanteOrdenFabricacionSAP = iFaltanteOrdenFabricacionSAP
            opComentarioUsuario = iopComentarioUsuario
            Procedencia = iProcedencia
            CodGranel = iCodGranel

            FechaFuturoInicio = dFechaFuturoInicio
            FechaFuturoFin = dFechaFuturoFin

            UnidadesPorCaja = iUnidadesPorCaja

            Me.bCreado = True
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal iCodigoFabricacion As Integer,
                  ByVal iCodigoPuestoTrabajo As Integer,
                  ByVal iEnMarcha As Byte,
                  ByVal iOrdernMaq As Integer,
                  ByVal sCodigoMaterial As String,
                  ByVal iCantidadPlanificada As Integer,
                  ByVal iCantidadFabricada As Integer,
                  ByVal dTfechaIni As Date,
                  ByVal cTurno As Char,
                  ByVal dFechaInicio As Date,
                  ByVal dFechaFin As Date,
                  ByVal iOrdenFabSAP As Integer,
                  ByVal iOrdernEnvSAP As Integer,
                  ByVal dFechaPreFin As Date,
                  ByVal sCodigoListaMaterial As String,
                  ByVal sGrupoHojaRuta As String,
                  ByVal sContadorHojaRuta As String,
                  ByVal iSigPtoTrabajo As Integer,
                  ByVal iCantFabSAP As Integer,
                  ByVal sNumeroLoteSAP As String,
                  ByVal iCodigoEquipo As Short,
                  ByVal Pedido_SAP As String,
                  ByVal PosPedisoSAP As Integer,
                  ByVal sFormato As String,
                  ByVal sMaterialPadre As String,
                  ByVal iCantidadPlanificadaPadre As Integer,
                  ByVal iCantidadFabRechazadas As Integer,
                  ByVal iCantidadFabBuenas As Integer,
                  ByVal Cantidad_ReprocTap As Integer,
                  ByVal Cantidad_ReprocEti As Integer,
                  ByVal Minutos_FabObj As Integer,
                  ByVal Minutos_FabReal As Integer,
                  ByVal Nombre_Material As String,
                  ByVal IdEnvio_proveedor As Integer)
        Try
            InicializarVariables()

            CodigoFabricacion = iCodigoFabricacion
            CodigoPuestoTrabajo = iCodigoPuestoTrabajo
            EnMarcha = iEnMarcha
            OrdenMaq = iOrdernMaq
            CodigoMaterial = sCodigoMaterial
            CantidadPlanificada = iCantidadPlanificada
            CantidadFabricada = iCantidadFabricada
            TfechaIni = dTfechaIni
            Turno = cTurno
            FechaInicio = dFechaInicio
            FechaFin = dFechaFin
            OrdenFabSAP = iOrdenFabSAP
            OrdenEnvSAP = iOrdernEnvSAP
            FechaPreFin = dFechaPreFin
            CodigoListaMaterial = sCodigoListaMaterial
            GrupoHojaRuta = sGrupoHojaRuta
            ContadorHojaRuta = sContadorHojaRuta
            SigPtoTrabajo = iSigPtoTrabajo
            CantidadFabSAP = iCantFabSAP
            NumeroLoteSAP = sNumeroLoteSAP
            CodEquipo = iCodigoEquipo
            PedidoSAP = Pedido_SAP
            PosPedidoSAP = PosPedisoSAP
            Formato = sFormato
            MaterialPadre = sMaterialPadre
            CantidadPlanificadaPadre = iCantidadPlanificadaPadre
            CantidadFabRechazada = iCantidadFabRechazadas
            CantidadFabBuenas = iCantidadFabBuenas
            CantidadReprocTap = Cantidad_ReprocTap
            CantidadReprocEti = Cantidad_ReprocEti
            MinutosFabObj = Minutos_FabObj
            MinutosFabReal = Minutos_FabReal
            NombreMaterial = Nombre_Material
            IdEnvioProveedor = IdEnvio_proveedor

            Me.bCreado = True
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal iCodigoFabricacion As Integer,
                  ByVal iCodigoPuestoTrabajo As Integer,
                  ByVal iEnMarcha As Byte,
                  ByVal iOrdernMaq As Integer,
                  ByVal sCodigoMaterial As String,
                  ByVal iCantidadPlanificada As Integer,
                  ByVal iCantidadFabricada As Integer,
                  ByVal dTfechaIni As Date,
                  ByVal cTurno As Char,
                  ByVal dFechaInicio As Date,
                  ByVal dFechaFin As Date,
                  ByVal iOrdenFabSAP As Integer,
                  ByVal iOrdernEnvSAP As Integer,
                  ByVal dFechaPreFin As Date,
                  ByVal sCodigoListaMaterial As String,
                  ByVal sGrupoHojaRuta As String,
                  ByVal sContadorHojaRuta As String,
                  ByVal iSigPtoTrabajo As Integer,
                  ByVal iCantFabSAP As Integer,
                  ByVal sNumeroLoteSAP As String,
                  ByVal iCodigoEquipo As Short,
                  ByVal Pedido_SAP As String,
                  ByVal PosPedisoSAP As Integer,
                  ByVal sFormato As String,
                  ByVal sMaterialPadre As String,
                  ByVal iCantidadPlanificadaPadre As Integer,
                  ByVal iCantidadFabRechazadas As Integer,
                  ByVal iCantidadFabBuenas As Integer,
                  ByVal Cantidad_ReprocTap As Integer,
                  ByVal Cantidad_ReprocEti As Integer,
                  ByVal Minutos_FabObj As Integer,
                  ByVal Minutos_FabReal As Integer,
                  ByVal Nombre_Material As String,
                  ByVal IdEnvio_proveedor As Integer,
                   ByVal iEsLiberada As String,
                   ByVal iMensajeError As String,
                   ByVal itieneError As String)
        Try
            InicializarVariables()

            CodigoFabricacion = iCodigoFabricacion
            CodigoPuestoTrabajo = iCodigoPuestoTrabajo
            EnMarcha = iEnMarcha
            OrdenMaq = iOrdernMaq
            CodigoMaterial = sCodigoMaterial
            CantidadPlanificada = iCantidadPlanificada
            CantidadFabricada = iCantidadFabricada
            TfechaIni = dTfechaIni
            Turno = cTurno
            FechaInicio = dFechaInicio
            FechaFin = dFechaFin
            OrdenFabSAP = iOrdenFabSAP
            OrdenEnvSAP = iOrdernEnvSAP
            FechaPreFin = dFechaPreFin
            CodigoListaMaterial = sCodigoListaMaterial
            GrupoHojaRuta = sGrupoHojaRuta
            ContadorHojaRuta = sContadorHojaRuta
            SigPtoTrabajo = iSigPtoTrabajo
            CantidadFabSAP = iCantFabSAP
            NumeroLoteSAP = sNumeroLoteSAP
            CodEquipo = iCodigoEquipo
            PedidoSAP = Pedido_SAP
            PosPedidoSAP = PosPedisoSAP
            Formato = sFormato
            MaterialPadre = sMaterialPadre
            CantidadPlanificadaPadre = iCantidadPlanificadaPadre
            CantidadFabRechazada = iCantidadFabRechazadas
            CantidadFabBuenas = iCantidadFabBuenas
            CantidadReprocTap = Cantidad_ReprocTap
            CantidadReprocEti = Cantidad_ReprocEti
            MinutosFabObj = Minutos_FabObj
            MinutosFabReal = Minutos_FabReal
            NombreMaterial = Nombre_Material
            IdEnvioProveedor = IdEnvio_proveedor

            EsLiberada = iEsLiberada
            MensajeError = iMensajeError
            tieneError = itieneError

            Me.bCreado = True
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal IdFab As Integer)
        Try
            InicializarVariables()
            Dim sSql As String = "SELECT * FROM dbo.Fabricaciones " &
                                 "WHERE opIdFab = " & IdFab.ToString
            Dim DTDatos As New DataTable
            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Dim elemento = DTDatos.Rows(0)

                CodigoFabricacion = CInt(NoNull(elemento.Item("opIdFab"), "D"))
                CodigoPuestoTrabajo = CInt(NoNull(elemento.Item("opPuestoTrabajo"), "D"))
                EnMarcha = CByte(NoNull(elemento.Item("opEnmarcha"), "D"))
                OrdenMaq = CInt(NoNull(elemento.Item("opOrdenMaq"), "D"))
                CodigoMaterial = CStr(NoNull(elemento.Item("opMaterial"), "A"))
                CantidadPlanificada = CInt(NoNull(elemento.Item("opCantidadPlanif"), "D"))
                CantidadFabricada = CInt(NoNull(elemento.Item("opCantidadFab"), "D"))
                TfechaIni = CDate(NoNull(elemento.Item("opTfecIni"), "DT"))
                Turno = CChar(Trim(CStr(NoNull(elemento.Item("opTurno"), "A"))))
                FechaInicio = CDate(NoNull(elemento.Item("opFechaIni"), "DT"))
                FechaFin = CDate(NoNull(elemento.Item("opFechaFin"), "DT"))
                OrdenFabSAP = CInt(NoNull(elemento.Item("opOrdenFabSAP"), "D"))
                OrdenEnvSAP = CInt(NoNull(elemento.Item("opOrdenEnvSAP"), "D"))
                FechaPreFin = CDate(NoNull(elemento.Item("opFechaPrevFin"), "DT"))
                CodigoListaMaterial = CStr(NoNull(elemento.Item("opListaMaterial"), "A"))
                GrupoHojaRuta = CStr(NoNull(elemento.Item("opGrupoHR"), "A"))
                ContadorHojaRuta = CStr(NoNull(elemento.Item("opContHR"), "A"))
                CantidadFabSAP = CInt(NoNull(elemento.Item("opCantidadFabSAP"), "D"))
                NumeroLoteSAP = CStr(NoNull(elemento.Item("opNumeroLoteSAP"), "A"))
                SigPtoTrabajo = CInt(NoNull(elemento.Item("opSigPtoTrabajo"), "D"))
                CodEquipo = CShort(NoNull(elemento.Item("opEquipo"), "D"))
                PedidoSAP = CStr(NoNull(elemento.Item("opNumPedSAP"), "A"))
                PosPedidoSAP = CInt(NoNull(elemento.Item("opPosPedSAP"), "D"))
                Formato = CStr(NoNull(elemento.Item("opFormato"), "A"))
                MaterialPadre = CStr(NoNull(elemento.Item("opMaterialPadre"), "A"))
                CantidadPlanificadaPadre = CInt(NoNull(elemento.Item("opCantidadPlanifPadre"), "D"))
                CantidadFabRechazada = CInt(NoNull(elemento.Item("opCantidadFabRechazada"), "D"))
                CantidadFabBuenas = CInt(NoNull(elemento.Item("opCantidadFabBuenas"), "D"))

                CantidadReprocTap = CInt(NoNull(elemento.Item("opCantidadReprocTap"), "D"))
                CantidadReprocEti = CInt(NoNull(elemento.Item("opCantidadReprocEti"), "D"))
                MinutosFabObj = CInt(NoNull(elemento.Item("opMinutosFabObj"), "D"))
                MinutosFabReal = CInt(NoNull(elemento.Item("opMinutosFabReal"), "D"))
                NombreMaterial = CStr(NoNull(elemento.Item("opNombreMaterial"), "A"))
                IdEnvioProveedor = CInt(NoNull(elemento.Item("opIdEnvio"), "D"))

                Me.bCreado = True
            End If
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New()
        Try
            InicializarVariables()
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

#End Region


#Region "Propiedades"

    Public Property CodigoFabricacion As Integer
    Public Property CodigoPuestoTrabajo As Integer
    Public Property EnMarcha As Byte
    Public Property OrdenMaq As Integer
    Public Property CodigoMaterial As String
    Public Property CantidadPlanificada As Integer
    Public Property CantidadFabricada As Integer = 0
    Public Property TfechaIni As Date
    Public Property Turno As Char
    Public Property FechaInicio As Date
    Public Property FechaFin As Date
    Public Property OrdenFabSAP As Integer
    Public Property OrdenEnvSAP As Integer
    Public Property FechaPreFin As Date
    Public Property CodigoListaMaterial As String
    Public Property GrupoHojaRuta As String
    Public Property ContadorHojaRuta As String
    Public Property SigPtoTrabajo As Integer
    Public Property CantidadFabSAP As Integer
    Public Property NumeroLoteSAP As String
    Public Property PedidoSAP As String
    Public Property PosPedidoSAP As Integer

    Public Property CodEquipo As Short

    Public Property Formato As String

    Public Property MaterialPadre As String

    Public Property CantidadPlanificadaPadre As Integer

    Public Property CantidadFabRechazada As Integer

    Public Property CantidadFabBuenas As Integer
    Public Property CantidadReprocTap As Integer
    Public Property CantidadReprocEti As Integer
    Public Property MinutosFabObj As Integer

    Public Property NombreMaterial As String

    Public Property MinutosFabReal As Integer
    Public Property IdEnvioProveedor As Integer

    Public Property EstadoDisponibilidad As EstadosDisponibilidad

    Public Property EsLiberada As String = "0"
    Public Property MensajeError As String = ""

    Public Property tieneError As String = ""

    Public Property EsLiberadaEnvasado As String = "0"

    Public Property FechaRotura As Date

    Public Property FaltanteOrdenEnvasadoSAP As String = ""

    Public Property FaltanteOrdenFabricacionSAP As String = ""

    Public Property opComentarioUsuario As String = ""

    Public Property Procedencia As String = ""

    Public Property CodGranel As String = ""

    Public Property FechaFuturoInicio As DateTime = FechaGlobal
    Public Property FechaFuturoFin As DateTime = FechaGlobal

    Public Property UnidadesPorCaja As Integer = 1

    Public ReadOnly Property Minutos_Tiempo_Preparacion(bIncluirPreparacion As Boolean) As Integer
        Get
            '2023-04-25 nuevos parametros para calcular el Tiempo de Fabricación
            '2023-04-25 
            'Return Dame_Minutos_Tiempo_Fabricacion(Cantidad:=Cantidad_Pte_Fabricar,
            '                                       Hoja_de_Ruta:=Hoja_Ruta,
            '                                       Puestotrabajo:=CodigoPuestoTrabajo,
            '                                       IncluirPreparacion:=bIncluirPreparacion)

            Return Dame_Minutos_Tiempo_Preparacion(Hoja_de_Ruta:=Hoja_Ruta,
                                                  Puestotrabajo:=CodigoPuestoTrabajo,
                                                  IncluirPreparacion:=bIncluirPreparacion, GrupoHojaRuta:=GrupoHojaRuta, ContadorHojaRuta:=ContadorHojaRuta)
        End Get
    End Property

    Public ReadOnly Property Minutos_Tiempo_FabricacionEnMarcha(bIncluirPreparacion As Boolean) As Integer
        Get
            '2023-04-25 nuevos parametros para calcular el Tiempo de Fabricación
            '2023-04-25 
            'Return Dame_Minutos_Tiempo_Fabricacion(Cantidad:=Cantidad_Pte_Fabricar,
            '                                       Hoja_de_Ruta:=Hoja_Ruta,
            '                                       Puestotrabajo:=CodigoPuestoTrabajo,
            '                                       IncluirPreparacion:=bIncluirPreparacion)

            Return Dame_Minutos_Tiempo_FabricacionEnMarcha(Cantidad:=Cantidad_Pte_Fabricar,
                                                  Hoja_de_Ruta:=Hoja_Ruta,
                                                  Puestotrabajo:=CodigoPuestoTrabajo,
                                                  IncluirPreparacion:=bIncluirPreparacion, GrupoHojaRuta:=GrupoHojaRuta, ContadorHojaRuta:=ContadorHojaRuta)
        End Get
    End Property
    Public ReadOnly Property Material As Material
        Get
            If miMaterial Is Nothing Then
                miMaterial = New Material(CodigoMaterial)
            ElseIf miMaterial.Creado = False Then
                miMaterial = New Material(CodigoMaterial)
            End If

            Return miMaterial
        End Get
    End Property

    Public ReadOnly Property MaterialPadreDetalle As Material
        Get
            If miMaterial Is Nothing Then
                miMaterial = New Material(MaterialPadre)
            ElseIf miMaterial.Creado = False Then
                miMaterial = New Material(MaterialPadre)
            End If

            Return miMaterial
        End Get
    End Property



    'Public ReadOnly Property Equipo As Equipos
    '    Get
    '        If miEquipo Is Nothing Then
    '            miEquipo = New Equipos(Codigo:=CodEquipo)
    '        ElseIf miEquipo.Creado = False Then
    '            miEquipo = New Equipos(Codigo:=CodEquipo)
    '        End If

    '        Return miEquipo
    '    End Get
    'End Property

    Public ReadOnly Property PuestoTrabajo As PuestosTrabajo
        Get
            If miPuestoTrabajo Is Nothing Then
                miPuestoTrabajo = New PuestosTrabajo(CodigoPuestoTrabajo)
            ElseIf miPuestoTrabajo.Creado = False Then
                miPuestoTrabajo = New PuestosTrabajo(CodigoPuestoTrabajo)
            End If

            Return miPuestoTrabajo
        End Get
    End Property

    Public ReadOnly Property Hoja_Ruta As HojaRuta
        Get
            If miHojaRuta Is Nothing Then
                miHojaRuta = New HojaRuta(GrupoHojaruta:=GrupoHojaRuta,
                                          ContadorGrupo:=ContadorHojaRuta)
            ElseIf miHojaRuta.Creado = False Then
                miHojaRuta = New HojaRuta(GrupoHojaruta:=GrupoHojaRuta,
                                          ContadorGrupo:=ContadorHojaRuta)
            End If

            Return miHojaRuta
        End Get
    End Property

    Public ReadOnly Property Operarios As Integer
        Get
            Dim iOperarios As Integer = 0

            For Each miOper In Hoja_Ruta.OperacHojaRutaLista
                If miOper.CodigoPuestoDeTrabajo = Me.CodigoPuestoTrabajo AndAlso miOper.Operarios > iOperarios Then
                    iOperarios = miOper.Operarios
                End If
            Next

            Return iOperarios
        End Get
    End Property

    ''' <summary>
    ''' Solo para el pedido en MRCHA. Fecha de fin prevista en base a los turnos creados y la planificación de fabricaciones- Dato calculado no guardado en base de datos
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Fecha_Fin_Previsto_PedidoenMarcha() As Date
        Get
            If Me.EnMarcha = EstadoFabricacion.EnMarcha Then

                Return DatosProduccion.DameHoraFin(SegundosFab:=Minutos_Para_Cambio * 60,
                                                  FechaInicio:=CDate(IIf(FechaInicio = FechaGlobal,
                                                                         PuestoTrabajo.Proxima_FechaInicio_Turno,
                                                                         Now)),
                                                  CodPuestoTrabajo:=Me.CodigoPuestoTrabajo)
            Else
                Return FechaGlobal
            End If
        End Get
    End Property

    Public ReadOnly Property Velocidad_Media_Actual As Double
        Get

            If Me.EnMarcha = EstadoFabricacion.EnMarcha Or
               Me.EnMarcha = EstadoFabricacion.Finalizada Then
                If Minutos_Total > 0 Then
                    Return Me.CantidadFabricada / Minutos_Total
                End If
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property Minutos_Total As Long
        Get
            If Me.EnMarcha = EstadoFabricacion.EnMarcha Or
               Me.EnMarcha = EstadoFabricacion.Finalizada Then
                Return DateDiff(DateInterval.Minute, CDate(IIf(Me.FechaInicio = FechaGlobal, PuestoTrabajo.Proxima_FechaInicio_Turno, FechaInicio)), Now)
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property Minutos_Para_Cambio(Optional bCambioFormato As Boolean = False) As Integer
        Get
            If Math.Round(Velocidad_Media_Actual) > 0 Then
                Return CInt(Cantidad_Pte_Fabricar / Velocidad_Media_Actual)
            Else
                Return Minutos_Tiempo_Fabricacion(bCambioFormato)
            End If
        End Get
    End Property

    Public ReadOnly Property Cantidad_Pte_Fabricar As Integer
        Get
            If EnMarcha = ConstantesGPP.EstadoFabricacion.PteFabricar Or
                 EnMarcha = ConstantesGPP.EstadoFabricacion.EnMarcha Then
                If CantidadPlanificada - CantidadFabricada > 0 Then
                    Return CantidadPlanificada - CantidadFabricada
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        End Get
    End Property







    Public ReadOnly Property Creado As Boolean
        Get
            Creado = bCreado
        End Get
    End Property

    Public ReadOnly Property Minutos_Tiempo_Fabricacion(bIncluirPreparacion As Boolean) As Integer
        Get
            '2023-04-25 nuevos parametros para calcular el Tiempo de Fabricación
            '2023-04-25 
            'Return Dame_Minutos_Tiempo_Fabricacion(Cantidad:=Cantidad_Pte_Fabricar,
            '                                       Hoja_de_Ruta:=Hoja_Ruta,
            '                                       Puestotrabajo:=CodigoPuestoTrabajo,
            '                                       IncluirPreparacion:=bIncluirPreparacion)

            Return Dame_Minutos_Tiempo_Fabricacion(Cantidad:=Cantidad_Pte_Fabricar,
                                                  Hoja_de_Ruta:=Hoja_Ruta,
                                                  Puestotrabajo:=CodigoPuestoTrabajo,
                                                  IncluirPreparacion:=bIncluirPreparacion, GrupoHojaRuta:=GrupoHojaRuta, ContadorHojaRuta:=ContadorHojaRuta)
        End Get
    End Property

    'Public ReadOnly Property Minutos_Tiempo_Fabricacion(bIncluirPreparacion As Boolean, Cantidad_Pte As Integer, CodPuesto As Integer) As Integer
    '    Get

    '        Return Dame_Minutos_Tiempo_Fabricacion(Cantidad:=Cantidad_Pte,
    '                                               Hoja_de_Ruta:=Hoja_Ruta,
    '                                               Puestotrabajo:=CodPuesto,
    '                                               IncluirPreparacion:=bIncluirPreparacion)
    '    End Get
    'End Property

#End Region

#Region "BBDD"

    Public Function Insertar() As Boolean
        Try
            If OrdenMaq = 0 Then
                OrdenMaq = DameSiguienteOrden()
            End If

            Dim sSql As String = "INSERT INTO Fabricaciones (opPuestoTrabajo, opEnmarcha, opOrdenMaq, opMaterial,opCantidadPlanif, opCantidadFab, opTFecIni, opTurno, opFechaIni," &
                                                            "opFechaFin, opOrdenFabSAP, opOrdenEnvSAP,opFechaPrevFin,opListaMaterial,opGrupoHR,opContHR," &
                                                            "opCantidadFabSAP,opNumeroLoteSAP,opSigPtoTrabajo,opEquipo,opNumPedSAP,opPosPedSAP,opFormato,opMaterialPadre,opCantidadPlanifPadre," &
                                                            "opCantidadFabRechazada,opCantidadReprocTap,opCantidadReprocEti,opMinutosFabObj,opNombreMaterial,opIdEnvio) VALUES (" &
                                                                CodigoPuestoTrabajo & ", " &
                                                                EnMarcha & ", " &
                                                                OrdenMaq & ", '" &
                                                                CodigoMaterial & "', " &
                                                                CantidadPlanificada & ", " &
                                                                CantidadFabricada & ", '" &
                                                                TfechaIni & "', '" &
                                                                Turno & "', '" &
                                                                FechaInicio & "', '" &
                                                                FechaFin & "', " &
                                                                OrdenFabSAP & ", " &
                                                                OrdenEnvSAP & ", '" &
                                                                FechaPreFin & "', '" &
                                                                CodigoListaMaterial & "', '" &
                                                                GrupoHojaRuta & "', '" &
                                                                ContadorHojaRuta & "', " &
                                                                CantidadFabSAP & ", '" &
                                                                NumeroLoteSAP & "', " &
                                                                SigPtoTrabajo & "," &
                                                                CodEquipo & ",'" &
                                                                PedidoSAP.Trim & "'," &
                                                                PosPedidoSAP & ", '" &
                                                                Formato.Trim & "','" &
                                                                MaterialPadre.Trim & "'," &
                                                                CantidadPlanificadaPadre & "," &
                                                                CantidadFabRechazada & "," &
                                                                CantidadReprocTap & "," &
                                                                CantidadReprocEti & "," &
                                                                MinutosFabObj & ",'" &
                                                                NombreMaterial.Trim & "'," &
                                                                IdEnvioProveedor & ") Select @@IDENTITY "

            Me.CodigoFabricacion = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If Me.CodigoFabricacion = -1 Then
                Insertar = False
            Else
                Insertar = True
                ' si insert de la fabricacion es correcto se valida si el material es de centro extero
            End If
            Datos.GuardarLog(TipoLogDescripcion.Alta & " " & Me.GetType().Name, CodigoFabricacion.ToString)
        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function





    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE Fabricaciones " &
                                 " Set opPuestoTrabajo = " & CodigoPuestoTrabajo & ", " &
                                 " opEnmarcha = " & EnMarcha & ", " &
                                 " opOrdenMaq = " & OrdenMaq & ", " &
                                 " opMaterial = '" & CodigoMaterial.Trim() & "', " &
                                 " opCantidadPlanif = " & CantidadPlanificada & ", " &
                                 " opCantidadFab = " & CantidadFabricada & ", " &
                                 " opTFecIni = '" & TfechaIni & "', " &
                                 " opTurno = '" & Turno & "', " &
                                 " opFechaIni = '" & FechaInicio & "', " &
                                 " opFechaFin = '" & FechaFin & "', " &
                                 " opOrdenFabSAP = " & OrdenFabSAP & ", " &
                                 " opOrdenEnvSAP = " & OrdenEnvSAP & ", " &
                                 " opFechaPrevFin = '" & FechaPreFin & "', " &
                                 " opListaMaterial = '" & CodigoListaMaterial & "', " &
                                 " opGrupoHR = '" & GrupoHojaRuta & "', " &
                                 " opContHR = '" & ContadorHojaRuta & "', " &
                                 " opSigPtoTrabajo = " & SigPtoTrabajo & ", " &
                                 " opCantidadFabSAP = " & CantidadFabSAP & ", " &
                                 " opNumeroLoteSAP = '" & NumeroLoteSAP.Trim() & "'," &
                                 " opNumPedSAP = '" & PedidoSAP.Trim & "'," &
                                 " opPosPedSAP = " & PosPedidoSAP & "," &
                                 " opEquipo = " & CodEquipo & "," &
                                 " opFormato = '" & Formato & "', " &
                                 " opMaterialPadre = '" & MaterialPadre.Trim() & "', " &
                                 " opCantidadPlanifPadre = " & CantidadPlanificadaPadre & "," &
                                 " opCantidadFabRechazada = " & CantidadFabRechazada & "," &
                                 " opCantidadReprocTap = " & CantidadReprocTap & "," &
                                 " opCantidadReprocEti = " & CantidadReprocEti & "," &
                                 " opMinutosFabObj = " & MinutosFabObj & "," &
                                 " opNombreMaterial = '" & NombreMaterial.Trim & "'," &
                                 " opIdEnvio = " & IdEnvioProveedor & "," &
                                    " opEsLiberada = '" & EsLiberada.Trim() & "', " &
                                    " opMensajeError = '" & MensajeError.Trim() & "', " &
                                    " opTieneError = '" & tieneError.Trim() & "' " &
                                 " WHERE opIdFab=" & CodigoFabricacion

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " Fabricaciones", CStr(CodigoFabricacion))
            End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM Fabricaciones " &
                                 "WHERE opIdFab=" & CodigoFabricacion

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Fabricaciones", CStr(CodigoFabricacion))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function DameSiguienteOrden() As Integer
        Try
            Dim sSql As String = "Select opOrdenMaq FROM Fabricaciones " &
                                 "WHERE opEnmarcha='" & EstadoFabricacion.PteFabricar & "' " &
                                 "ORDER BY opOrdenMaq DESC "
            Dim DTDatos As New DataTable
            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Return CInt(NoNull(DTDatos.Rows(0).Item("opOrdenMaq"), "D")) + 100
            Else
                Return 1
            End If
        Catch ex As Exception
            DameSiguienteOrden = 1
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

#End Region

End Class
