
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP


Public Class Material

#Region "Atributos"

    Private sCodigo As String
    Private sTipo As String
    Private sGrupo As String
    Private miGrupo As GrupoArticulo
    Private sUnidadMedida As String
    Private sNombre As String
    Private sListaMat As String
    Private miCabListaMaterial As CabListaMaterial
    Private iFamiliaEnvasado As Byte
    Private dFechaIniPS As Date
    Private dFechaFinPS As Date
    Private iDiasPP As Byte
    Private iStockMaxPS As Integer
    Private iStockMinPS As Integer
    Private misHojasRuta As List(Of HojaRuta)
    Private miHojaRutaDefecto As HojaRuta
    Private bActivo As Boolean
    Private iLoteMinimo As Integer
    Private iLoteMaximo As Integer
    Private iLoteFijo As Integer
    Private iRedondeoLote As Integer
    Private iDiasFabPropia As Integer
    Private sTipoTamañoLote As String
    Private sGrupoHR As String
    Private sContadorHR As String
    Private sGrupoCompra As String
    Private miGrupoCompra As New GrupoCompras
    Private misProductoFormula As List(Of Material)
    Private misSemielaborados As List(Of Material)
    Private misPackaging As List(Of Material)
    Private iStockActual As Integer
    Private bMostrarInformes As Boolean
    Private iUnidadesCaja As Integer
    Private iUnidadesPalet As Integer
    Private iMesesLoteCarga As Integer
    Private bCreado As Boolean

#End Region

#Region "Constructores"

    ''' <summary>
    ''' Inicializa todos los atributos a sus valores por defecto.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InicializarVariables()
        Try
            sCodigo = ""
            sTipo = ""
            sGrupo = ""
            sUnidadMedida = ""
            sNombre = ""
            Me.dFechaFinPS = FechaGlobal
            Me.dFechaIniPS = FechaGlobal
            Me.iStockMaxPS = 0
            Me.iStockMinPS = 0
            Me.iDiasPP = 0
            Me.sListaMat = ""
            Me.iFamiliaEnvasado = 0
            Me.bActivo = True
            Me.iLoteMinimo = 0
            Me.iLoteMaximo = 0
            Me.iRedondeoLote = 0
            Me.sTipoTamañoLote = ""
            Me.iDiasFabPropia = 0
            Me.iLoteFijo = 0
            Me.sGrupoCompra = ""
            Me.sGrupoHR = ""
            Me.sContadorHR = ""
            Me.iStockActual = -999999999
            Me.bMostrarInformes = False
            Me.iUnidadesCaja = 0
            iUnidadesPalet = 0
            Me.miGrupo = Nothing
            Me.miCabListaMaterial = Nothing
            Me.misHojasRuta = Nothing
            Me.miHojaRutaDefecto = Nothing
            misProductoFormula = Nothing
            misSemielaborados = Nothing
            misPackaging = Nothing
            iMesesLoteCarga = 0
            bCreado = False

        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    ''' <summary>
    ''' Constructor vacio.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Try
            InicializarVariables()
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Constructo parametrizado.
    ''' </summary>
    ''' <param name="Codigo"></param>
    ''' <param name="Tipo"></param>
    ''' <param name="Grupo"></param>
    ''' <param name="UnidadMedida"></param>
    ''' <param name="Nombre"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Codigo As String,
                   ByVal Tipo As String,
                   ByVal Grupo As String,
                   ByVal UnidadMedida As String,
                   ByVal Nombre As String,
                   ByVal Lista_Mat As String,
                   ByVal Familia_Envasado As Byte,
                   ByVal Fecha_IniPS As Date,
                   ByVal Fecha_FinPS As Date,
                   ByVal Dias_PP As Byte,
                   ByVal Stock_MaxPS As Integer,
                   ByVal Stock_MinPS As Integer,
                   ByVal Activo As Boolean,
                   ByVal Lote_Minimo As Integer,
                   ByVal Lote_Maximo As Integer,
                   ByVal Lote_Fijo As Integer,
                   ByVal Redondeo_Lote As Integer,
                   ByVal Tipo_TamañoLote As String,
                   ByVal Dias_FabPropia As Integer,
                   ByVal Grupo_HojaRuta As String,
                   ByVal Contador_HojaRuta As String,
                   ByVal Grupo_Compra As String,
                   ByVal Mostrar_Informes As Boolean,
                   ByVal Unidades_Pack As Integer,
                   ByVal UnidadesPorPalet As Integer,
                   ByVal MesesLoteCarga As Integer)
        Try
            InicializarVariables()

            Me.sCodigo = Codigo
            Me.sTipo = Tipo
            Me.sGrupo = Grupo
            Me.sUnidadMedida = UnidadMedida
            Me.sNombre = Nombre
            Me.sListaMat = Lista_Mat
            Me.iFamiliaEnvasado = Familia_Envasado
            Me.dFechaIniPS = Fecha_IniPS
            Me.dFechaFinPS = Fecha_FinPS
            Me.iDiasPP = Dias_PP
            Me.iStockMaxPS = Stock_MaxPS
            Me.iStockMinPS = Stock_MinPS
            Me.sGrupoCompra = Grupo_Compra
            Me.iLoteMinimo = Lote_Minimo
            Me.iLoteMaximo = Lote_Maximo
            Me.iLoteFijo = Lote_Fijo
            Me.iRedondeoLote = Redondeo_Lote
            Me.sTipoTamañoLote = Tipo_TamañoLote
            Me.iDiasFabPropia = Dias_FabPropia
            Me.sGrupoHR = Grupo_HojaRuta
            Me.sContadorHR = Contador_HojaRuta
            Me.bMostrarInformes = Mostrar_Informes
            Me.iUnidadesCaja = Unidades_Pack
            Me.iUnidadesPalet = UnidadesPorPalet
            iMesesLoteCarga = MesesLoteCarga
            Me.bActivo = Activo

            Me.bCreado = True
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Constructor parametrizado contra Base de Datos.
    ''' </summary>
    ''' <param name="Codigo"></param>
    ''' <remarks></remarks>
    Public Sub New(Codigo As String)
        Try
            Dim sSql As String = "SELECT * " &
                                 " FROM Materiales " &
                                 " WHERE maCod = '" & Codigo & "'"
            Dim DTDatos As New DataTable

            InicializarVariables()

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Me.sCodigo = Codigo
                Me.sTipo = UTrim(DTDatos.Rows(0).Item("maTipoMat"))
                Me.sGrupo = UTrim(DTDatos.Rows(0).Item("maGrupoArt"))
                Me.sUnidadMedida = UTrim(DTDatos.Rows(0).Item("maUMBase"))
                Me.sNombre = UTrim(DTDatos.Rows(0).Item("maNombre"))
                Me.sListaMat = UTrim(DTDatos.Rows(0).Item("maListaMaterial"))
                Me.iFamiliaEnvasado = CByte(DTDatos.Rows(0).Item("maFamiliaEnvasado"))
                Me.dFechaIniPS = CDate(DTDatos.Rows(0).Item("maFecIniPS"))
                Me.dFechaFinPS = CDate(DTDatos.Rows(0).Item("maFecFinPS"))
                Me.iDiasPP = CByte(DTDatos.Rows(0).Item("maDiasPP"))
                Me.iStockMaxPS = CInt(DTDatos.Rows(0).Item("maStokMaxPS"))
                Me.iStockMinPS = CInt(DTDatos.Rows(0).Item("maStokMinPS"))

                Me.iLoteMinimo = CInt(DTDatos.Rows(0).Item("maLoteMin"))
                Me.iLoteMaximo = CInt(DTDatos.Rows(0).Item("maLoteMax"))
                Me.iLoteFijo = CInt(DTDatos.Rows(0).Item("maLoteFijo"))
                Me.iRedondeoLote = CInt(DTDatos.Rows(0).Item("maRedondeo"))
                Me.sTipoTamañoLote = CStr(DTDatos.Rows(0).Item("maTipoTamLote")).Trim
                Me.iDiasFabPropia = CInt(DTDatos.Rows(0).Item("maDiasFabPropia"))
                Me.sGrupoHR = CStr(NoNull(DTDatos.Rows(0).Item("maGrupoHR"), "A")).Trim
                Me.sContadorHR = CStr(NoNull(DTDatos.Rows(0).Item("maContHR"), "A")).Trim
                Me.sGrupoCompra = CStr(NoNull(DTDatos.Rows(0).Item("maGrupoCompra"), "A")).Trim
                Me.bActivo = CBool(DTDatos.Rows(0).Item("maActivo"))
                Me.bMostrarInformes = CBool(DTDatos.Rows(0).Item("mnMostrarInformes"))
                Me.iUnidadesCaja = CInt(DTDatos.Rows(0).Item("maUnidadesPACK"))
                Me.iUnidadesPalet = CInt(DTDatos.Rows(0).Item("maUnidadesPalet"))
                iMesesLoteCarga = CInt(DTDatos.Rows(0).Item("maMesesLoteCarga"))
                Me.bCreado = True
            End If
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

#End Region

#Region "BBDD"

    Public Function Insertar() As Boolean
        Try
            Dim sSql As String = "INSERT INTO Materiales (maCod,maNombre,maTipoMat,maGrupoArt,maUMBase,maFamiliaEnvasado,maFecIniPS,maFecFinPS,maDiasPP,maStokMinPS,maStokMaxPS," &
                                 "maListaMaterial,maActivo,maLoteFijo,maLoteMin,maLoteMax,maRedondeo,maDiasFabPropia,maTipoTamLote,maGrupoHR,maContHR,maGrupoCompra," &
                                 "mnMostrarInformes,maUnidadesPACK,maUnidadesPalet,maMesesLoteCarga) VALUES ('" &
                                             Me.sCodigo & "','" &
                                             Me.sNombre.Trim.ToUpper & "','" &
                                             Me.sTipo.Trim & "','" &
                                             Me.sGrupo.Trim & "','" &
                                             Me.sUnidadMedida.Trim & "'," &
                                             Me.iFamiliaEnvasado & ",'" &
                                             Me.dFechaIniPS & "','" &
                                             Me.dFechaFinPS & "'," &
                                             Me.iDiasPP & "," &
                                             Me.iStockMaxPS & "," &
                                             Me.iStockMinPS & ",'" &
                                             Me.sListaMat.Trim & "','" &
                                             Me.bActivo & "'," &
                                             Me.iLoteFijo & "," &
                                             Me.iLoteMinimo & "," &
                                             Me.iLoteMaximo & "," &
                                             Me.iRedondeoLote & "," &
                                             Me.iDiasFabPropia & ",'" &
                                             Me.sTipoTamañoLote.Trim & "','" &
                                             Me.sGrupoHR.Trim & "','" &
                                             Me.sContadorHR.Trim & "','" &
                                             Me.sGrupoCompra & "','" &
                                             Me.bMostrarInformes & "'," &
                                             Me.iUnidadesCaja & "," &
                                             Me.iUnidadesPalet & "," &
                                             Me.iMesesLoteCarga & ")"

            Insertar = Datos.CGPL.EjecutarConsulta(sSql)

            If Insertar Then
                Me.bCreado = Insertar
                Datos.GuardarLog(TipoLogDescripcion.Alta & " Materiales ", Me.sCodigo)
            End If
        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function


    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE materiales SET maNombre='" & Me.sNombre.Trim.ToUpper & "'," &
                                 "maTipoMat='" & Me.sTipo.Trim & "'," &
                                 "maGrupoArt='" & Me.sGrupo.Trim & "'," &
                                 "maUMBase='" & Me.sUnidadMedida.Trim & "'," &
                                 "maFamiliaEnvasado=" & Me.iFamiliaEnvasado & "," &
                                 "maFecIniPS='" & Me.dFechaIniPS.Date & "'," &
                                 "maFecFinPS='" & Me.dFechaFinPS.Date & "'," &
                                 "maDiasPP=" & Me.iDiasPP & "," &
                                 "maStokMinPS=" & Me.iStockMinPS & "," &
                                 "maStokMaxPS=" & Me.iStockMaxPS & "," &
                                 "maListaMaterial='" & Me.sListaMat.Trim & "'," &
                                 "maActivo='" & Me.bActivo & "'," &
                                 "maLoteFijo=" & Me.iLoteFijo & "," &
                                 "maLoteMin=" & Me.iLoteMinimo & "," &
                                 "maLoteMax=" & Me.iLoteMaximo & "," &
                                 "maRedondeo=" & Me.iRedondeoLote & "," &
                                 "maDiasFabPropia=" & Me.iDiasFabPropia & "," &
                                 "maTipoTamLote='" & Me.sTipoTamañoLote.Trim & "'," &
                                 "maGrupoHR='" & Me.sGrupoHR.Trim & "'," &
                                 "maContHR='" & Me.sContadorHR.Trim & "'," &
                                 "maGrupoCompra='" & Me.sGrupoCompra & "'," &
                                 "mnMostrarInformes='" & Me.bMostrarInformes & "'," &
                                 "maUnidadesPACK=" & Me.iUnidadesCaja & "," &
                                 "maUnidadesPalet=" & Me.iUnidadesPalet & "," &
                                 "maMesesLoteCarga=" & Me.iMesesLoteCarga &
                                 " WHERE maCod='" & Me.sCodigo & "'"

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " Materiales ", Me.sCodigo)
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function


    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = " DELETE FROM materiales " &
                                  " WHERE maCod='" & Me.sCodigo & "'"

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)

            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " Materiales ", Me.sCodigo)
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function
#End Region

#Region "Propiedades"
    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property


    'Public Property Stock As Integer

    Public ReadOnly Property Stock As Integer
        Get
            If iStockActual = -999999999 Then
                iStockActual = DatosSAPConexion.DatosSAP.DameStockMARD(CENTRO:="12",
                                                                       ALMACEN:="",
                                                                       MATERIAL:=Me.sCodigo,
                                                                       GTIN:=0,
                                                                       Tipo_material:="")

                If "12" <> "12" Then
                    Dim i = ""
                End If
            End If

            Return Me.iStockActual
        End Get
    End Property

    Public Property Codigo As String
        Get
            Codigo = Me.sCodigo
        End Get
        Set(value As String)
            Me.sCodigo = value
        End Set
    End Property

    Public Property Activo As Boolean
        Get
            Return Me.bActivo
        End Get
        Set(value As Boolean)
            Me.bActivo = value
        End Set
    End Property

    Public Property MostrarEnPedidosVenta As Boolean
        Get
            Return Me.bMostrarInformes
        End Get
        Set(value As Boolean)
            Me.bMostrarInformes = value
        End Set
    End Property

    Public Property Unidades_Caja As Integer
        Get
            Return Me.iUnidadesCaja
        End Get
        Set(value As Integer)
            Me.iUnidadesCaja = value
        End Set
    End Property

    Public Property Unidades_Palet As Integer
        Get
            Return Me.iUnidadesPalet
        End Get
        Set(value As Integer)
            Me.iUnidadesPalet = value
        End Set
    End Property

    Public ReadOnly Property Cajas_por_Palet As Integer
        Get
            If Me.iUnidadesCaja > 0 Then
                Return CInt(Math.Round(iUnidadesPalet / iUnidadesCaja))
            Else
                Return 0
            End If
        End Get
    End Property

    Public Property Tipo As String
        Get
            Tipo = Me.sTipo
        End Get
        Set(value As String)
            Me.sTipo = value
        End Set
    End Property

    Public Property CodigoGrupo As String
        Get
            CodigoGrupo = Me.sGrupo
        End Get
        Set(value As String)
            Me.sGrupo = value
            miGrupo = New GrupoArticulo
        End Set
    End Property

    Public ReadOnly Property Grupo As GrupoArticulo
        Get
            If miGrupo Is Nothing Then
                miGrupo = New GrupoArticulo(_Codigo:=sGrupo)
            ElseIf Me.miGrupo.Creado = False Then
                miGrupo = New GrupoArticulo(_Codigo:=sGrupo)
            End If

            Return miGrupo
        End Get
    End Property

    Public Property UnidadMedida As String
        Get
            UnidadMedida = Me.sUnidadMedida
        End Get
        Set(value As String)
            Me.sUnidadMedida = value
        End Set
    End Property

    Public Property Nombre As String
        Get
            Nombre = Me.sNombre
        End Get
        Set(value As String)
            Me.sNombre = value
        End Set
    End Property

    Public ReadOnly Property NombreCompleto As String
        Get
            NombreCompleto = ""
            If Me.bCreado Then
                NombreCompleto = Me.sCodigo.Trim & " - " & Me.sNombre.Trim
            End If
        End Get
    End Property

    Public ReadOnly Property CabLista_Material As CabListaMaterial
        Get
            If miCabListaMaterial Is Nothing Then
                Me.miCabListaMaterial = New CabListaMaterial(Codigo:=Me.sListaMat)
            ElseIf miCabListaMaterial.Creado = False Then
                Me.miCabListaMaterial = New CabListaMaterial(Codigo:=Me.sListaMat)
            End If

            Return Me.miCabListaMaterial
        End Get

    End Property

    Public Property FamiliaEnvasado As Byte
        Get
            Return Me.iFamiliaEnvasado
        End Get
        Set(value As Byte)
            Me.iFamiliaEnvasado = value
        End Set
    End Property

    Public Property FechaIniPS As Date
        Get
            Return Me.dFechaIniPS
        End Get
        Set(value As Date)
            Me.dFechaIniPS = value
        End Set
    End Property

    Public Property FechaFinPS As Date
        Get
            Return Me.dFechaFinPS
        End Get
        Set(value As Date)
            Me.dFechaFinPS = value
        End Set
    End Property

    Public Property DiasPP As Byte
        Get
            Return Me.iDiasPP
        End Get
        Set(value As Byte)
            Me.iDiasPP = value
        End Set
    End Property

    Public Property StockMaxPS As Integer
        Get
            Return Me.iStockMaxPS
        End Get
        Set(value As Integer)
            Me.iStockMaxPS = value
        End Set
    End Property

    Public Property StockMinPS As Integer
        Get
            Return Me.iStockMinPS
        End Get
        Set(value As Integer)
            Me.iStockMinPS = value
        End Set
    End Property

    Public Property LoteFijo As Integer
        Get
            Return Me.iLoteFijo
        End Get
        Set(value As Integer)
            Me.iLoteFijo = value
        End Set
    End Property

    Public Property LoteMinimo As Integer
        Get
            Return Me.iLoteMinimo
        End Get
        Set(value As Integer)
            Me.iLoteMinimo = value
        End Set
    End Property

    Public Property LoteMaximo As Integer
        Get
            Return Me.iLoteMaximo
        End Get
        Set(value As Integer)
            Me.iLoteMaximo = value
        End Set
    End Property

    Public ReadOnly Property Enviar_Exterior As Boolean
        Get

            'If ProductoSemielaborado.Count > 0 Then

            'End If
            If HojasRutaDefecto IsNot Nothing Then
                For Each mioper In HojasRutaDefecto.OperacHojaRutaLista
                    If mioper.ClaveControlSAP = ClaveControl_PtoTrabajo_Externo Then
                        Return True
                    End If
                Next
            Else
                For Each miHR In HojasDeRuta
                    For Each mioper In miHR.OperacHojaRutaLista
                        If mioper.ClaveControlSAP = ClaveControl_PtoTrabajo_Externo Then
                            Return True
                        End If
                    Next
                Next
            End If

            Return False
        End Get
    End Property

    Public Property Redondeo As Integer
        Get
            Return Me.iRedondeoLote
        End Get
        Set(value As Integer)
            Me.iRedondeoLote = value
        End Set
    End Property

    Public Property DiasFabPropia As Integer
        Get
            Return Me.iDiasFabPropia
        End Get
        Set(value As Integer)
            Me.iDiasFabPropia = value
        End Set
    End Property

    Public Property TipoTamLote As String
        Get
            Return Me.sTipoTamañoLote
        End Get
        Set(value As String)
            Me.sTipoTamañoLote = value
        End Set
    End Property

    Public Property GrupoHojaRuta As String
        Get
            Return Me.sGrupoHR
        End Get
        Set(value As String)
            Me.sGrupoHR = value
        End Set
    End Property

    Public Property ContadorHojaRuta As String
        Get
            Return Me.sContadorHR
        End Get
        Set(value As String)
            Me.sContadorHR = value
        End Set
    End Property

    Public Property CodigoListMaterial As String
        Get
            Return Me.sListaMat
        End Get
        Set(value As String)
            Me.sListaMat = value
        End Set
    End Property

    Public Property GrupoCompra As String
        Get
            Return Me.sGrupoCompra
        End Get
        Set(value As String)
            Me.sGrupoCompra = value
        End Set
    End Property

    Public ReadOnly Property HojasDeRuta As List(Of HojaRuta)
        Get
            If misHojasRuta Is Nothing Then
                misHojasRuta = DameHojasRutaMaterial(CodigoMaterial:=Me.sCodigo.Trim)
            ElseIf misHojasRuta.Count = 0 Then
                misHojasRuta = DameHojasRutaMaterial(CodigoMaterial:=Me.sCodigo.Trim)
            End If

            Return misHojasRuta
        End Get
    End Property

    Public ReadOnly Property HojasRutaDefecto As HojaRuta
        Get

            If Me.sGrupoHR = "" And Me.sContadorHR = "" Then
                Return miHojaRutaDefecto
            End If

            If miHojaRutaDefecto Is Nothing Then
                miHojaRutaDefecto = New HojaRuta(GrupoHojaruta:=Me.sGrupoHR,
                                                ContadorGrupo:=Me.sContadorHR)
            ElseIf miHojaRutaDefecto.Creado = False Then
                miHojaRutaDefecto = New HojaRuta(GrupoHojaruta:=Me.sGrupoHR,
                                                ContadorGrupo:=Me.sContadorHR)
            End If

            Return miHojaRutaDefecto
        End Get
    End Property

    Public ReadOnly Property GrupoComprasDetalle As GrupoCompras
        Get
            If Not String.IsNullOrEmpty(sGrupoCompra) Then
                miGrupoCompra = New GrupoCompras(sGrupoCompra.Trim)
            End If
            Return miGrupoCompra
        End Get
    End Property

    Public ReadOnly Property ProductoFormula As List(Of Material)
        Get
            If Not String.IsNullOrEmpty(sCodigo) Then
                misProductoFormula = DameProducto(CodigoMaterial:=Me.sCodigo.Trim, tipoMaterial:=TipoMaterial.Fabricaciones)
            End If
            Return misProductoFormula
        End Get
    End Property
    Public ReadOnly Property ProductoSemielaborado As List(Of Material)
        Get
            If Not String.IsNullOrEmpty(sCodigo) Then
                misSemielaborados = DameProducto(CodigoMaterial:=Me.sCodigo.Trim, tipoMaterial:=TipoMaterial.Semielaborado)
            End If
            Return misSemielaborados
        End Get
    End Property
    Public ReadOnly Property ProductoPackaging As List(Of Material)
        Get
            If Not String.IsNullOrEmpty(sCodigo) Then
                misPackaging = DameProducto(CodigoMaterial:=Me.sCodigo.Trim, tipoMaterial:=TipoMaterial.Packaging)
            End If
            Return misPackaging
        End Get
    End Property
    Public Property MesesLoteCargaMaxima As Integer
        Get
            Return iMesesLoteCarga
        End Get
        Set(value As Integer)
            iMesesLoteCarga = value
        End Set
    End Property

#End Region
End Class
