
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class PuestosTrabajo
#Region "Atributos"

    Private bCreado As Boolean

    Private miOperHojaRutaLista As New List(Of OperacionesHojaRuta)
    Private miCentroProd As CentroProd
    'Private miProveedor As Proveedor
#End Region

#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            CodigoPuestoTrabajo = 0
            Nombre = String.Empty
            Centro = String.Empty
            Operarios = 0
            Tipo = String.Empty
            AreaProduccion = String.Empty
            Recurso = False
            CambioPedido = False
            VelocidadMax = 0
            VelocidadActual = 0
            Orden = 0
            miOperHojaRutaLista = New List(Of OperacionesHojaRuta)
            Recurso = False
            CodCentroProd = 0
            CodProveedor = String.Empty
            miCentroProd = Nothing
            'miProveedor = Nothing
            EsCentroExterno = False
            PuedeRealizarTraspasos = False
            bCreado = False

        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(iCodigoPuestoTrabajo As Integer)
        Try

            Dim sSQl As String = "SELECT * " &
                              " FROM PuestosTrabajo with(nolock) " &
                              " WHERE ptCod=" & UTrim(iCodigoPuestoTrabajo)

            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                CodigoPuestoTrabajo = iCodigoPuestoTrabajo
                Nombre = UTrim(DTDatos.Rows(0).Item("ptNombre"))
                Centro = UTrim(DTDatos.Rows(0).Item("ptCentro"))
                Operarios = CInt((NoNull(DTDatos.Rows(0).Item("ptOperarios"), "N")))
                Tipo = UTrim(DTDatos.Rows(0).Item("ptTipo"))
                AreaProduccion = UTrim(DTDatos.Rows(0).Item("ptAreaProd"))
                Activo = CBool(NoNull(DTDatos.Rows(0).Item("ptActivo"), "N"))
                Recurso = CBool(NoNull(DTDatos.Rows(0).Item("ptRecurso"), "N"))
                CambioPedido = CBool(NoNull(DTDatos.Rows(0).Item("ptCambiarPedido"), "N"))
                Orden = CInt((NoNull(DTDatos.Rows(0).Item("ptOrden"), "N")))
                VelocidadMax = CInt((NoNull(DTDatos.Rows(0).Item("ptVelMax"), "N")))
                VelocidadActual = CInt((NoNull(DTDatos.Rows(0).Item("ptVelActual"), "N")))
                CodCentroProd = CInt((NoNull(DTDatos.Rows(0).Item("ptCentroProd"), "D")))
                CodProveedor = UTrim(DTDatos.Rows(0).Item("ptProveedor"))
                EsCentroExterno = CBool(NoNull(DTDatos.Rows(0).Item("ptEsCoperativaExt"), "N"))
                Me.bCreado = True
            End If


        Catch ex As Exception
            bCreado = False
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)

        End Try
    End Sub

    Public Sub New(sCodigoProveedor As String)
        Try

            Dim sSQl As String = "SELECT * " &
                              " FROM PuestosTrabajo with(nolock) " &
                              " WHERE ptProveedor='" & UTrim(sCodigoProveedor) + "'"

            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                CodigoPuestoTrabajo = CInt((NoNull(DTDatos.Rows(0).Item("ptOperarios"), "N")))
                Nombre = UTrim(DTDatos.Rows(0).Item("ptNombre"))
                Centro = UTrim(DTDatos.Rows(0).Item("ptCentro"))
                Operarios = CInt((NoNull(DTDatos.Rows(0).Item("ptOperarios"), "N")))
                Tipo = UTrim(DTDatos.Rows(0).Item("ptTipo"))
                AreaProduccion = UTrim(DTDatos.Rows(0).Item("ptAreaProd"))
                Activo = CBool(NoNull(DTDatos.Rows(0).Item("ptActivo"), "N"))
                Recurso = CBool(NoNull(DTDatos.Rows(0).Item("ptRecurso"), "N"))
                CambioPedido = CBool(NoNull(DTDatos.Rows(0).Item("ptCambiarPedido"), "N"))
                Orden = CInt((NoNull(DTDatos.Rows(0).Item("ptOrden"), "N")))
                VelocidadMax = CInt((NoNull(DTDatos.Rows(0).Item("ptVelMax"), "N")))
                VelocidadActual = CInt((NoNull(DTDatos.Rows(0).Item("ptVelActual"), "N")))
                CodCentroProd = CInt((NoNull(DTDatos.Rows(0).Item("ptCentroProd"), "D")))
                CodProveedor = UTrim(DTDatos.Rows(0).Item("ptProveedor"))
                EsCentroExterno = CBool(NoNull(DTDatos.Rows(0).Item("ptEsCoperativaExt"), "N"))
                PuedeRealizarTraspasos = CBool(NoNull(DTDatos.Rows(0).Item("ptPuedeRealizarTraspaso"), "N"))
                Me.bCreado = True
            End If


        Catch ex As Exception
            bCreado = False
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)

        End Try
    End Sub

    Public Sub New(_CodigoPuestoTrabajo As Integer,
                   _Nombre As String,
                   _Centro As String,
                   _Operarios As Integer,
                   _Tipo As String,
                   _AreaProduccion As String,
                   _Activo As Boolean,
                   _Recurso As Boolean,
                   _CambioPedido As Boolean,
                   _Orden As Integer,
                   _VelocidadMax As Integer,
                   _VelocidadActual As Integer,
                   _CodCentroProd As Integer,
                   _CodProveedor As String,
                   _EsCentroExterno As Boolean,
                   _PuedeRealizarTraspasos As Boolean)
        Try
            InicializarVariables()
            CodigoPuestoTrabajo = _CodigoPuestoTrabajo
            Nombre = _Nombre
            Centro = _Centro
            Operarios = _Operarios
            Tipo = _Tipo
            AreaProduccion = _AreaProduccion
            Activo = _Activo
            Recurso = _Recurso
            CambioPedido = _CambioPedido
            Orden = _Orden
            VelocidadMax = _VelocidadMax
            VelocidadActual = _VelocidadActual
            CodCentroProd = _CodCentroProd
            CodProveedor = _CodProveedor
            EsCentroExterno = _EsCentroExterno
            PuedeRealizarTraspasos = _PuedeRealizarTraspasos

            Me.bCreado = True

        Catch ex As Exception
            bCreado = False
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)

        End Try
    End Sub


#End Region

#Region "Propiedades"
    Public ReadOnly Property Creado As Boolean
        Get
            Creado = bCreado
        End Get
    End Property

    Public Property CodigoPuestoTrabajo As Integer
    Public Property Nombre As String
    Public Property Centro As String
    Public Property Operarios As Integer
    Public Property Tipo As String
    Public Property AreaProduccion As String
    Public Property Activo As Boolean
    Public Property Recurso As Boolean
    Public Property CambioPedido As Boolean
    Public Property Orden As Integer
    Public Property VelocidadMax As Integer
    Public Property VelocidadActual As Integer
    Public Property CodCentroProd As Integer
    Public Property CodProveedor As String
    Public Property EsCentroExterno As Boolean
    Public Property PuedeRealizarTraspasos As Boolean


    Public ReadOnly Property Centro_Prod As CentroProd
        Get
            If Me.miCentroProd Is Nothing Then
                Me.miCentroProd = New CentroProd(CodCentroProd)
            End If
            Return Me.miCentroProd
        End Get
    End Property

    Public ReadOnly Property Proxima_FechaInicio_Turno As DateTime
        Get
            Dim sSql As String = "SELECT TOP (1) clInicioTurno " &
                                 "FROM Calendario " &
                                 "WHERE clPuestoTrabajo = " & Me.CodigoPuestoTrabajo &
                                 " AND clInicioTurno >= CURRENT_TIMESTAMP"
            Dim DTDatos As New DataTable

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Return CDate(DTDatos.Rows(0).Item("clInicioTurno"))
            Else
                Return FechaGlobal
            End If
        End Get
    End Property

    'Public ReadOnly Property Proveedor As Proveedor
    '    Get
    '        If Me.miProveedor Is Nothing Then
    '            Me.miProveedor = New Proveedor(CodProveedor)
    '        End If
    '        Return Me.miProveedor

    '    End Get
    'End Property

#End Region

#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            Dim CodigoSQL As Integer
            Dim sSql As String = "INSERT INTO PuestosTrabajo (ptCod,ptNombre,ptCentro,ptOperarios,ptTipo,ptAreaProd,ptActivo,ptRecurso," &
                                 "ptCambiarPedido,ptOrden,ptVelMax,ptVelActual,ptCentroProd,ptProveedor,ptEsCoperativaExt,ptPuedeRealizarTraspaso) " &
                                 " VALUES (" & CodigoPuestoTrabajo & ",'" &
                                               Nombre & "','" &
                                               Centro & "'," &
                                               Operarios & ",'" &
                                               Tipo & "','" &
                                               AreaProduccion & "','" &
                                               Activo & "','" &
                                               Recurso & "','" &
                                               CambioPedido & "'," &
                                               Orden & "," &
                                               VelocidadMax & "," &
                                               VelocidadActual & "," &
                                               CodCentroProd & ",'" &
                                               CodProveedor & "','" &
                                               EsCentroExterno & "','" &
                                               PuedeRealizarTraspasos & "') SELECT @@IDENTITY "

            CodigoSQL = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))



            If CodigoSQL = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & " PuestosTrabajo", CStr(CodigoPuestoTrabajo))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE PuestosTrabajo " &
                                 " SET ptNombre = '" & UTrim(Nombre) & "', " &
                                 " ptCentro = '" & Centro & "', " &
                                 " ptOperarios = " & Operarios & ", " &
                                 " ptTipo = '" & Tipo & "', " &
                                 " ptAreaProd = '" & AreaProduccion & "', " &
                                 " ptActivo = '" & Activo & "', " &
                                 " ptRecurso = '" & Recurso & "', " &
                                 " ptCambiarPedido = '" & CambioPedido & "', " &
                                 " ptOrden = " & Orden & ", " &
                                 " ptVelMax = " & VelocidadMax & ", " &
                                 " ptVelActual = " & VelocidadActual & "," &
                                 " ptCentroProd =" & CodCentroProd & "," &
                                 " ptProveedor = '" & CodProveedor & "'," &
                                 " ptEsCoperativaExt = '" & EsCentroExterno & "'," &
                                 " ptPuedeRealizarTraspaso = '" & PuedeRealizarTraspasos & "' " &
                                 " WHERE ptCod=" & CodigoPuestoTrabajo

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " PuestosTrabajo", CStr(CodigoPuestoTrabajo))
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM PuestosTrabajo " &
                                 "WHERE ptCod=" & CodigoPuestoTrabajo

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " PuestosTrabajo", CStr(CodigoPuestoTrabajo))
            End If
        Catch ex As Exception
            Eliminar = False
            'Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Shared Narrowing Operator CType(v As List(Of PuestosTrabajo)) As PuestosTrabajo
        Throw New NotImplementedException()
    End Operator

#End Region
End Class
