
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class CabListaMaterial

#Region "Atributos"
    Private sCodigo As String
    Private sUnidadBase As String
    Private dCantidadBase As Double
    Private misMaterialesLista As List(Of ListaMaterial)
    Private bCreado As Boolean
#End Region

#Region "Constructores"
    Private Sub InicializarVariables()
        Try
            sCodigo = ""
            sUnidadBase = ""
            dCantidadBase = 0
            Me.misMaterialesLista = New List(Of ListaMaterial)
            Me.bCreado = False
        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(Codigo As String,
                   UnidadBase As String,
                   CantidadBase As Double)
        Try
            InicializarVariables()
            Me.sCodigo = Codigo
            Me.sUnidadBase = UnidadBase
            Me.dCantidadBase = CantidadBase

            Me.bCreado = True
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(Codigo As String)
        Try
            Dim sSQl As String = "SELECT * " &
                                 " FROM ListaMaterialesCab " &
                                 " WHERE upper(rtrim(clLista)) = '" & UTrim(Codigo) & "'"
            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                Me.sCodigo = Codigo
                Me.sUnidadBase = UTrim(DTDatos.Rows(0).Item("clUM"))
                Me.dCantidadBase = CDbl(NoNull(DTDatos.Rows(0).Item("clCantidad"), "D"))
                Me.bCreado = True
            End If
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub
#End Region

#Region "Propiedades"
    Public Property UnidadBase As String
        Get
            UnidadBase = Me.sUnidadBase
        End Get
        Set(value As String)
            Me.sUnidadBase = value
        End Set
    End Property

    Public Property CantidadBase As Double
        Get
            CantidadBase = Me.dCantidadBase
        End Get
        Set(value As Double)
            Me.dCantidadBase = value
        End Set
    End Property

    Public Property Codigo As String
        Get
            Codigo = Me.sCodigo
        End Get
        Set(value As String)
            Me.sCodigo = value
        End Set
    End Property

    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property

    Public ReadOnly Property MaterialesLista As List(Of ListaMaterial)

        Get
            If Me.misMaterialesLista.Count = 0 And Not String.IsNullOrEmpty(Me.Codigo) Then
                Me.misMaterialesLista = DameListasMaterial(Me.Codigo)
            End If
            MaterialesLista = Me.misMaterialesLista
        End Get

    End Property
#End Region

#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            Dim CodigoSQL As Integer
            Dim sSql As String = "INSERT INTO ListaMaterialesCab (clLista,clCantidad,clUM) " &
                                 " VALUES ('" & UTrim(sCodigo) & "'," &
                                                dCantidadBase & ",'" & sUnidadBase & "'" & ") SELECT @@IDENTITY "

            CodigoSQL = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If CodigoSQL = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & " ListaMaterialesCab", CStr(Codigo))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE ListaMaterialesCab " &
                                 " SET clUM = '" & UTrim(sUnidadBase) & "', " &
                                 " clCantidad = " & dCantidadBase & " " &
                                 " WHERE clLista=" & UTrim(sCodigo)

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " ListaMaterialesCab", CStr(Codigo))
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM ListaMaterialesCab " &
                                 "WHERE clLista=" & sCodigo

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " ListaMaterialesCab", CStr(Codigo))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

#End Region

End Class
