
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class GrupoArticulo

#Region "Atributos"
    'Private sCodigo As String
    Public Property Codigo As String
    Private sNombre As String
    Private sDescripcion As String
    Private bMostrarEntregas As Boolean

    Private bCreado As Boolean
#End Region

#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            Codigo = ""
            Me.sNombre = ""
            Me.sDescripcion = ""

            Me.bMostrarEntregas = False

            Me.bCreado = False
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

    Public Sub New(ByVal _Codigo As String,
                   ByVal Nombre As String,
                   ByVal Descripcion As String,
                   ByVal Mostrar_Entregas As Boolean)
        Try
            InicializarVariables()

            Codigo = _Codigo
            Me.sNombre = Nombre
            Me.sDescripcion = Descripcion
            Me.bMostrarEntregas = Mostrar_Entregas

            Me.bCreado = True

        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(_Codigo As String)
        Try
            Dim sSql As String = " SELECT * " &
                                 " FROM GrupoArticulos " &
                                 " WHERE upper(rtrim(toCod)) = '" & UTrim(Codigo) & "'"
            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Codigo = _Codigo
                Me.sNombre = UTrim(DTDatos.Rows(0).Item("toNombre"))
                Me.sDescripcion = UTrim(DTDatos.Rows(0).Item("toDesc"))
                Me.bMostrarEntregas = CBool(NoNull(DTDatos.Rows(0).Item("toMostrarEntregas"), "D"))
                Me.bCreado = True
            End If
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

#End Region

#Region "BBDD"

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = " UPDATE GrupoArticulos " &
                                 " SET toNombre='" & Me.sNombre.Trim.ToUpper & "'," &
                                 " toTNDesperdicioFIN='" & Me.sDescripcion.Trim.ToUpper & "'," &
                                 " toMostrarEntregas='" & Me.bMostrarEntregas & "' " &
                                 " WHERE UPPER(RTRIM(toCod)) = '" & Codigo.Trim & "'"

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " GrupoArticulo ", Codigo)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name &
                                                            "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function


#End Region

#Region "Propiedades"

    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property

    'Public Property Codigo As String
    '    Get
    '        Codigo = codigo
    '    End Get
    '    Set(value As String)
    '        codigo = value
    '    End Set
    'End Property

    Public Property Nombre As String
        Get
            Nombre = Me.sNombre
        End Get
        Set(value As String)
            Me.sNombre = value
        End Set
    End Property

    Public Property Descripcion As String
        Get
            Descripcion = Me.sDescripcion
        End Get
        Set(value As String)
            Me.sDescripcion = value
        End Set
    End Property

    Public Property MostrarEntregas As Boolean
        Get
            MostrarEntregas = Me.bMostrarEntregas
        End Get
        Set(value As Boolean)
            Me.bMostrarEntregas = value
        End Set
    End Property
#End Region

End Class
