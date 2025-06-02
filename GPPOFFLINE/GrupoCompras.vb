
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP
Public Class GrupoCompras

#Region "Atributos"

    Private bCreado As Boolean

#End Region

#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            Codigo = ""
            Nombre = ""
            bCreado = False
        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub
    Public Sub New(ByVal sCodigo As String, ByVal sNombre As String)
        Try
            InicializarVariables()
            Codigo = sCodigo
            Nombre = sNombre
            bCreado = True
        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal sCodigo As String, ByVal sNombre As String, ByVal sViewPedido As Boolean)
        Try
            InicializarVariables()
            Codigo = sCodigo
            Nombre = sNombre
            ViewPedido = sViewPedido
            bCreado = True
        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(sCodigo As String)
        Try

            Dim sSQl As String = "SELECT * " &
                              " FROM GrupoCompras with(nolock) " &
                              " WHERE gcCod='" & UTrim(sCodigo) & "'"

            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                Codigo = sCodigo
                Nombre = CStr((NoNull(DTDatos.Rows(0).Item("gcNombre"), "A")))
                Me.bCreado = True
            End If


        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)

        End Try
    End Sub



#End Region

#Region "Propiedades"

    Public Property Codigo As String
    Public Property Nombre As String
    Public ReadOnly Property Creado As Boolean
        Get
            Creado = bCreado
        End Get
    End Property

    Public Property ViewPedido As Boolean


#End Region

#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            Dim CodigoSQL As Integer
            Dim sSql As String = "INSERT INTO GrupoCompras (gcCod,gcNombre) " &
                                 " VALUES (" & Codigo & "," &
                                                Nombre & ") SELECT @@IDENTITY "

            CodigoSQL = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If CodigoSQL = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & "GrupoCompras", CStr(Codigo))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE GrupoCompras SET gcNombre=" & Nombre & "" &
                                 " WHERE gcCod=" & Codigo & ""

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & "GrupoCompras", Nombre)
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function ModificarVistaPedidos() As Boolean
        Try
            'Dim sSql As String = "UPDATE GrupoCompras SET gcNombre=" & Nombre & "" &
            '                     " WHERE gcCod=" & Codigo & ""

            Dim Ver = IIf(ViewPedido, "1", "0")
            Dim sSql As String = "UPDATE materiales SET mnMostrarInformes = " & Ver.ToString() & "" &
                                 " WHERE magrupoCompra='" & Codigo & "'"



            ModificarVistaPedidos = Datos.CGPL.EjecutarConsulta(sSql)

            Dim sSqlG As String = "UPDATE GrupoCompras SET gcVerPedidos=" & Ver.ToString() & "" &
                                 " WHERE gcCod='" & Codigo & "'"


            ModificarVistaPedidos = Datos.CGPL.EjecutarConsulta(sSqlG)

            If ModificarVistaPedidos Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & "GrupoCompras", Codigo)
            End If
        Catch ex As Exception
            ModificarVistaPedidos = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM GrupoCompras " &
                                  " WHERE gcCod=" & Codigo & ""

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " GrupoCompras", CStr(Codigo))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function
#End Region

End Class
