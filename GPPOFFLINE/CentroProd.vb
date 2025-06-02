
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class CentroProd

#Region "Atributos"

    Private iCodigo As Integer
    Private sNombre As String
    Private iCodSociedad As Integer
    Private miSociedad As Sociedad

    Private bCreado As Boolean

#End Region

#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            Me.iCodigo = 0
            Me.sNombre = ""
            Me.iCodSociedad = 0
            Me.miSociedad = Nothing

            Me.bCreado = False
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(Codigo As Integer,
                   Nombre As String,
                   Cod_Sociedad As Integer)
        Try
            InicializarVariables()

            Me.iCodigo = Codigo
            Me.sNombre = Nombre
            Me.iCodSociedad = Cod_Sociedad

            Me.bCreado = True
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Try
            Dim DTDAtos As New DataTable
            Dim sSql As String = "SELECT * " &
                                 "FROM CentrosProd " &
                                 "WHERE ceCod = " & Codigo

            Me.InicializarVariables()

            If Datos.CGPL.DameDatosDT(sSql, DTDAtos) Then
                With DTDAtos.Rows(0)
                    Me.iCodigo = CShort(NoNull(.Item("ceCod"), "D"))
                    Me.sNombre = CStr(NoNull(.Item("ceNombre"), "A")).Trim
                    Me.iCodSociedad = CInt(NoNull(.Item("ceSociedad"), "D"))
                End With
                Me.bCreado = True
            End If

        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Public Function Insertar() As Boolean
        Try
            Dim sSql As String = "INSERT CentrosProd (ceNombre,ceSociedad) VALUES ('" &
                                                               Me.sNombre.ToUpper.Trim & "'," &
                                                               Me.iCodSociedad & ")  SELECT @@IDENTITY "

            Me.iCodigo = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If iCodigo = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & " Perfil", CStr(iCodigo))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE CentrosProd SET " &
                                 " ceNombre='" & Me.sNombre.ToUpper.Trim & "'," &
                                 " ceSociedad=" & iCodSociedad &
                                 " WHERE ceCod=" & Me.iCodigo

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Alta & " " & Me.GetType().Name & " ", Me.iCodigo.ToString)
            End If

        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM CentrosProd " &
                                 "WHERE ceCod=" & Me.iCodigo

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " ", CStr(Me.Codigo))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function




    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property

    Public Property Codigo As Integer
        Get
            Codigo = Me.iCodigo
        End Get
        Set(value As Integer)
            Me.iCodigo = value
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

    Public Property CodigoSociedadO As Integer
        Get
            Return Me.iCodSociedad
        End Get
        Set(value As Integer)
            Me.iCodSociedad = value
        End Set
    End Property

    Public ReadOnly Property Sociedad As Sociedad
        Get
            If miSociedad Is Nothing Then
                miSociedad = New Sociedad(Codigo:=Me.iCodSociedad)
            End If
            Return miSociedad
        End Get
    End Property
#End Region
End Class
