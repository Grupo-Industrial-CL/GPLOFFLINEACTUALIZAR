
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class Empresa
    Private iCodigo As Integer
    Private sNombre As String
    Private bOcultar As Boolean
    Private iSociedad As Integer
    Private bImprimirLogo As Boolean 'Si es true significa que es la empresa principal, es decir en A
    Private miSociedad As New Sociedad
    Private bCreado As Boolean

    Public Sub New()
        Try
            InicializarVariables()
        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Private Sub InicializarVariables()
        Try
            Me.iCodigo = 0
            Me.sNombre = ""
            Me.bOcultar = False
            Me.iSociedad = 0
            Me.miSociedad = New Sociedad
            Me.bImprimirLogo = False
            Me.bCreado = False

        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New(ByVal Codigo As Integer)

        Try
            Dim dtDatos As New DataTable

            InicializarVariables()

            Dim sSql As String = "SELECT * " &
                                 " FROM Empresas " &
                                 " WHERE emCodigo = " & Codigo

            If Datos.CGPL.DameDatosDT(sSql, dtDatos) Then
                Me.iCodigo = Codigo
                With dtDatos.Rows(0)
                    Me.sNombre = UTrim(.Item("emNombre"))
                    Me.bOcultar = CBool(NoNull(.Item("emOcultar"), "D"))
                    Me.iSociedad = CInt(NoNull(.Item("emSociedad"), "D"))
                    Me.bImprimirLogo = CBool(NoNull(.Item("emLogo"), "D"))
                End With
                Me.bCreado = True
            End If

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try

    End Sub

    Public Sub New(ByVal Codigo As Integer,
                   ByVal Nombre As String,
                   ByVal Sociedad As Integer,
                   ByVal Ocultar As Boolean,
                   ByVal ImprimirLogo As Boolean)

        Try
            InicializarVariables()

            Me.iCodigo = Codigo
            Me.sNombre = Nombre
            Me.iSociedad = Sociedad
            Me.bOcultar = Ocultar
            Me.bImprimirLogo = ImprimirLogo
            bCreado = True

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try

    End Sub

    Public Function Insertar() As Boolean
        Try
            Dim sSql As String = "INSERT INTO empresas (emNombre,emOcultar,emSociedad,emLogo)" &
                                 " VALUES ('" & UTrim(Me.sNombre) & " ','" &
                                                Me.bOcultar & "'," &
                                                Me.iSociedad & ",'" &
                                                Me.bImprimirLogo & "') SELECT @@IDENTITY "

            Me.iCodigo = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If Me.iCodigo = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
            End If

            If Insertar Then
                Datos.GuardarLog(TipoLogDescripcion.Alta & " Empresa ", Me.iCodigo.ToString)
            End If
        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = " UPDATE empresas " &
                                 " SET emNombre = '" & UTrim(Me.sNombre) & "'," &
                                 " emOcultar = '" & Me.bOcultar & "'," &
                                 " emSociedad = " & Me.iSociedad & "," &
                                 " emLogo = '" & Me.bImprimirLogo & "' " &
                                 " WHERE emCodigo=" & Me.iCodigo

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " Empresa", Me.iCodigo.ToString)
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public ReadOnly Property Creado() As Boolean
        Get
            Creado = bCreado
        End Get
    End Property

    Public Property Codigo() As Integer
        Get
            Codigo = iCodigo
        End Get
        Set(ByVal lValor As Integer)
            iCodigo = lValor
        End Set
    End Property

    Public Property Nombre() As String
        Get
            Nombre = sNombre
        End Get
        Set(ByVal sValor As String)
            sNombre = sValor
        End Set
    End Property

    Public ReadOnly Property Sociedad() As Sociedad
        Get
            If Me.miSociedad.Creado = False Then
                Me.miSociedad = New Sociedad(Me.iSociedad)
            End If
            Sociedad = Me.miSociedad
        End Get
    End Property

    Public Property CodigoSociedad() As Integer
        Get
            CodigoSociedad = iSociedad
        End Get
        Set(ByVal value As Integer)
            Me.iSociedad = value
            Me.miSociedad = New Sociedad
        End Set
    End Property

    Public Property Ocultar() As Boolean
        Get
            Ocultar = Me.bOcultar
        End Get
        Set(ByVal value As Boolean)
            Me.bOcultar = value
        End Set
    End Property

    'Indica si la empresa es A o B
    Public Property ImprimirLogo As Boolean
        Get
            ImprimirLogo = Me.bImprimirLogo
        End Get
        Set(value As Boolean)
            Me.bImprimirLogo = value
        End Set
    End Property

    Public ReadOnly Property CodigoYNombre As String
        Get
            CodigoYNombre = "(" & Me.iCodigo & ") " & Me.sNombre
        End Get
    End Property
End Class
