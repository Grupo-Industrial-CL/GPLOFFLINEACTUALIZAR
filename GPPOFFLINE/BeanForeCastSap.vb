
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP


Public Class BeanForeCastSap

#Region "Atributos"

    Private sCodMaterial As String
    Private sNombreMaterial As String
    Private sClaseNecesidad As String
    Private iMes As Integer
    Private iAnio As Integer
    Private iCantidadPlan As Integer
    Private sUnidad As String
    Private iStockActual As Double
    Private iFechaRotura As Date
    Private iFechaRoturaEntradas As Date
    Private bCreado As Boolean
    Private iStockBloqueado As Integer
    Private sEstatus As String
    Private iFechaCorta As Date
    Private iNecesidad As String

#End Region

#Region "Constructores"

    Public Sub New()
        Try
            InicializarVariables()
        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Private Sub InicializarVariables()
        Try
            sCodMaterial = ""
            sNombreMaterial = ""
            sClaseNecesidad = ""
            iMes = 0
            iAnio = 0
            iCantidadPlan = 0
            sUnidad = ""
            Me.iStockActual = -999999999
            bCreado = False
        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Public Sub New(ByVal CodMaterial As String,
                   ByVal NombreMaterial As String,
                   ByVal ClaseNecesidad As String,
                   ByVal Mes As Integer,
                   ByVal Anio As Integer,
                   ByVal CantidadPlan As Integer,
                   ByVal Unidad As String,
                   ByVal StockActual As Double,
                   ByVal FechaRotura As String,
                   ByVal FechaRoturaEntradas As String,
                   ByVal Fecha As String,
                   ByVal miStockBloquedo As Integer,
                   ByVal miEstatus As String,
                   ByVal FechaCorta As String,
                   ByVal Necesidad As String)
        Try
            InicializarVariables()

            sCodMaterial = CodMaterial
            sNombreMaterial = NombreMaterial
            sClaseNecesidad = ClaseNecesidad
            iMes = Mes
            iAnio = Anio
            iCantidadPlan = CantidadPlan
            sUnidad = Unidad
            iStockActual = StockActual
            Dim FechaRoturaLocal As Date = ConstantesGPP.FechaGlobal
            If FechaRotura.Contains("9999") Or FechaRotura.Contains("0000") Then

            Else
                FechaRoturaLocal = CDate(FechaRotura)
            End If
            iFechaRotura = FechaRoturaLocal
            Dim FechaRoturaEntradasLocal As Date = ConstantesGPP.FechaGlobal
            If FechaRoturaEntradas.Contains("9999") Or FechaRoturaEntradas.Contains("0000") Then

            Else
                FechaRoturaEntradasLocal = CDate(FechaRoturaEntradas)
            End If
            iFechaRoturaEntradas = FechaRoturaEntradasLocal
            iStockBloqueado = miStockBloquedo
            sEstatus = miEstatus
            Dim FechaCortaLocal As Date = ConstantesGPP.FechaGlobal
            If FechaCorta.Contains("9999") Or FechaCorta.Contains("0000") Then

            Else
                FechaCortaLocal = CDate(FechaCorta)
            End If
            iFechaCorta = FechaCortaLocal
            Dim NecesidadLocal = ""
            If Necesidad.Trim() = "VSF" Then
                NecesidadLocal = "Previsión"
            End If
            If Necesidad.Trim() = "05" Then
                NecesidadLocal = "Pedido"
            End If
            iNecesidad = NecesidadLocal
            bCreado = True
        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try

    End Sub

#End Region

#Region "Propiedades"

    Public ReadOnly Property CodMaterial As String
        Get
            Return sCodMaterial
        End Get
    End Property
    Public ReadOnly Property NombreMaterial As String
        Get
            Return sNombreMaterial
        End Get
    End Property
    Public ReadOnly Property ClaseNecesidad As String
        Get
            Return sClaseNecesidad
        End Get
    End Property

    Public ReadOnly Property Mes As Integer
        Get
            Return iMes
        End Get
    End Property
    Public ReadOnly Property Anio As Integer
        Get
            Return iAnio
        End Get
    End Property
    Public Property CantidadPlanificada As Integer
        Get
            Return iCantidadPlan
        End Get
        Set(value As Integer)
            iCantidadPlan = value
        End Set
    End Property
    Public ReadOnly Property Unidad As String
        Get
            Return If(sUnidad.Trim() = "ST", LiteralesSAP.Unidad, sUnidad)
        End Get
    End Property
    Public ReadOnly Property UnidadSAP As String
        Get
            Return sUnidad.Trim()
        End Get
    End Property
    Public ReadOnly Property Stock As Double
        Get
            Return Me.iStockActual
        End Get
    End Property

    Public ReadOnly Property FechaCorta As String
        Get
            Return iAnio.ToString & "-" & iMes.ToString.PadLeft(2, CChar("0"))
        End Get
    End Property

    Public ReadOnly Property FechaRotura As Date
        Get
            Return Me.iFechaRotura
        End Get
    End Property

    Public ReadOnly Property FechaRoturaEntradas As Date
        Get
            Return Me.iFechaRoturaEntradas
        End Get
    End Property
    Public ReadOnly Property Creado As Boolean
        Get
            Return bCreado
        End Get
    End Property

    Public ReadOnly Property StockBloqueado As Integer
        Get
            Return iStockBloqueado
        End Get
    End Property

    Public ReadOnly Property Estatus As String
        Get
            Return sEstatus.Trim()
        End Get
    End Property

    Public ReadOnly Property Fecha As Date
        Get
            Return Me.iFechaCorta
        End Get
    End Property

    Public ReadOnly Property Necesidad As String
        Get
            Return iNecesidad.Trim()
        End Get
    End Property

#End Region

End Class
