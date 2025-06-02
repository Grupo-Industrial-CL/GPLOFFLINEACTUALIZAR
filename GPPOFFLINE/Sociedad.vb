

Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class Sociedad

    Private iCodigo As Integer
    Private sNombre As String
    Private sCif As String
    Private iSociedadSAP As Integer
    Private bIncluirProduccionPtePS As Boolean
    Private bDescontarEntregasPS As Boolean
    Private iDiasLaborales As Byte
    Private sCentroSAP As String
    Private misEmpresas As New List(Of Empresa)
    Private sRamo As String
    Private miNumPaletsNotificar As Integer

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
            Me.sCif = ""
            Me.sRamo = ""

            Me.iDiasLaborales = 0
            miNumPaletsNotificar = 0
            Me.iSociedadSAP = 0
            Me.sCentroSAP = ""
            Me.bIncluirProduccionPtePS = False
            Me.bDescontarEntregasPS = False
            Me.misEmpresas = New List(Of Empresa)

            Me.bCreado = False

        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Private Sub Cargar_Datos(ByVal Codigo As Integer)

        Try
            Dim dtDatos As New DataTable

            InicializarVariables()

            Dim sSql As String = "SELECT * " &
                                 " FROM Sociedad " &
                                 " WHERE scCodigo = " & Codigo

            If Datos.CGPL.DameDatosDT(sSql, dtDatos) Then
                Me.iCodigo = Codigo
                With dtDatos.Rows(0)
                    Me.sNombre = UTrim(.Item("scNombre"))
                    Me.sCif = UTrim(.Item("scCIF"))
                    Me.iSociedadSAP = CInt(NoNull(.Item("scSociedadSAP"), "D"))
                    Me.sCentroSAP = CStr(NoNull(.Item("scCentroSAP"), "A")).Trim
                    Me.bIncluirProduccionPtePS = CBool(NoNull(.Item("scSumarProduccionPtePS"), "D"))
                    Me.bDescontarEntregasPS = CBool(NoNull(.Item("scDescontarEntregasPS"), "D"))
                    Me.iDiasLaborales = CByte(NoNull(.Item("scDiasLaborales"), "D"))
                    Me.sRamo = CStr(NoNull(.Item("scRamoMaterialesSAP"), "A"))
                    Me.miNumPaletsNotificar = CInt(NoNull(.Item("scNumPaletsNotif"), "D"))
                End With

                Me.bCreado = True
            End If

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try

    End Sub

    Public Sub New(ByVal Codigo As Integer)

        Try
            Cargar_Datos(Codigo)

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal Codigo As Integer,
                   ByVal Nombre As String,
                   ByVal Cif As String,
                   ByVal SociedadSAP As Integer,
                   ByVal Incluir_ProduccionPtePS As Boolean,
                   ByVal Descontar_EntregasPS As Boolean,
                   ByVal DiasLaborales As Byte,
                   ByVal Centro_SAP As String,
                   ByVal Ramo_SAP As String,
                   ByVal NumPaletsNotificar As Integer)

        Try
            InicializarVariables()

            Me.iCodigo = Codigo
            Me.sNombre = Nombre
            Me.sCif = Cif

            Me.iSociedadSAP = SociedadSAP
            Me.sCentroSAP = Centro_SAP
            Me.bIncluirProduccionPtePS = Incluir_ProduccionPtePS
            Me.bDescontarEntregasPS = Descontar_EntregasPS
            Me.iDiasLaborales = DiasLaborales
            Me.sRamo = Ramo_SAP
            Me.miNumPaletsNotificar = NumPaletsNotificar
            bCreado = True
        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try

    End Sub

    Public Function Insertar() As Boolean
        Try
            Dim sSql As String = "INSERT INTO sociedad (scCodigo,scNombre,scCIF,scDescontarEntregasPS,scSumarProduccionPtePS,scDiasLaborales,scCentroSAP,scRamoMaterialesSAP,scNumPaletsNotif) " &
                                 " VALUES ('" & UTrim(Me.sNombre) & "','" &
                                                UTrim(Me.sCif) & "'," &
                                                Me.iSociedadSAP & ",'" &
                                                Me.bDescontarEntregasPS & "','" &
                                                Me.bIncluirProduccionPtePS & "'," &
                                                Me.iDiasLaborales & ",'" &
                                                Me.sCentroSAP.Trim & "','" &
                                                Me.sRamo.Trim & "'," &
                                                Me.miNumPaletsNotificar & ") SELECT @@IDENTITY "

            Me.iCodigo = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If Me.iCodigo = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
            End If

            If Insertar Then
                Datos.GuardarLog(TipoLogDescripcion.Alta & " Sociedad", Me.iCodigo.ToString)
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = " UPDATE sociedad " &
                                 " SET scNombre='" & UTrim(Me.sNombre) &
                                 "',scCIF='" & UTrim(Me.sCif) & "'," &
                                 " scSociedadSAP = " & Me.iSociedadSAP & "," &
                                 " scSumarProduccionPtePS ='" & Me.bIncluirProduccionPtePS & "'," &
                                 " scDescontarEntregasPS='" & Me.bDescontarEntregasPS & "'," &
                                 " scDiasLaborales=" & Me.iDiasLaborales & "," &
                                 " scCentroSAP ='" & Me.sCentroSAP.Trim & "'," &
                                 " scRamoMaterialesSAP='" & Me.sRamo.Trim & "'," &
                                 " scNumPaletsNotif=" & Me.miNumPaletsNotificar & "" &
                                 " WHERE scCodigo = " & Me.iCodigo

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " Sociedad", Me.iCodigo.ToString)
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
        Set(ByVal Value As Integer)
            iCodigo = Value
        End Set
    End Property

    Public Property Nombre() As String
        Get
            Nombre = sNombre
        End Get
        Set(ByVal Value As String)
            sNombre = Value
        End Set
    End Property

    Public Property RamoMaterialesSAP As String
        Get
            Return Me.sRamo
        End Get
        Set(value As String)
            Me.sRamo = value
        End Set
    End Property

    Public Property Centro_SAP As String
        Get
            Return Me.sCentroSAP
        End Get
        Set(value As String)
            Me.sCentroSAP = ""
        End Set
    End Property

    Public Property CIF() As String
        Get
            CIF = sCif
        End Get

        Set(ByVal Value As String)
            sCif = Value
        End Set
    End Property

    Public Property DiasLaborales As Byte
        Get
            Return Me.iDiasLaborales
        End Get
        Set(value As Byte)
            Me.iDiasLaborales = value
        End Set
    End Property

    'Public Property Empresas() As List(Of Empresa)
    '    Get
    '        If Me.misEmpresas.Count = 0 Then
    '            Me.misEmpresas = Datos.DameEmpresas(Me.iCodigo, True, False)
    '        End If
    '        Empresas = Me.misEmpresas
    '    End Get
    '    Set(ByVal value As List(Of Empresa))
    '        Me.misEmpresas = value
    '    End Set
    'End Property

    ''' <summary>
    ''' Nos dice que empresa esta activa en la sociedad. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public ReadOnly Property EmpresaActiva As Empresa
    '    Get
    '        EmpresaActiva = (From elemento In Empresas
    '                         Where elemento.ImprimirLogo = True).ToList.FirstOrDefault
    '        If EmpresaActiva Is Nothing Then
    '            EmpresaActiva = New Empresa
    '        End If
    '    End Get
    'End Property

    Public Property CodigoSociedadSAP As Integer
        Get
            Return Me.iSociedadSAP
        End Get
        Set(value As Integer)
            Me.iSociedadSAP = value
        End Set
    End Property

    Public Property IncluirProduccionPTEPullSystem As Boolean
        Get
            IncluirProduccionPTEPullSystem = Me.bIncluirProduccionPtePS
        End Get
        Set(value As Boolean)
            Me.bIncluirProduccionPtePS = value
        End Set
    End Property

    Public Property DescontarEntregasPullSystem As Boolean
        Get
            DescontarEntregasPullSystem = Me.bDescontarEntregasPS
        End Get
        Set(value As Boolean)
            Me.bDescontarEntregasPS = value
        End Set
    End Property

    Public Property NumPaletsNotificar As Integer
        Get
            Return miNumPaletsNotificar
        End Get
        Set(value As Integer)
            miNumPaletsNotificar = value
        End Set
    End Property
    Public Sub Recargar_Sociedad()
        Try
            Cargar_Datos(Me.iCodigo)
        Catch ex As Exception

        End Try
    End Sub

End Class
