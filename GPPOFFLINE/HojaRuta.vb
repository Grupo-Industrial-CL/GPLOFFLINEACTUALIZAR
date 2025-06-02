
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class HojaRuta
#Region "Atributos"

    Private miOpeHojaRutaLista As List(Of OperacionesHojaRuta)
    Private misPuestosTrabajo As List(Of PuestosTrabajo)
    Private miFormato As Formatos
    'Private misHojaRutaxMaterial As List(Of HojaRutaxMaterial)
    Private bCreado As Boolean

#End Region

#Region "Constructores"



    Private Sub InicializarVariables()
        Try
            Grupo = String.Empty
            ContGrupo = String.Empty
            UnidadMedida = String.Empty
            Nombre = String.Empty
            Centro = String.Empty
            TipoHR = String.Empty
            Borrada = False
            Formato = String.Empty
            'miFormato = New Formatos()
            'miOpeHojaRutaLista = New List(Of OperacionesHojaRuta)
            'misPuestosTrabajo = New List(Of PuestosTrabajo)
            'misHojaRutaxMaterial = New List(Of HojaRutaxMaterial)
            bCreado = False

        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(sGrupo As String,
                   sContGrupo As String,
                   sUnidadMedida As String,
                   sNombre As String,
                   sCentro As String,
                   sTipoHR As String,
                   bBorrada As Boolean,
                   sFormato As String)
        Try
            InicializarVariables()

            Grupo = sGrupo
            ContGrupo = sContGrupo
            UnidadMedida = sUnidadMedida
            Nombre = sNombre
            Centro = sCentro
            TipoHR = sTipoHR
            Borrada = bBorrada
            Formato = sFormato

            bCreado = True

        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(GrupoHojaruta As String,
                   ContadorGrupo As String)
        Try
            Dim sSQl As String = "SELECT * " &
                                 " FROM HojaRuta with(nolock) " &
                                 " WHERE hrGrupo='" & GrupoHojaruta.Trim &
                                 "' AND hrContGrupo='" & ContadorGrupo & "'"

            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                Grupo = UTrim(DTDatos.Rows(0).Item("hrGrupo"))
                ContGrupo = UTrim(DTDatos.Rows(0).Item("hrContGrupo"))
                UnidadMedida = UTrim(DTDatos.Rows(0).Item("hrUnidadMedida"))
                Nombre = UTrim(DTDatos.Rows(0).Item("hrNombre"))
                Centro = UTrim(DTDatos.Rows(0).Item("hrCentro"))
                TipoHR = UTrim(DTDatos.Rows(0).Item("hrTipoHR"))
                Borrada = CBool(DTDatos.Rows(0).Item("hrBorrada"))
                Formato = UTrim(DTDatos.Rows(0).Item("hrFormato"))

                Me.bCreado = True
            End If
        Catch ex As Exception
            Me.bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

#End Region

#Region "Propiedades"

    ' ahorramos codigo utiliando propiedades auto implementadas. codigo mas limpio !

    Public Property Grupo As String
    Public Property ContGrupo As String
    Public Property UnidadMedida As String
    Public Property Nombre As String
    Public Property Centro As String
    Public Property TipoHR As String
    Public Property Borrada As Boolean
    Public Property Formato As String


    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property

    Public ReadOnly Property Nombre2 As String
        Get
            Return Grupo.Trim & "-" & ContGrupo.Trim & "  " & Nombre.Trim
        End Get
    End Property
    Public ReadOnly Property Codigo2 As String
        Get
            Return Grupo.Trim & "-" & ContGrupo.Trim
        End Get
    End Property



    Public ReadOnly Property OperacHojaRutaLista As List(Of OperacionesHojaRuta)
        Get
            If miOpeHojaRutaLista Is Nothing Then
                miOpeHojaRutaLista = DamelistaOperacionesHojaRuta(Grupo, ContGrupo, 0)
            ElseIf miOpeHojaRutaLista.Count = 0 Then
                miOpeHojaRutaLista = DamelistaOperacionesHojaRuta(Grupo, ContGrupo, 0)
            End If

            Return miOpeHojaRutaLista
        End Get
    End Property

    Public ReadOnly Property FormatoDetalle As Formatos
        Get
            If miFormato.Creado = False And Formato.Trim() <> "" Then
                miFormato = New Formatos(sCodigo:=Formato)
            End If

            Return miFormato
        End Get
    End Property

    Public ReadOnly Property PuestosTrabajoHojaRuta As List(Of PuestosTrabajo)
        Get
            If misPuestosTrabajo Is Nothing Then
                misPuestosTrabajo = DamePuestosTrabajoHojaRuta(GrupoHojaRuta:=Grupo,
                                                               ContHojaRuta:=ContGrupo,
                                                               TipoPuestoTrabajo:=String.Empty)
            ElseIf misPuestosTrabajo.Count = 0 Then
                misPuestosTrabajo = DamePuestosTrabajoHojaRuta(GrupoHojaRuta:=Grupo,
                                                               ContHojaRuta:=ContGrupo,
                                                               TipoPuestoTrabajo:=String.Empty)
            End If

            Return misPuestosTrabajo
        End Get
    End Property

#End Region

#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            Dim CodigoSQL As Integer
            Dim sSql As String = "INSERT INTO HojaRuta (hrGrupo,hrContGrupo,hrUnidadMedida,hrNombre,hrCentro,hrTipoHR,hrBorrada,hrFormato) " &
                                 " VALUES ('" & UTrim(Grupo) & "'," &
                                                ContGrupo & ",'" &
                                                UnidadMedida.Trim.ToUpper & ",'" &
                                                Nombre.Trim.ToUpper & ",'" &
                                                Centro.Trim.ToUpper & "','" &
                                                TipoHR.Trim & "','" &
                                                Borrada & "'" & ",'" &
                                                Formato & "'" & ")"

            Insertar = Datos.CGPL.EjecutarConsulta(sSql)

            If Insertar = True Then
                Datos.GuardarLog(TipoLogDescripcion.Alta & " HojaRuta", CStr(Grupo))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE HojaRuta " &
                                 " SET hrContGrupo = '" & UTrim(ContGrupo) & "', " &
                                 " hrUnidadMedida = '" & UTrim(UnidadMedida) & "', " &
                                 " hrNombre = '" & UTrim(Nombre) & "', " &
                                 " hrCentro = '" & UTrim(Centro) & "', " &
                                 " hrTipoHR = '" & UTrim(TipoHR) & "', " &
                                 " hrBorrada = '" & UTrim(Borrada) & "', " &
                                 " hrFormato = '" & UTrim(Borrada) & "' " &
                                 " WHERE hrGrupo='" & UTrim(Grupo) & "'"

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " HojaRuta", CStr(Grupo))
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM HojaRuta " &
                                 "WHERE hrContGrupo='" & Grupo & "'"

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " HojaRuta", CStr(Grupo))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

#End Region

End Class
