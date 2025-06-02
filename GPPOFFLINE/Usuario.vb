
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class Usuario

    Private iCodigo As Integer
    Private sId As String
    Private sPwd As String
    Private sNombre As String
    Private dFecAlta As Date
    Private dFecBaja As Date
    Private bActivo As Boolean
    Private sEmail As String
    Private iSociedadActual As Integer
    Private sPIN As String
    Private iCodPuestoTrabajo As Integer

    'Private misPerfiles As List(Of Perfil)
    'Private misSociedades As New List(Of SociedadUsuario)
    'Private miOperacionActual As New Operacion
    'Private misFavoritos As New List(Of Operacion)
    Private miPuestoTrabajo As New PuestosTrabajo

    'Private miAgendaPersonal As New List(Of TareaAgenda)
    'Private miAgendaGrupo As New List(Of TareaAgenda)
    'Private miAgendaSinAsignar As New List(Of TareaAgenda)
    'Private miAgendaPendiente As New List(Of TareaAgenda)

    Private misPalabrasIdioma As New Dictionary(Of String, String)

    Private bCreado As Boolean


    Public Sub New()
        Try
            Inicializar_Variables()
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal IdUsuario As String,
                   ByVal Codigo As Integer,
                   ByVal Pass As String,
                   ByVal Nombre As String,
                   ByVal FecAlta As Date,
                   ByVal FecBaja As Date,
                   ByVal Activo As Boolean,
                   ByVal Email As String,
                   ByVal CodigoPIN As String,
                   ByVal CodPuestotrabajo As Integer)
        Try
            Inicializar_Variables()

            Me.sId = IdUsuario
            Me.iCodigo = Codigo
            Me.sPwd = Pass
            Me.sNombre = Nombre
            Me.dFecAlta = FecAlta
            Me.dFecBaja = FecBaja
            Me.bActivo = Activo
            Me.sEmail = Email
            Me.sPIN = CodigoPIN
            Me.iCodPuestoTrabajo = CodPuestotrabajo
            bCreado = True

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(ByVal Id As String)
        Try
            Inicializar_Variables()
            'Cargar_Datos(Id)
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try

    End Sub

    Public Sub New(ByVal CodUsuario As Integer)
        Try
            Inicializar_Variables()
            'Cargar_Datos(CodUsuario)
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try

    End Sub
    Public Sub New(ByVal Id As String,
                             ByVal Pin As Boolean)
        Try
            Dim TUsuario As New DataTable
            If Pin AndAlso Id = "" Then
                Me.bCreado = False
                Exit Sub
            End If
            Dim sSql As String = "SELECT * " &
                                 " FROM Usuarios " &
                                 CStr(IIf(Not Pin, " WHERE USID='" & Trim(Id) & "'",
                                                   " WHERE usPIN='" & Trim(Id) & "'"))

            Inicializar_Variables()
            If Datos.CGPL.DameDatosDT(sSql, TUsuario) Then
                'Cargar_Variables(TUsuario)
            End If

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Private Sub Inicializar_Variables()
        Try
            Me.iCodigo = 0
            Me.sId = ""
            Me.sNombre = ""
            Me.sPwd = ""
            Me.dFecAlta = FechaGlobal
            Me.dFecBaja = FechaGlobal
            Me.sEmail = ""
            Me.bActivo = False
            Me.sPIN = ""
            Me.iCodPuestoTrabajo = 0

            Me.iSociedadActual = 0
            'Me.misSociedades = New List(Of SociedadUsuario)
            'Me.misPerfiles = New List(Of Perfil)

            'Me.misFavoritos = New List(Of Operacion)
            'Me.miOperacionActual = New Operacion
            miPuestoTrabajo = Nothing

            'Me.miAgendaGrupo = New List(Of TareaAgenda)
            'Me.miAgendaPersonal = New List(Of TareaAgenda)
            'Me.miAgendaSinAsignar = New List(Of TareaAgenda)
            'Me.miAgendaPendiente = New List(Of TareaAgenda)

            Me.misPalabrasIdioma = New Dictionary(Of String, String)

            Me.bCreado = False

        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    'Private Sub Cargar_Datos(ByVal Id As String)
    '    Try
    '        Dim TUsuario As New DataTable
    '        Dim sSql As String = "SELECT * " &
    '                             " FROM Usuarios " &
    '                             " WHERE USID='" & Trim(Id) & "'"

    '        Inicializar_Variables()
    '        If Datos.CGPL.DameDatosDT(sSql, TUsuario) Then
    '            Cargar_Variables(TUsuario)
    '        End If

    '    Catch ex As Exception
    '        bCreado = False
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Sub

    'Private Sub Cargar_Datos(ByVal Codigo As Integer)
    '    Try
    '        Dim TUsuario As New DataTable
    '        Dim sSql As String = "SELECT * " &
    '                             " FROM Usuarios " &
    '                             " WHERE USCodigo=" & Codigo

    '        Inicializar_Variables()

    '        If Datos.CGPL.DameDatosDT(sSql, TUsuario) Then
    '            Cargar_Variables(TUsuario)
    '        End If

    '    Catch ex As Exception
    '        bCreado = False
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Sub

    'Private Sub Cargar_Variables(ByVal dTDatos As DataTable)
    '    Try
    '        Dim i As Integer = 0
    '        Dim bEncontrado As Boolean = False

    '        With dTDatos.Rows(0)
    '            Me.sId = CStr(NoNull(.Item("USID"), "A"))
    '            Me.iCodigo = CInt(NoNull(.Item("USCODIGO"), "D"))
    '            Me.iCodPuestoTrabajo = CInt(NoNull(.Item("usPuestoTrabajo"), "D"))
    '            Me.sPwd = CStr(NoNull(.Item("USPWD"), "A"))
    '            Me.sNombre = CStr(NoNull(.Item("USNOMBRE"), "A"))
    '            Me.dFecAlta = CDate(NoNull(.Item("USFECALTA"), "DT"))
    '            Me.dFecBaja = CDate(NoNull(.Item("USFECBAJA"), "DT"))
    '            Me.bActivo = CBool(.Item("usActivo"))
    '            Me.sEmail = CStr(NoNull(.Item("USMAIL"), "A"))
    '            Me.sPIN = CStr(NoNull(.Item("USPIN"), "A"))

    '            'Cargamos la sociedad por defecto, en el caso de que no tenga ninguna se le pone por defecto la primera que encontremos.
    '            If Me.Sociedades.Count > 0 Then
    '                i = 0
    '                While i < Me.Sociedades.Count And Not bEncontrado
    '                    If Me.Sociedades(i).PorDefecto = True Then
    '                        bEncontrado = True
    '                        Me.iSociedadActual = Me.Sociedades(i).CodigoSociedad
    '                    End If
    '                    i = i + 1
    '                End While

    '                If Me.iSociedadActual = 0 Then
    '                    Me.iSociedadActual = Me.Sociedades(0).CodigoSociedad
    '                    Me.Sociedades(0).PorDefecto = True
    '                    Me.Sociedades(0).Modificar()
    '                End If
    '            End If
    '        End With

    '        bCreado = True

    '    Catch ex As Exception
    '        Inicializar_Variables()
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Sub

    Public Function Insertar() As Boolean
        Try
            Dim sSql As String = "INSERT INTO Usuarios (USID,USPWD,USNOMBRE,USFECALTA,USACTIVO," &
                                 "USMAIL,USFECBAJA,USPIN,usPuestoTrabajo)" &
                                 " VALUES ('" & Me.sId & " ','" &
                                              Trim(sPwd) & "','" &
                                              UTrim(sNombre) & "',current_timestamp,'" &
                                              Me.bActivo & "','" &
                                              Trim(sEmail) & "','" &
                                              FechaGlobal & "','" &
                                              sPIN.Trim & "'," &
                                              Me.iCodPuestoTrabajo & ") SELECT @@IDENTITY "

            If sPIN.Trim <> "" AndAlso Datos.CodigoPINRepetido(Me.iCodigo, Me.sPIN) = True Then
                Throw New Exception("CÓDIGO PIN REPETIDO")
                Insertar = False
                Exit Function
            End If

            Me.iCodigo = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If Me.iCodigo = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                bCreado = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & " Usuario", CStr(Me.iCodigo))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            If sPIN.Trim <> "" AndAlso Datos.CodigoPINRepetido(Me.iCodigo, sPIN) = True Then
                Throw New Exception("CÓDIGO PIN REPETIDO")
                Modificar = False
                Exit Function
            End If

            Dim sSql As String = " UPDATE usuarios " &
                                 " SET USID = '" & Trim(Me.Id) & "'," &
                                 " USPWd = '" & Me.sPwd & "'," &
                                 " USNOMBRE = '" & UTrim(Me.sNombre) & "'," &
                                 " USACTIVO = '" & bActivo & "'," &
                                 " USFECBAJA = '" & Me.dFecBaja & "'," &
                                 " USMAIL = '" & Trim(Me.sEmail) & "'," &
                                 " USPIN = '" & Trim(Me.sPIN) & "'," &
                                 " USPUESTOTRABAJO = " & Me.iCodPuestoTrabajo &
                                 " WHERE USCODIGO=" & Me.iCodigo

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)

            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " usuario", CStr(Me.iCodigo))
            End If

        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    'Public Property Sociedades As List(Of SociedadUsuario)
    '    Get
    '        If Me.misSociedades Is Nothing Then
    '            Me.misSociedades = Datos.DameSociedadesUsuario(Me.iCodigo)
    '        ElseIf Me.misSociedades.Count = 0 Then
    '            Me.misSociedades = Datos.DameSociedadesUsuario(Me.iCodigo)
    '        End If
    '        Return Me.misSociedades
    '    End Get

    '    Set(value As List(Of SociedadUsuario))
    '        Me.misSociedades = value
    '    End Set
    'End Property

    'Public ReadOnly Property SociedadActual As SociedadUsuario
    '    Get
    '        If Sociedades.Count > 0 AndAlso Me.iSociedadActual > 0 Then
    '            SociedadActual = (From miSociedad In Me.misSociedades
    '                              Where miSociedad.CodigoSociedad = Me.iSociedadActual).ToList.First
    '        Else
    '            SociedadActual = New SociedadUsuario
    '        End If
    '    End Get
    'End Property

    Public ReadOnly Property Puesto_Trabajo As PuestosTrabajo
        Get
            If Me.miPuestoTrabajo Is Nothing Then
                Me.miPuestoTrabajo = New PuestosTrabajo(Me.iCodPuestoTrabajo)
            End If
            Return Me.miPuestoTrabajo
        End Get
    End Property

    Public Property CodigoPuestoTrabajo() As Integer
        Get
            Return Me.iCodPuestoTrabajo
        End Get
        Set(ByVal value As Integer)
            Me.iCodPuestoTrabajo = value
        End Set
    End Property

    'Public Function OperacionValida(ByVal Clave As Integer,
    '                                ByRef Optional miOperacion As Operacion = Nothing) As Boolean
    '    Try
    '        Dim i As Integer = 0
    '        Dim j As Integer = 0

    '        OperacionValida = False
    '        For j = 0 To Perfiles.Count - 1
    '            'esto es un pelin más rapido pero no lo pongo para no variar el resultado:

    '            If Perfiles(j).Operaciones.Count > 0 AndAlso Perfiles(j).OperacionesHash.ContainsKey(Clave) Then
    '                miOperacion = Perfiles(j).OperacionesHash.Item(Clave)
    '                Return True
    '            End If
    '        Next

    '    Catch ex As Exception
    '        OperacionValida = False
    '    End Try
    'End Function

    Public ReadOnly Property Creado() As Boolean
        Get
            Creado = bCreado
        End Get
    End Property

    Public Property Codigo() As Integer
        Get
            Codigo = iCodigo
        End Get
        Set(ByVal value As Integer)
            iCodigo = value
        End Set
    End Property

    Public Property Id() As String
        Get
            Id = Trim(sId)
        End Get
        Set(ByVal value As String)
            sId = value
        End Set
    End Property

    'Public Property Favoritos As List(Of Operacion)
    '    Get
    '        If Me.misFavoritos Is Nothing Then
    '            Me.misFavoritos = Datos.DameFavoritos(Me.iCodigo)
    '        ElseIf Me.misFavoritos.Count = 0 Then
    '            Me.misFavoritos = Datos.DameFavoritos(Me.iCodigo)
    '        End If
    '        Return Me.misFavoritos
    '    End Get

    '    Set(value As List(Of Operacion))
    '        Me.misFavoritos = value
    '    End Set
    'End Property

    Public Property PIN As String
        Get
            PIN = Me.sPIN
        End Get
        Set(value As String)
            Me.sPIN = value
        End Set
    End Property

    'Public Function AñadirAFavoritos(CodigoOperacion As Integer) As Boolean
    '    Try
    '        Dim sSQl As String = "INSERT INTO Favoritos (faUsuario,faOperacion) " &
    '                             " VALUES(" & Me.iCodigo & "," &
    '                                          CodigoOperacion & ")"

    '        AñadirAFavoritos = Datos.CGPL.EjecutarConsulta(sSQl)

    '        If AñadirAFavoritos Then
    '            Me.misFavoritos = New List(Of Operacion)
    '        End If

    '    Catch ex As Exception
    '        AñadirAFavoritos = False
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    'Public Function EliminarDeFavoritos(CodigoOperacion As Integer) As Boolean
    '    Try
    '        Dim sSQl As String = "DELETE FROM Favoritos " &
    '                             " WHERE faUsuario = " & Me.iCodigo &
    '                             " AND faOperacion = " & CodigoOperacion

    '        EliminarDeFavoritos = Datos.CGPL.EjecutarConsulta(sSQl)

    '        If EliminarDeFavoritos Then
    '            Me.misFavoritos = New List(Of Operacion)
    '        End If

    '    Catch ex As Exception
    '        EliminarDeFavoritos = False
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    'Public Function EstaEnFavoritos(CodigoOperacion As Integer) As Boolean
    '    Try
    '        Dim i As Integer = 0
    '        EstaEnFavoritos = False

    '        While i < Favoritos.Count And Not EstaEnFavoritos
    '            If Me.misFavoritos(i).Codigo = CodigoOperacion Then
    '                EstaEnFavoritos = True
    '            End If
    '            i = i + 1
    '        End While
    '    Catch ex As Exception
    '        EstaEnFavoritos = False
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    Public Property Email() As String
        Get
            Email = sEmail
        End Get
        Set(ByVal value As String)
            sEmail = value
        End Set
    End Property

    Public Property FechaAlta() As Date
        Get
            FechaAlta = dFecAlta
        End Get
        Set(ByVal value As Date)
            dFecAlta = value
        End Set
    End Property

    Public Property FechaBaja() As Date
        Get
            FechaBaja = dFecBaja
        End Get
        Set(ByVal value As Date)
            dFecBaja = value
        End Set
    End Property

    Public Property Nombre() As String
        Get
            Nombre = sNombre
        End Get
        Set(ByVal value As String)
            sNombre = value
        End Set
    End Property

    Public Property Activo() As Boolean
        Get
            Activo = bActivo
        End Get
        Set(ByVal value As Boolean)
            bActivo = value
        End Set
    End Property

    Public Property Contraseña() As String
        Get
            Contraseña = sPwd
        End Get
        Set(ByVal value As String)
            sPwd = value
        End Set
    End Property

    ''' <summary>
    ''' Empresa a la que está conectado el usuario actualmente
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CodigoSociedadActual() As Integer
        Get
            CodigoSociedadActual = Me.iSociedadActual
        End Get
        Set(ByVal value As Integer)
            iSociedadActual = value
        End Set
    End Property

    'Public ReadOnly Property EsPlanificador() As Boolean
    '    Get
    '        For Each miPerfil In Me.Perfiles
    '            If miPerfil.EsPlanificador Then
    '                Return True
    '                Exit Property
    '            End If
    '        Next

    '        Return False
    '    End Get
    'End Property

    'Public Property Perfiles() As List(Of Perfil)
    '    Get
    '        If misPerfiles Is Nothing Then
    '            misPerfiles = Datos.DamePerfilesUsuario(Me.iCodigo)
    '        ElseIf misPerfiles.Count = 0 Then
    '            misPerfiles = Datos.DamePerfilesUsuario(Me.iCodigo)
    '        End If
    '        Return misPerfiles
    '    End Get
    '    Set(ByVal value As List(Of Perfil))
    '        misPerfiles = value
    '    End Set
    'End Property

    'Public Sub AñadirPerfil(ByVal Perf As Perfil)
    '    Try
    '        Dim miPerfilUser As New PerfilXUsuario(Perf.Codigo, Me.iCodigo)
    '        If miPerfilUser.Insertar Then
    '            Me.misPerfiles.Add(Perf)
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Sub

    'Public Sub AñadirSociedad(ByVal Sociedad As Sociedad, PorDefecto As Boolean)
    '    Try
    '        Dim miSocUser As New SociedadUsuario(Sociedad.Codigo, Me.iCodigo, PorDefecto)
    '        If miSocUser.Insertar Then
    '            Me.misSociedades.Add(miSocUser)
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Sub

    'Public Sub EliminarPerfil(ByVal Perf As Perfil)
    '    Try
    '        Dim i As Integer = 0
    '        Dim bEncontrado As Boolean = False
    '        Dim miPerfilUser As New PerfilXUsuario(Perf.Codigo, Me.iCodigo)

    '        While i < Perfiles.Count And Not bEncontrado
    '            If Perf.Codigo = Perfiles(i).Codigo Then
    '                misPerfiles.RemoveAt(i)
    '                miPerfilUser.Eliminar()
    '                bEncontrado = True
    '            End If
    '            i = i + 1
    '        End While
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Sub

    'Public Sub EliminarSociedad(ByVal Sociedad As Sociedad)
    '    Try
    '        Dim i As Integer = 0
    '        Dim bEncontrado As Boolean = False
    '        Dim miSocUser As New SociedadUsuario(Sociedad.Codigo, Me.iCodigo, False)

    '        While i < Sociedades.Count And Not bEncontrado
    '            If Sociedad.Codigo = Sociedades(i).CodigoSociedad Then
    '                misSociedades.RemoveAt(i)
    '                miSocUser.Eliminar()
    '                bEncontrado = True
    '            End If
    '            i = i + 1
    '        End While
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Sub

    'Public Property OperacionActual As Operacion
    '    Get
    '        OperacionActual = Me.miOperacionActual
    '    End Get
    '    Set(value As Operacion)
    '        Me.miOperacionActual = value
    '    End Set
    'End Property

    'Public Property AgendaPersonal(ByVal Fecha As Date) As List(Of TareaAgenda)
    '    Get
    '        If Me.miAgendaPersonal.Count = 0 Then
    '            Me.miAgendaPersonal = Datos.DameAgenda(Me.iCodigo, New List(Of Perfil), Fecha, TipoAgenda.Personal, False)
    '        End If
    '        AgendaPersonal = Me.miAgendaPersonal
    '    End Get
    '    Set(value As List(Of TareaAgenda))
    '        Me.miAgendaPersonal = value
    '    End Set
    'End Property

    'Public Property AgendaGrupo(ByVal Fecha As Date) As List(Of TareaAgenda)
    '    Get
    '        If Me.miAgendaGrupo.Count = 0 Then
    '            Me.miAgendaGrupo = Datos.DameAgenda(Me.iCodigo, Perfiles, Fecha, TipoAgenda.Grupo, False)
    '        End If
    '        AgendaGrupo = Me.miAgendaGrupo
    '    End Get
    '    Set(value As List(Of TareaAgenda))
    '        Me.miAgendaGrupo = value
    '    End Set
    'End Property

    'Public Property AgendaSinAsignar(ByVal Fecha As Date) As List(Of TareaAgenda)
    '    Get
    '        If Me.miAgendaSinAsignar.Count = 0 Then
    '            Me.miAgendaSinAsignar = Datos.DameAgenda(0, New List(Of Perfil), Fecha, TipoAgenda.SinAsignar, False)
    '        End If
    '        AgendaSinAsignar = Me.miAgendaSinAsignar
    '    End Get
    '    Set(value As List(Of TareaAgenda))
    '        Me.miAgendaSinAsignar = value
    '    End Set
    'End Property

    'Public Property AgendaPendiente() As List(Of TareaAgenda)
    '    Get
    '        If Me.miAgendaPendiente.Count = 0 Then
    '            Me.miAgendaPendiente = Datos.DameAgenda(Me.iCodigo, New List(Of Perfil), FechaGlobal, TipoAgenda.Ninguno, True)
    '        End If
    '        AgendaPendiente = Me.miAgendaPendiente
    '    End Get
    '    Set(ByVal value As List(Of TareaAgenda))
    '        Me.miAgendaPendiente = value
    '    End Set
    'End Property


    'Public Function FechasAgendaCompleta() As List(Of Date)
    '    Try
    '        Dim miLista As New List(Of TareaAgenda)
    '        Dim miFecha As New Date

    '        TareaAgenda.OrdenarPor = "FechaEjecucion"
    '        FechasAgendaCompleta = New List(Of Date)
    '        miAgendaSinAsignar = New List(Of TareaAgenda)
    '        miAgendaPendiente = New List(Of TareaAgenda)
    '        miAgendaGrupo = New List(Of TareaAgenda)
    '        miAgendaPersonal = New List(Of TareaAgenda)

    '        miLista.AddRange(AgendaPendiente)
    '        miLista.AddRange(AgendaSinAsignar(CDate(FechaGlobal)))
    '        miLista.AddRange(AgendaGrupo(CDate(FechaGlobal)))
    '        miLista.AddRange(AgendaPersonal(CDate(FechaGlobal)))

    '        FechasAgendaCompleta = miLista.Select(Function(p As TareaAgenda) p.FechaEjecucion.Date).Distinct.ToList
    '        FechasAgendaCompleta.Sort()

    '        miAgendaSinAsignar = New List(Of TareaAgenda)
    '        miAgendaPendiente = New List(Of TareaAgenda)
    '        miAgendaGrupo = New List(Of TareaAgenda)

    '    Catch ex As Exception
    '        FechasAgendaCompleta = New List(Of Date)
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    'Public Function FechasTareasPendientes() As List(Of Date)
    '    Try
    '        Dim i As Integer = 0

    '        TareaAgenda.OrdenarPor = "FechaEjecucion"
    '        FechasTareasPendientes = New List(Of Date)

    '        Dim Fecha As Date = FechaGlobal

    '        While i < AgendaPendiente.Count
    '            If Fecha <> AgendaPendiente(i).FechaEjecucion Then
    '                Fecha = AgendaPendiente(i).FechaEjecucion
    '                FechasTareasPendientes.Add(Fecha)
    '            End If
    '            i = i + 1
    '        End While

    '        FechasTareasPendientes.Sort()
    '    Catch ex As Exception
    '        FechasTareasPendientes = New List(Of Date)
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    'Public Function Insertar_MensajeAgenda(Mensaje As String,
    '                                       CodPerfil As Integer,
    '                                       CodUsuario As Integer,
    '                                       CodCliente As Integer,
    '                                       CargoCliente As Integer) As Boolean
    '    Try
    '        Dim miAgenda = New TareaAgenda(0,
    '                                       CodPerfil,
    '                                       CodUsuario,
    '                                       Me.iCodigo,
    '                                       0,
    '                                       False,
    '                                       Now,
    '                                       Now,
    '                                       FechaGlobal,
    '                                       Now,
    '                                       CodCliente,
    '                                       CargoCliente,
    '                                       0,
    '                                       UTrim(Mensaje))

    '        Insertar_MensajeAgenda = miAgenda.Insertar()

    '        If Insertar_MensajeAgenda Then
    '            Me.AgendaPendiente.Add(miAgenda)
    '        End If

    '    Catch ex As Exception
    '        Insertar_MensajeAgenda = False
    '        Throw New NegocioDatosExcepction(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
    '    End Try
    'End Function

    Public Function DameTraduccionIdioma(Palabras As String) As String
        Try
            If Idioma.ContainsKey(UTrim(Palabras)) Then
                DameTraduccionIdioma = Idioma(UTrim(Palabras))
            Else
                DameTraduccionIdioma = Palabras
            End If
        Catch ex As Exception
            DameTraduccionIdioma = ""
        End Try
    End Function

    Public Property Idioma() As Dictionary(Of String, String)
        Get
            If Me.misPalabrasIdioma.Count = 0 Then
                Me.misPalabrasIdioma = Datos.DamePalabrasIdioma()
            End If

            Idioma = Me.misPalabrasIdioma
        End Get
        Set(value As Dictionary(Of String, String))
            Me.misPalabrasIdioma = value
        End Set
    End Property
End Class
