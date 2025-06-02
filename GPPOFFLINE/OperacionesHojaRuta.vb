
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP
Public Class OperacionesHojaRuta

#Region "Atributos"

    Private bCreado As Boolean
    Private miPuestoTrabajo As PuestosTrabajo
    Private miHojaRuta As HojaRuta

#End Region

#Region "Constructores"

    Private Sub InicializarVariables()
        Try
            GrupoSAP = ""
            NodoSAP = 0
            ContOperSAP = 0
            NumOperacionSAP = ""

            UnidadMedida = String.Empty
            Nombre = String.Empty
            CantidadBase = 0
            Operarios = 0
            CodigoPuestoDeTrabajo = 0
            MinutosMaquina = 0
            ContGrupoSAP = ""
            ClaveControlSAP = ""
            MinutosLimpieza = 0
            MinutosPreparacion = 0

            'miPuestoTrabajo = New PuestosTrabajo
            'miHojaRuta = New HojaRuta

            MinutosLimpieza = 0
            MinutosPreparacion = 0

            CodProveedor = String.Empty

            bCreado = False
        Catch ex As Exception
            Me.bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub



    Public Sub New(GrupoHojaRuta As String,
                   Nodo_SAP As Integer,
                   ContOper_SAP As Integer,
                   NumOperacion_SAP As String,
                   sUnidadMedida As String,
                   sNombre As String,
                   iCantidadBase As Integer,
                   iOperarios As Integer,
                   iPuestoDeTrabajo As Integer,
                   dMinutosMaq As Decimal,
                   ContGrupo_SAP As String,
                   ClaveControl_SAP As String,
                   dMinutosPrep As Decimal,
                   dMinutosLimp As Decimal,
                   sCodProveedor As String)
        Try
            GrupoSAP = GrupoHojaRuta
            NodoSAP = Nodo_SAP
            ContOperSAP = ContOper_SAP
            NumOperacionSAP = NumOperacion_SAP

            UnidadMedida = sUnidadMedida
            Nombre = sNombre
            CantidadBase = iCantidadBase
            Operarios = iOperarios
            CodigoPuestoDeTrabajo = iPuestoDeTrabajo
            MinutosMaquina = dMinutosMaq
            MinutosPreparacion = dMinutosPrep
            MinutosLimpieza = dMinutosLimp
            ContGrupoSAP = ContGrupo_SAP
            ClaveControlSAP = ClaveControl_SAP
            MinutosLimpieza = dMinutosLimp
            MinutosPreparacion = dMinutosPrep
            CodProveedor = sCodProveedor

            bCreado = True

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)

        End Try
    End Sub



    Public Sub New(Grupo_SAP As String,
                   Nodo_SAP As Integer,
                   Cont_OperSAP As Integer)
        Try
            Dim sSQl As String = "SELECT * " &
                               " FROM OperacionesHojaRuta with(nolock) " &
                               " WHERE opGrupoSAP='" & Grupo_SAP.Trim &
                               "' AND opNodoSAP=" & Nodo_SAP &
                               " AND opContOperSAP=" & Cont_OperSAP

            Dim DTDatos As New DataTable

            InicializarVariables()
            If Datos.CGPL.DameDatosDT(sSQl, DTDatos) Then
                GrupoSAP = Grupo_SAP
                NodoSAP = Nodo_SAP
                ContOperSAP = Cont_OperSAP
                NumOperacionSAP = UTrim(DTDatos.Rows(0).Item("opNumOper"))
                UnidadMedida = UTrim(DTDatos.Rows(0).Item("opUnidadMedida"))
                Nombre = UTrim(DTDatos.Rows(0).Item("opNombre"))
                CantidadBase = CInt((NoNull(DTDatos.Rows(0).Item("opCantidadBase"), "N")))
                Operarios = CInt((NoNull(DTDatos.Rows(0).Item("opOperarios"), "N")))
                CodigoPuestoDeTrabajo = CInt((NoNull(DTDatos.Rows(0).Item("opPuestoTrabajo"), "N")))
                MinutosMaquina = CDec((NoNull(DTDatos.Rows(0).Item("opMinutos"), "D")))
                MinutosPreparacion = CDec((NoNull(DTDatos.Rows(0).Item("opMinutosPrep"), "D")))
                MinutosLimpieza = CDec((NoNull(DTDatos.Rows(0).Item("opMinutosLimpieza"), "D")))
                ContGrupoSAP = CStr((NoNull(DTDatos.Rows(0).Item("opContGrupoSAP"), "A")))
                ClaveControlSAP = CStr((NoNull(DTDatos.Rows(0).Item("opClaveControl"), "A")))
                MinutosLimpieza = CDec((NoNull(DTDatos.Rows(0).Item("opMinutosLimpieza"), "N")))
                MinutosPreparacion = CDec((NoNull(DTDatos.Rows(0).Item("opMinutosPrep"), "N")))
                CodProveedor = UTrim(DTDatos.Rows(0).Item("opProveedor"))

                Me.bCreado = True
            End If

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)

        End Try
    End Sub


#End Region

#Region "Propiedades"

    Public Property GrupoSAP As String
    Public Property NodoSAP As Integer
    Public Property ContOperSAP As Integer
    Public Property NumOperacionSAP As String
    Public Property UnidadMedida As String
    Public Property Nombre As String
    Public Property CantidadBase As Integer
    Public Property Operarios As Integer
    Public Property CodigoPuestoDeTrabajo As Integer
    Public Property MinutosMaquina As Decimal
    Public Property ContGrupoSAP As String
    Public Property ClaveControlSAP As String
    Public Property MinutosLimpieza As Decimal
    Public Property MinutosPreparacion As Decimal
    Public Property CodProveedor As String

    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property


    Public ReadOnly Property PuestoTrabajo As PuestosTrabajo
        Get
            If Me.miPuestoTrabajo Is Nothing And ClaveControlSAP <> ClaveControl_PtoTrabajo_Externo Then
                Me.miPuestoTrabajo = New PuestosTrabajo(CodigoPuestoDeTrabajo)
            End If

            If Me.miPuestoTrabajo Is Nothing And ClaveControlSAP = ClaveControl_PtoTrabajo_Externo Then
                Me.miPuestoTrabajo = New PuestosTrabajo(CInt(TipoPuestoTrabajo.Externo))
            End If
            PuestoTrabajo = miPuestoTrabajo
        End Get
    End Property

    Public ReadOnly Property HojaDeRuta As HojaRuta
        Get

            If miHojaRuta Is Nothing Then
                miHojaRuta = New HojaRuta(GrupoSAP,
                                          ContGrupoSAP)
            End If
            HojaDeRuta = miHojaRuta
        End Get
    End Property

#End Region



#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            Dim CodigoSQL As Integer
            Dim sSql As String = "INSERT INTO OperacionesHojaRuta (opGrupoSAP,opNodoSAP,opContOperSAP,opNumOper,opUnidadMedida,opNombre," &
                                 "opCantidadBase,opOperarios,opPuestoTrabajo,opMinutos,opContGrupoSAP,opClaveControl,opMinutosLimpieza,opMinutosPrep,opProveedor) " &
                                 " VALUES ('" & UTrim(GrupoSAP) & "'," &
                                               NodoSAP & "," &
                                               ContOperSAP & "," &
                                               "'" & NumOperacionSAP & "'," &
                                               "'" & UnidadMedida & "'," &
                                               "'" & Nombre & "'," &
                                               CantidadBase & "," &
                                               Operarios & "," &
                                               CodigoPuestoDeTrabajo & "," &
                                               MinutosMaquina & "," &
                                               "'" & ContGrupoSAP & "'," &
                                               "'" & ClaveControlSAP & "'," &
                                               MinutosLimpieza & "," &
                                               MinutosPreparacion & ",'" &
                                               CodProveedor & "') SELECT @@IDENTITY "

            CodigoSQL = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If CodigoSQL = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & " OperacionesHojaRuta", CStr(GrupoSAP))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE OperacionesHojaRuta " &
                                 " SET opContOperSAP = " & ContOperSAP & ", " &
                                 " opNumOper= '" & UTrim(NumOperacionSAP) & "', " &
                                 " opUnidadMedida= '" & UTrim(UnidadMedida) & "', " &
                                 " opNombre= '" & UTrim(Nombre) & "', " &
                                 " opCantidadBase= " & CantidadBase & ", " &
                                 " opOperarios= " & Operarios & ", " &
                                 " opPuestoTrabajo= " & CodigoPuestoDeTrabajo & ", " &
                                 " opMinutos= " & MinutosMaquina & ", " &
                                 " opContGrupoSAP= '" & UTrim(ContGrupoSAP) & "', " &
                                 " opClaveControl= '" & UTrim(ClaveControlSAP) & "'," &
                                 " opMinutosLimpieza = " & MinutosLimpieza & "," &
                                 " opMinutosPrep = " & MinutosPreparacion & "," &
                                 " opProveedor = '" & CodProveedor &
                                 " WHERE opGrupoSAP=" & UTrim(GrupoSAP) & " AND opNodoSAP=" & NodoSAP

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " OperacionesHojaRuta", CStr(ContOperSAP))
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM OperacionesHojaRuta " &
                                 "WHERE opGrupoSAP=" & ContOperSAP & " AND opNodoSAP=" & NodoSAP

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " OperacionesHojaRuta", CStr(ContOperSAP))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

#End Region

End Class
