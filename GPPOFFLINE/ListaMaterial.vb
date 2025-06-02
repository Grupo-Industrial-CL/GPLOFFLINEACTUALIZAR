
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class ListaMaterial
#Region "Atributos"
    Private sCodigo As String
    Private iNodo As Integer
    Private sMaterial As String
    Private sPosicion As String
    Private dCantidad As Double
    Private sUnidadMedida As String
    Private sTipoPosicion As String
    Private dPorceMerma As Double
    Private miMaterial As Material
    Private miResumenMaterial As ResumenMaterial
    Private miCabecera As CabListaMaterial

    Private bCreado As Boolean
#End Region

#Region "Constructores"
    Private Sub InicializarVariables()
        Try
            sCodigo = ""
            iNodo = 0
            sMaterial = ""
            Me.miMaterial = Nothing
            sPosicion = ""
            dCantidad = 0
            sUnidadMedida = ""
            sTipoPosicion = ""
            dPorceMerma = 0
            miResumenMaterial = Nothing
            Me.miCabecera = Nothing

            bCreado = False

        Catch ex As Exception
            bCreado = False
        End Try
    End Sub

    Public Sub New()
        InicializarVariables()
    End Sub

    Public Sub New(Codigo As String,
                   Nodo As Integer,
                   Material As String,
                   Posicion As String,
                   Cantidad As Double,
                   UnidadMedida As String,
                   TipoPosicion As String,
                   PorcMerma As Double)
        Try
            InicializarVariables()
            Me.iNodo = Nodo
            Me.sCodigo = Codigo
            Me.sMaterial = Material
            Me.sPosicion = Posicion
            Me.dCantidad = Cantidad
            Me.sUnidadMedida = UnidadMedida
            Me.sTipoPosicion = TipoPosicion
            Me.dPorceMerma = PorcMerma
            Me.bCreado = True
        Catch ex As Exception
            Me.bCreado = True
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(Lista As String,
                   Nodo As Integer)
        Try
            Dim sSql As String = "SELECT * " &
                                 " FROM ListaMateriales " &
                                 " WHERE UPPER(RTRIM(dlLista)) = '" & UTrim(Lista) & "'," &
                                 " AND dlNodo = " & Nodo
            Dim DTDatos As New DataTable
            InicializarVariables()

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Me.sCodigo = Lista
                Me.iNodo = Nodo
                Me.sMaterial = UTrim(DTDatos.Rows(0).Item("dlMaterial"))
                Me.sPosicion = UTrim(DTDatos.Rows(0).Item("dlPosicion"))
                Me.dCantidad = CDbl(NoNull(DTDatos.Rows(0).Item("dlCantidad"), "D"))
                Me.sUnidadMedida = UTrim(DTDatos.Rows(0).Item("dlUM"))
                Me.sTipoPosicion = UTrim(DTDatos.Rows(0).Item("dlTipoPos"))
                Me.dPorceMerma = CDbl(NoNull(DTDatos.Rows(0).Item("dlPorcMerma"), "D"))
                Me.bCreado = True

            End If
        Catch ex As Exception
            Me.bCreado = True
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub
#End Region

#Region "Propiedades"

    Public Property Codigo As String
        Get
            Codigo = Me.sCodigo
        End Get
        Set(value As String)
            Me.sCodigo = value
            Me.miCabecera = New CabListaMaterial
        End Set
    End Property

    Public ReadOnly Property Cabecera As CabListaMaterial
        Get
            If miCabecera Is Nothing Then
                Me.miCabecera = New CabListaMaterial(Me.sCodigo)
            ElseIf miCabecera.Creado = False Then
                Me.miCabecera = New CabListaMaterial(Me.sCodigo)
            End If

            If miCabecera.Creado = False Then
                Return New CabListaMaterial
            Else
                Return miCabecera
            End If
        End Get
    End Property

    Public Property Nodo As Integer
        Get
            Nodo = Me.iNodo
        End Get
        Set(value As Integer)
            Me.iNodo = value
        End Set
    End Property

    Public Property CodigoMaterial As String
        Get
            CodigoMaterial = Me.sMaterial
        End Get
        Set(value As String)
            Me.sMaterial = value
            Me.miMaterial = New Material
        End Set
    End Property

    Public ReadOnly Property Material As Material
        Get

            If miMaterial Is Nothing Then
                Me.miMaterial = New Material(Me.sMaterial)
            ElseIf miMaterial.Creado = False Then
                Me.miMaterial = New Material(Me.sMaterial)
            End If

            Return miMaterial
        End Get
    End Property
    Public ReadOnly Property MaterialResumen As ResumenMaterial
        Get
            If miResumenMaterial Is Nothing Then
                Me.miResumenMaterial = New ResumenMaterial(Me.sMaterial)
            ElseIf miResumenMaterial.Creado = False Then
                Me.miResumenMaterial = New ResumenMaterial(Me.sMaterial)
            End If

            Return miResumenMaterial
        End Get
    End Property

    Public Property Posicion As String
        Get
            Posicion = Me.sPosicion
        End Get
        Set(value As String)
            Me.sPosicion = value
        End Set
    End Property

    Public Property Cantidad As Double
        Get
            Cantidad = Me.dCantidad
        End Get
        Set(value As Double)
            Me.dCantidad = value
        End Set
    End Property

    Public Property UnidadMedida As String
        Get
            UnidadMedida = Me.sUnidadMedida
        End Get
        Set(value As String)
            Me.sUnidadMedida = value
        End Set
    End Property

    Public Property TipoPosicion As String
        Get
            TipoPosicion = Me.sTipoPosicion
        End Get
        Set(value As String)
            Me.sTipoPosicion = value
        End Set
    End Property

    Public Property PorcentajeMerma As Double
        Get
            Return Me.dPorceMerma
        End Get
        Set(ByVal value As Double)
            Me.dPorceMerma = value
        End Set
    End Property


    Public ReadOnly Property Creado As Boolean
        Get
            Creado = Me.bCreado
        End Get
    End Property
#End Region



#Region "BBDD"
    Public Function Insertar() As Boolean
        Try
            Dim CodigoSQL As Integer
            Dim sSql As String = "INSERT INTO ListaMateriales (dlLista,dlNodo,dlMaterial,dlPosicion,dlCantidad,dlUM,dlTipoPos,dlPorcMerma) " &
                                 " VALUES ('" & UTrim(sCodigo) & "'," &
                                                iNodo & ",'" &
                                                sMaterial & "','" &
                                                sPosicion & "'," &
                                                dCantidad & ",'" &
                                                sUnidadMedida & "','" &
                                                sTipoPosicion & "'," &
                                                PuntoComa(Me.dPorceMerma) & ") SELECT @@IDENTITY "

            CodigoSQL = CInt(Datos.CGPL.EjecutarConsultaEscalar(sSql))

            If CodigoSQL = -1 Then
                Insertar = False
                bCreado = False
            Else
                Insertar = True
                Datos.GuardarLog(TipoLogDescripcion.Alta & " ListaMateriales", CStr(Codigo))
            End If

        Catch ex As Exception
            Insertar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Modificar() As Boolean
        Try
            Dim sSql As String = "UPDATE ListaMateriales " &
                                 " SET dlNodo = " & iNodo & ", " &
                                 " dlMaterial = '" & sMaterial & "', " &
                                 " dlPosicion = '" & sPosicion & "', " &
                                 " dlCantidad = " & dCantidad & ", " &
                                 " dlUM = '" & sUnidadMedida & "', " &
                                 " dlTipoPos = '" & sTipoPosicion & "," &
                                 " dlPorcMerma = " & PuntoComa(Me.dPorceMerma) &
                                 " WHERE dlLista=" & UTrim(sCodigo)

            Modificar = Datos.CGPL.EjecutarConsulta(sSql)
            If Modificar Then
                Datos.GuardarLog(TipoLogDescripcion.Modificar & " ListaMateriales", CStr(Codigo))
            End If
        Catch ex As Exception
            Modificar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

    Public Function Eliminar() As Boolean
        Try
            Dim sSql As String = "DELETE FROM ListaMateriales " &
                                 "WHERE dlLista=" & sCodigo

            Eliminar = Datos.CGPL.EjecutarConsulta(sSql)
            If Eliminar Then
                Datos.GuardarLog(TipoLogDescripcion.Eliminar & " ListaMateriales", CStr(Codigo))
            End If
        Catch ex As Exception
            Eliminar = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Function

#End Region
End Class
