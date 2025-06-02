

Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP

Public Class PullSystem
#Region "Atributos "
    Private sMaterial As String
    Private miMaterial As Material
    Private miResumenMaterial As ResumenMaterial
    Private bCreado As Boolean
    Private miMesLaboral As DiasLaborales
    Private miPrevicionDiaria As Integer
    Private miCalculoStockMinimo As Integer
    Private miLoteOptimo As Integer
    Private miFabricacionPendiente As Integer
    Private miSituacionActual As Integer
    Private miNuevaFabricacion As Integer
    Private iStockActual As Double
    Private iFechaRotura As Date
    Private iFechaRoturaEntradas As Date
    Private misFabricaciones As List(Of Fabricaciones)

    Private diasControl As Integer
    Private iStockBloqueado As Double
    Private sEstatus As String
    Private iFechaCorta As Date
    Private iNecesidad As String

#End Region

#Region "Constructores"
    Private Sub InicializarVariables()
        Try
            sMaterial = ""
            miMaterial = Nothing
            miResumenMaterial = Nothing
            miMesLaboral = New DiasLaborales
            Mes = 0
            Año = 0
            miPrevicionDiaria = 0
            miCalculoStockMinimo = 0
            miLoteOptimo = 0
            bCreado = False
            misFabricaciones = New List(Of Fabricaciones)
            miFabricacionPendiente = 0
            miSituacionActual = 0
            miNuevaFabricacion = 0
            diasControl = 0
            Me.iStockActual = -999999999
            fecha_Fin_Previsto = FechaGlobal
        Catch ex As Exception
            bCreado = False
        End Try
    End Sub

    Public Sub New(sCodigoMaterial As String,
                   iMes As Integer,
                   iAño As Integer,
                   iCantidad As Integer,
                   idiasControl As Integer,
                   ByVal StockActual As Double)
        Try
            InicializarVariables()

            Me.sMaterial = sCodigoMaterial
            Mes = iMes
            Año = iAño
            Cantidad = iCantidad
            diasControl = idiasControl
            fecha_Fin_Previsto = FechaGlobal
            iStockActual = StockActual
            Me.bCreado = True
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(sCodigoMaterial As String,
                   iMes As Integer,
                   iAño As Integer,
                   iCantidad As Integer,
                   idiasControl As Integer,
                   ByVal StockActual As Double,
                   ByVal FechaRotura As Date,
                   ByVal miStockBloqueado As Double,
                   ByVal miEstatus As String,
                   ByVal FechaCorta As Date,
                   ByVal Necesidad As String)
        Try
            InicializarVariables()

            Me.sMaterial = sCodigoMaterial
            Mes = iMes
            Año = iAño
            Cantidad = iCantidad
            diasControl = idiasControl
            fecha_Fin_Previsto = FechaGlobal
            iStockActual = StockActual
            iFechaRotura = FechaRotura
            iStockBloqueado = miStockBloqueado
            sEstatus = miEstatus
            iFechaCorta = FechaCorta
            iNecesidad = Necesidad
            Me.bCreado = True
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(sCodigoMaterial As String,
                   iMes As Integer,
                   iAño As Integer,
                   iCantidad As Integer,
                   idiasControl As Integer,
                   ByVal StockActual As Double,
                   ByVal FechaRotura As Date,
                    ByVal FechaRoturaEntradas As Date,
                   ByVal miStockBloqueado As Double,
                   ByVal miEstatus As String,
                   ByVal FechaCorta As Date,
                   ByVal Necesidad As String)
        Try
            InicializarVariables()

            Me.sMaterial = sCodigoMaterial
            Mes = iMes
            Año = iAño
            Cantidad = iCantidad
            diasControl = idiasControl
            fecha_Fin_Previsto = FechaGlobal
            iStockActual = StockActual
            iFechaRotura = FechaRotura
            iFechaRoturaEntradas = FechaRoturaEntradas
            iStockBloqueado = miStockBloqueado
            sEstatus = miEstatus
            iFechaCorta = FechaCorta
            iNecesidad = Necesidad
            Me.bCreado = True
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(sCodigoMaterial As String,
                   iMes As Integer,
                   iAño As Integer,
                   iCantidad As Integer,
                   idiasControl As Integer,
                   ByVal StockActual As Double,
                   ByVal FechaRotura As Date,
                   ByVal miStockBloqueado As Double,
                   ByVal miEstatus As String
                   )
        Try
            InicializarVariables()

            Me.sMaterial = sCodigoMaterial
            Mes = iMes
            Año = iAño
            Cantidad = iCantidad
            diasControl = idiasControl
            fecha_Fin_Previsto = FechaGlobal
            iStockActual = StockActual
            iFechaRotura = FechaRotura
            iStockBloqueado = miStockBloqueado
            sEstatus = miEstatus

            Me.bCreado = True
        Catch ex As Exception
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub

    Public Sub New(sCodigoMaterial As String,
                   iMes As Integer,
                   iAño As Integer)
        Try
            Dim sSql As String = " SELECT * " &
                                 " FROM ForeCastVentas " &
                                 " WHERE fcMaterial = '" & sCodigoMaterial & "'" &
                                 " AND fcMes = " & Mes &
                                 " AND fcAnio = " & Año

            Dim DTDatos As New DataTable
            InicializarVariables()

            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then
                Me.sMaterial = sCodigoMaterial
                Mes = iMes
                Año = iAño
                Cantidad = CInt(NoNull(DTDatos.Rows(0).Item("fcCantidad"), "D"))
                fecha_Fin_Previsto = FechaGlobal
                Me.bCreado = True
            End If
        Catch ex As Exception
            InicializarVariables()
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "() ", ex)
        End Try
    End Sub
#End Region

#Region "Propiedades"

    Public ReadOnly Property Estatus As String
        Get
            Return Me.sEstatus
        End Get
    End Property
    Public ReadOnly Property StockBloqueado As Double
        Get
            Return Me.iStockBloqueado
        End Get
    End Property
    Public ReadOnly Property Stock As Double
        Get
            Return Me.iStockActual
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
    Public Property Mes As Integer
    Public Property Año As Integer
    Public Property Cantidad As Integer

    Public ReadOnly Property Creado As Boolean
        Get
            Creado = bCreado
        End Get
    End Property
    Public Property CodigoMaterial As String
        Get
            CodigoMaterial = Me.sMaterial
        End Get
        Set(value As String)
            Me.sMaterial = value
        End Set
    End Property

    Public ReadOnly Property Material As Material
        Get
            If miMaterial Is Nothing Then
                miMaterial = New Material(Me.sMaterial)
            ElseIf miMaterial.Creado = False Then
                miMaterial = New Material(Me.sMaterial)
            End If

            Return miMaterial
        End Get
    End Property

    Public ReadOnly Property DiasLaborales As DiasLaborales
        Get
            If miMesLaboral.Creado = False And Mes > 0 Then
                miMesLaboral = New DiasLaborales(Mes)
            End If
            Return miMesLaboral
        End Get
    End Property
    Public Property fecha_Fin_Previsto As Date

    Public ReadOnly Property Fabricaciones As List(Of Fabricaciones)
        Get
            If diasControl > 0 Then

                If fecha_Fin_Previsto = FechaGlobal Then

                    misFabricaciones = DatosPullSystem.DameFabricaciones(EstadosFabricacion:=EstadoFabricacion.PteFabricar & "," &
                                                                                      EstadoFabricacion.EnMarcha,
                                                                  FechaPrevistaFIN:=Now.AddDays(diasControl).Date,
                                                                  CodMaterial:=CodigoMaterial.Trim)
                Else
                    If Now.AddDays(diasControl) >= fecha_Fin_Previsto Then

                        misFabricaciones = DatosPullSystem.DameFabricaciones(EstadosFabricacion:=EstadoFabricacion.PteFabricar & "," &
                                                                                      EstadoFabricacion.EnMarcha, CodMaterial:=CodigoMaterial.Trim)
                    End If


                End If


            End If
            Return misFabricaciones
        End Get

    End Property


    Public ReadOnly Property PrevicionDiaria As Integer
        Get
            If Cantidad > 0 And DiasLaborales.Creado = True Then
                miPrevicionDiaria = CInt(Cantidad / DiasLaborales.DiasLaborales)
            End If
            Return miPrevicionDiaria
        End Get

    End Property


    Public ReadOnly Property CalculoStockMinimo As Integer
        Get
            If miResumenMaterial.Creado Then
                If miResumenMaterial.StockMinPS > 0 Then
                    miCalculoStockMinimo = miResumenMaterial.StockMinPS
                ElseIf PrevicionDiaria > 0 And miResumenMaterial.DiasPP > 0 Then

                    miCalculoStockMinimo = PrevicionDiaria * miResumenMaterial.DiasPP
                End If

                Return miCalculoStockMinimo
            Else
                Return 0
            End If

        End Get
    End Property

    Public ReadOnly Property LoteOptimo As Integer
        Get
            If miResumenMaterial.Creado = True Then

                If miResumenMaterial.StockMaxPS > 0 And miCalculoStockMinimo > 0 Then
                    miLoteOptimo = miResumenMaterial.StockMaxPS - miCalculoStockMinimo
                End If
                Return miLoteOptimo
            Else
                Return 0
            End If
        End Get
    End Property


    Public ReadOnly Property FabricacionPendiente As Integer
        Get
            If Fabricaciones.Count > 0 Then
                miFabricacionPendiente = Fabricaciones.Sum((Function(P) (P.CantidadPlanificada - P.CantidadFabBuenas)))
            End If
            Return miFabricacionPendiente
        End Get
    End Property



    Public ReadOnly Property SituacionActual As Integer
        Get
            miSituacionActual = MaterialResumen.Stock + FabricacionPendiente

            Return miSituacionActual
        End Get
    End Property

    Public ReadOnly Property UnidadesFabricar As Integer
        Get
            If miResumenMaterial.Creado = True Then
                'If miResumenMaterial.StockMaxPS > 0 Then
                '    If Cantidad > miResumenMaterial.StockMaxPS Then
                '        miNuevaFabricacion = Cantidad - miSituacionActual
                '    Else
                '        If SituacionActual < miResumenMaterial.StockMaxPS Then
                '            miNuevaFabricacion = miResumenMaterial.StockMaxPS - miSituacionActual
                '        Else
                '            miNuevaFabricacion = 0
                '        End If

                '    End If
                'End If



                'If miResumenMaterial.StockMaxPS > 0 Then

                '    If SituacionActual < miResumenMaterial.StockMaxPS Then
                '        miNuevaFabricacion = miResumenMaterial.StockMaxPS - miSituacionActual
                '    Else
                '        miNuevaFabricacion = 0
                '    End If
                'End If

                miNuevaFabricacion = Cantidad - miSituacionActual

                Return miNuevaFabricacion
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property KgNuevaFabricacion As Double
        Get
            Dim kgFabricacion As Double = 0
            If miNuevaFabricacion > 0 Then
                Dim cantidadFormulaBase = miMaterial.CabLista_Material.MaterialesLista.Where(Function(p) p.MaterialResumen.Tipo = TipoMaterial.Fabricaciones).Sum(Function(p) p.Cantidad)
                If cantidadFormulaBase > 0 Then
                    kgFabricacion = (miNuevaFabricacion * cantidadFormulaBase) / miMaterial.CabLista_Material.CantidadBase
                End If
            End If
            Return kgFabricacion
        End Get
    End Property
    Public ReadOnly Property ValorRdoTanque As Integer
        Get
            ' esto es temporal se tiene que rediseñar con la alta de Tanques y ahi debe tener como
            ' propiedad la cacidad del mismo y al centro producutivo al que pertenece
            Dim TanqueRedondeo As Integer = 0
            If miMaterial.HojasRutaDefecto Is Nothing Then
                Return TanqueRedondeo
            End If

            Dim centrosMaquina = miMaterial.HojasRutaDefecto.PuestosTrabajoHojaRuta.Where(Function(p) p.Tipo = TipoPuestoTrabajo.Maquina).ToList

            If centrosMaquina.Count > 0 Then
                Select Case centrosMaquina.First.CodCentroProd
                    Case 1 ' Cosmetica
                        'Cosmética: 200 Kg, 500 Kg, 3.000 Kg y 6.000 Kg.
                        Select Case KgNuevaFabricacion
                            Case 0
                                TanqueRedondeo = 0
                            Case 1 To 200
                                TanqueRedondeo = TanqueCosmetica.TC200
                            Case 201 To 500
                                TanqueRedondeo = TanqueCosmetica.TC500
                            Case 501 To 3000
                                TanqueRedondeo = TanqueCosmetica.TC3000
                            Case 3001 To 6000
                                TanqueRedondeo = TanqueCosmetica.TC6000
                            Case Else
                                TanqueRedondeo = TanqueCosmetica.TC6000
                        End Select

                    Case 2 ' Fragancias
                        ' No hay tanques para fragancias
                        TanqueRedondeo = 0
                    Case 3 ' Higiene 
                        'Higiene: 3.000 Kg, 9.000 Kg y 20.000 Kg
                        Select Case KgNuevaFabricacion
                            Case 0
                                TanqueRedondeo = 0
                            Case 1 To 3000
                                TanqueRedondeo = TanqueHigiene.TH3000
                            Case 3001 To 9000
                                TanqueRedondeo = TanqueHigiene.TH9000
                            Case 9001 To 20000
                                TanqueRedondeo = TanqueHigiene.TH20000
                            Case Else
                                TanqueRedondeo = TanqueHigiene.TH20000
                        End Select
                End Select
            End If
            Return TanqueRedondeo
        End Get
    End Property
    Public ReadOnly Property NuevaFabricacion As Double
        Get
            Dim UNFabricacion As Double = 0
            Dim cantidadFormulaBase = miMaterial.CabLista_Material.MaterialesLista.Where(Function(p) p.MaterialResumen.Tipo = TipoMaterial.Fabricaciones).Sum(Function(p) p.Cantidad)

            If ValorRdoTanque > 0 Then
                If cantidadFormulaBase > 0 Then
                    UNFabricacion = CInt((ValorRdoTanque * miMaterial.CabLista_Material.CantidadBase)) / cantidadFormulaBase
                End If
            End If

            Return UNFabricacion
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

    Public ReadOnly Property FechaCorta As Date
        Get
            Return Me.iFechaCorta
        End Get
    End Property

    Public ReadOnly Property Necesidad As String
        Get
            Return Me.iNecesidad
        End Get

    End Property
    'Higiene: 3.000 Kg, 9.000 Kg y 20.000 Kg

    'Cosmética: 200 Kg, 500 Kg, 3.000 Kg y 6.000 Kg.

#End Region
End Class
