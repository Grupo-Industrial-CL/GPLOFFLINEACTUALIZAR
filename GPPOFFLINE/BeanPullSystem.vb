
Option Explicit On
Option Strict On
Imports GPLOFFLINEACTUALIZAR.Util
Imports GPLOFFLINEACTUALIZAR.DatosGPP
Imports GPLOFFLINEACTUALIZAR.ConstantesGPP


Public Class BeanPullSystem
#Region "Atributos"

    Private sCodigo As String
    Private sNombre As String
    Private sFormato As String
    Private sGrupoPlanif As String
    Private iCantidad As Integer
    Private iDiasPP As Integer
    Private iStockMinimo As Integer
    Private iStockMaximo As Integer
    Private iStockActual As Double
    Private iFabricacionesPendientes As Integer
    Private iSituacionActual As Integer
    Private iUnidadesFabricar As Integer
    Private iKgNuevaFabricacion As Integer
    Private iValorRdoTanque As Integer
    Private iNuevaFabricacion As Integer
    Private dtFechaFin As Date
    Private iValorRedondeo As Integer
    Private iMesPS As Integer
    Private iAnioPS As Integer

    Private sHojaRuraDefault As String
    Private sContHojaRutaDefault As String
    Private iCodPuestoTrabajo As Integer
    Private sNombrePuestoTrabajo As String
    Private iCodCentroProd As Integer
    Private iCantidadBaseFormula As Integer
    Private iCantidadBaseMaterias As Integer
    Private iDiasLaborables As Integer
    Private dtFechaFinFanPendientes As Date
    Private iCantidadPendientePedVentas As Integer

    Private bCreado As Boolean
    Private sNombresPuestoTrabajo As List(Of PuestosTrabajo)
    Private sPuestoTrabajo As String

    Private sDetallePullSystem As List(Of BeanPullSystem)
    Private iStockBloqueado As Double
    Private iEstatus As String

    Private iFechaRotura As Date
    Private sNecesidad As String
    Private sFechaCorta As Date
    Private sCodigoMaterialFab As String


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
            sCodigo = ""
            sNombre = ""
            sFormato = ""
            sGrupoPlanif = ""
            iCantidad = 0
            iDiasPP = 0
            iStockMinimo = 0
            iStockMaximo = 0
            iStockActual = -99999
            iFabricacionesPendientes = 0
            iSituacionActual = 0
            iUnidadesFabricar = 0
            iKgNuevaFabricacion = 0
            iValorRdoTanque = 0
            iNuevaFabricacion = 0
            dtFechaFin = FechaGlobal
            iValorRedondeo = 0
            iMesPS = 0
            iAnioPS = 0

            sHojaRuraDefault = ""
            sContHojaRutaDefault = ""
            iCodPuestoTrabajo = 0
            sNombrePuestoTrabajo = ""
            iCodCentroProd = 0
            iCantidadBaseFormula = 0
            iCantidadBaseMaterias = 0
            iDiasLaborables = 0
            dtFechaFinFanPendientes = FechaGlobal
            bCreado = False
            iCantidadPendientePedVentas = 0
            iFechaRotura = FechaGlobal
            sFechaCorta = FechaGlobal
            sNecesidad = ""

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub

    Public Sub New(ByVal diasControl As Integer,
                   ByVal stockActual As Double,
                   ByVal fechaPrevistaFin As Date,
                   ByVal cantidadFC As Integer,
                   ByVal codMaterial As String,
                   ByVal MesPS As Integer,
                   ByVal AnioPS As Integer,
                   ByVal DiasLaborables As Integer,
                   ByVal DetallePullSystem As List(Of BeanPullSystem), ByVal Necesidad As String, ByVal fechaCorta As Date)
        Try

            Dim sSql As String = " " &
                                "Declare @diasControl int = " & diasControl & " " &
                                "Declare @StockActual As float = " & stockActual & " " &
                                "Declare @FechaPrevistaFin  As datetime  = '" & fechaPrevistaFin & "' " & " " &
                                "declare @CantidadForeCast as int= " & cantidadFC & " " &
                                "declare @MaterialPlan	  varchar(10)='" & codMaterial & "' " & " " &
                                "--Parametros internos para calculos

                                declare @CantidadPlanificada as int
                                declare @CantidadFabricada as int
                                declare @FabricacionPendiente as int
                                declare @SituacionActual as int
                                declare @UnidadesFabricar as int
                                declare @cantidadFormulaBase as int
                                declare @cantidadBase as int
                                declare @KgNuevaFabricacion as int
                                declare @FechaFinFabCurso as datetime
    
                                if  @diasControl >0 
                                begin
	                                if @FechaPrevistaFin ='01/01/1900'
	                                begin
		                                SELECT 
		                                @CantidadPlanificada= SUM(FAB.opCantidadPlanif)  ,
		                                @CantidadFabricada = SUM(FAB.opCantidadFabricadaSAPV2) --opCantidadFabBuenas
		                                FROM Fabricaciones FAB 
		                                WHERE  opFechaPrevFin <= DATEADD(day,@diasControl,GETDATE())  AND  
		                                opMaterialPadre = @MaterialPlan  AND  opEnmarcha IN (0,1,3)  
		
                                        SET @FechaFinFabCurso = (SELECT TOP 1 opFechaFin   FROM Fabricaciones with(nolock) where opMaterial = @MaterialPlan and opEnmarcha IN (0,1)  order by opFechaFin desc )					                                
                                        SET @FabricacionPendiente = ISNULL( @CantidadPlanificada,0)- isnull(@CantidadFabricada,0)
	                                end
	                                else
	                                begin
		                                if  DATEADD(day,@diasControl,GETDATE()) >=  @FechaPrevistaFin
		                                begin 
			                                SELECT 
			                                @CantidadPlanificada= SUM(FAB.opCantidadPlanif)  ,
			                                @CantidadFabricada = SUM(FAB.opCantidadFabricadaSAPV2)  --opCantidadFabBuenas
			                                FROM Fabricaciones FAB 
			                                WHERE  opMaterialPadre = @MaterialPlan  AND  opEnmarcha IN (0,1,3) 

                                            SET @FechaFinFabCurso = (SELECT TOP 1 opFechaFin   FROM Fabricaciones with(nolock) where opMaterial = @MaterialPlan and opEnmarcha IN (0,1)  order by opFechaFin desc )					
			                                SET @FabricacionPendiente =  ISNULL( @CantidadPlanificada,0)- isnull(@CantidadFabricada,0)
		                                end
	                                end 
                                end
        
                                SET @SituacionActual = @StockActual+@FabricacionPendiente
                                SET @UnidadesFabricar= @CantidadForeCast - @SituacionActual


                                select @cantidadFormulaBase=SUM(LM.dlCantidad )
                                from ListaMateriales LM  with(nolock) 
                                LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
                                where MA.maCod=@MaterialPlan and LM.dlUM='KG'

                                --Si la cantidad base es nula se busca en su lista de materiales para obtener esa cantidad base								
								if @cantidadFormulaBase is null 
								begin
								 select @cantidadFormulaBase= SUM(LM.dlCantidad )
																from ListaMateriales LM  with(nolock) 
																where dlUM = 'KG' and dlLista in ( select maListaMaterial
																from ListaMateriales LM  with(nolock) 
																LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
																where MA.maCod in (select dlMaterial
																from ListaMateriales LM  with(nolock) 
																LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
																where MA.maCod=@MaterialPlan))
								end
    
                                select top 1 @cantidadBase  =clCantidad 
                                from ListaMaterialesCab where clLista = (select top 1 maListaMaterial from Materiales  where maCod = @MaterialPlan)

                                if  @UnidadesFabricar >0
	                            begin
	                            set @KgNuevaFabricacion= (@UnidadesFabricar * @cantidadFormulaBase)/ @cantidadBase
	                            end
	                            else
	                            begin
	                            set @KgNuevaFabricacion=0
	                            end

                                SELECT 
                                MA.maCod as Codigo,
                                MA.maNombre as NombreMat ,
                                MA.maGrupoHR as HojaRuraDefault,
                                MA.maContHR  as ContHojaRutaDefault,
                                isnull(F.fmNombre,'')  as Formato ,
                                isnull(F.fmCod ,'')  as GrupoPlanif ,
                                @CantidadForeCast as CantidadFC,
                                isnull(MA.maDiasPP,0)  as DiasPP,
                                MA.maStokMinPS as StockMin ,
                                MA.maStokMaxPS as StockMax,
                                isnull(@StockActual,0) as StockActual,
                                isnull(@FabricacionPendiente,0) AS  FabriacionPendiente,
                                @FechaFinFabCurso as FechaFinFabricacionActual,
                                isnull(@SituacionActual,0) as  SituacionActual,
                                isnull(@UnidadesFabricar,0) as  NuevaFabricacion,
                                isnull(@KgNuevaFabricacion,0) as KgNuevaFabricacion,
                                isnull(@cantidadFormulaBase,0) as CantidadFormulaBase,
                                isnull(@cantidadBase,0) as CantidadBase,
                                (select PT.ptCod  from PuestosTrabajo PT  
                                where ptCod = 
                                (select top 1 opPuestoTrabajo  from OperacionesHojaRuta  
                                 where opGrupoSAP=  MA.maGrupoHR and opContOperSAP=MA.maContHR)) as CodPuestoTrabajo,
                                (select PT.ptNombre   from PuestosTrabajo PT  
                                where ptCod = 
                                (select top 1 opPuestoTrabajo  from OperacionesHojaRuta  
                                 where opGrupoSAP=  MA.maGrupoHR and opContOperSAP=MA.maContHR)) as NombrePuestoTrabajo,
                                (select PT.ptCentroProd    from PuestosTrabajo PT  
                                where ptCod = 
                                (select top 1 opPuestoTrabajo  from OperacionesHojaRuta  
                                 where opGrupoSAP=  MA.maGrupoHR and opContOperSAP=MA.maContHR)) as CodCentroProd,
                                MA.maRedondeo  as RedondeoLote
                                from materiales MA  with(nolock) 
                                LEFT JOIN HojaRuta HR with(nolock)  on MA.maGrupoHR= HR.hrGrupo and MA.maContHR = HR.hrContGrupo 
                                LEFT JOIN Formatos F with(nolock)  on HR.hrFormato = F.fmCod 
                                where MA.maCod =@MaterialPlan"

            Dim DTDatos As New DataTable
            Dim DTDatosPuestos As New DataTable
            InicializarVariables()








            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then

                'Dim sSqlPuestos As String = " " &
                '                " select PT.ptNombre as PuestosTrabajo , PT.ptCod as Codigo  from PuestosTrabajo PT  
                '                where ptCod in
                '                (select  opPuestoTrabajo  from OperacionesHojaRuta  
                '                 where opGrupoSAP in  (select maGrupoHR
                '                from ListaMateriales LM  with(nolock) 
                '                LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
                '                where MA.maCod='" & UTrim(DTDatos.Rows(0).Item("Codigo")) & "' ) and opContOperSAP in (select maContHR
                '                from ListaMateriales LM  with(nolock) 
                '                LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
                '                where MA.maCod='" & UTrim( DTDatos.Rows(0).Item("Codigo")) & "' ))"

                Dim sSqlPuestos As String = " " &
                                " select PT.ptNombre as PuestosTrabajo , PT.ptCod as Codigo  from PuestosTrabajo PT  
                                where ptCod in
                                (select  opPuestoTrabajo  from OperacionesHojaRuta  
                                 where opGrupoSAP in  (select maGrupoHR
                                from ListaMateriales LM  with(nolock) 
                                LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
                                where MA.maCod='" & UTrim(DTDatos.Rows(0).Item("Codigo")) & "' ))"

                If UTrim(DTDatos.Rows(0).Item("Codigo")) = "70907491" Then
                    Dim g = 7
                End If

                Dim listaPuestosTrabajo As New List(Of PuestosTrabajo)
                If Datos.CGPL.DameDatosDT(sSqlPuestos, DTDatosPuestos) Then
                    'If DTDatos.Rows.Count > 1 Then
                    For index = 0 To DTDatosPuestos.Rows.Count - 1
                        Dim pue As New PuestosTrabajo
                        pue.Nombre = UTrim(DTDatosPuestos.Rows(index).Item("PuestosTrabajo"))
                        pue.CodigoPuestoTrabajo = CInt(NoNull(DTDatosPuestos.Rows(index).Item("Codigo"), "N"))
                        listaPuestosTrabajo.Add(pue)
                        'listaPuestosTrabajo.Add(UTrim(DTDatosPuestos.Rows(index).Item("PuestosTrabajo")))
                    Next
                    'End If
                End If
                sDetallePullSystem = DetallePullSystem
                sNombresPuestoTrabajo = listaPuestosTrabajo
                sCodigo = UTrim(DTDatos.Rows(0).Item("Codigo"))
                sNombre = UTrim(DTDatos.Rows(0).Item("NombreMat"))
                sFormato = UTrim(DTDatos.Rows(0).Item("Formato"))
                sGrupoPlanif = UTrim(DTDatos.Rows(0).Item("GrupoPlanif"))
                iCantidad = cantidadFC
                iDiasPP = CInt(NoNull(DTDatos.Rows(0).Item("DiasPP"), "N"))
                iStockMinimo = CInt(NoNull(DTDatos.Rows(0).Item("StockMin"), "N"))
                iStockMaximo = CInt(NoNull(DTDatos.Rows(0).Item("StockMax"), "N"))
                iStockActual = stockActual
                iFabricacionesPendientes = CInt(NoNull(DTDatos.Rows(0).Item("FabriacionPendiente"), "N"))
                iSituacionActual = CInt(NoNull(DTDatos.Rows(0).Item("SituacionActual"), "N"))
                iUnidadesFabricar = CInt(NoNull(DTDatos.Rows(0).Item("NuevaFabricacion"), "N"))
                iKgNuevaFabricacion = CInt(NoNull(DTDatos.Rows(0).Item("KgNuevaFabricacion"), "N"))
                dtFechaFin = fechaPrevistaFin
                iValorRedondeo = CInt(NoNull(DTDatos.Rows(0).Item("RedondeoLote"), "N"))
                iMesPS = MesPS
                iAnioPS = AnioPS

                sHojaRuraDefault = UTrim(DTDatos.Rows(0).Item("HojaRuraDefault"))
                sContHojaRutaDefault = UTrim(DTDatos.Rows(0).Item("ContHojaRutaDefault"))
                iCodPuestoTrabajo = CInt(NoNull(DTDatos.Rows(0).Item("CodPuestoTrabajo"), "N"))
                sNombrePuestoTrabajo = UTrim(DTDatos.Rows(0).Item("NombrePuestoTrabajo"))
                iCodCentroProd = CInt(NoNull(DTDatos.Rows(0).Item("CodCentroProd"), "N"))
                iCantidadBaseFormula = CInt(NoNull(DTDatos.Rows(0).Item("CantidadFormulaBase"), "N"))
                iCantidadBaseMaterias = CInt(NoNull(DTDatos.Rows(0).Item("CantidadBase"), "N"))
                iDiasLaborables = DiasLaborables
                dtFechaFinFanPendientes = CDate(NoNull(DTDatos.Rows(0).Item("FechaFinFabricacionActual"), "DT"))

                sNecesidad = Necesidad
                sFechaCorta = fechaCorta
                Me.bCreado = True
            End If

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub


    Public Sub New(ByVal diasControl As Integer,
                   ByVal stockActual As Double,
                   ByVal fechaPrevistaFin As Date,
                   ByVal cantidadFC As Integer,
                   ByVal codMaterial As String,
                   ByVal MesPS As Integer,
                   ByVal AnioPS As Integer,
                   ByVal DiasLaborables As Integer,
                   ByVal DetallePullSystem As List(Of BeanPullSystem),
                   ByVal FechaRotura As Date, ByVal StockBloqueado As Double, ByVal Estatus As String)
        Try

            Dim sSql As String = " " &
                                "Declare @diasControl int = " & diasControl & " " &
                                "Declare @StockActual As float = " & stockActual & " " &
                                "Declare @FechaPrevistaFin  As datetime  = '" & fechaPrevistaFin & "' " & " " &
                                "declare @CantidadForeCast as int= " & cantidadFC & " " &
                                "declare @MaterialPlan	  varchar(10)='" & codMaterial & "' " & " " &
                                "--Parametros internos para calculos

                                declare @CantidadPlanificada as int
                                declare @CantidadFabricada as int
                                declare @FabricacionPendiente as int
                                declare @SituacionActual as int
                                declare @UnidadesFabricar as int
                                declare @cantidadFormulaBase as int
                                declare @cantidadBase as int
                                declare @KgNuevaFabricacion as int
                                declare @FechaFinFabCurso as datetime
                                DECLARE  @opOrdenEnvSAP as int 
                                declare @CantidadFabricadaLOCAL as int

                                if  @diasControl >0 
                                begin
	                                if @FechaPrevistaFin ='01/01/1900'
	                                begin
		                                SELECT 
		                                @CantidadPlanificada= SUM(FAB.opCantidadPlanif)  ,
		                                @CantidadFabricada = SUM(FAB.opCantidadFabricadaSAPV2)--opCantidadFabBuenas
		                                FROM Fabricaciones FAB 
		                                WHERE  --opFechaPrevFin <= DATEADD(day,@diasControl,GETDATE())  AND  
		                                opMaterial = @MaterialPlan  AND  opEnmarcha IN (0,1)  
		
                                            set @CantidadFabricada = 0									
                                           set  @opOrdenEnvSAP  = 0;  
                                           
  
  
                                        DECLARE vendor_cursor CURSOR FOR   
                                         SELECT distinct FAB.opOrdenEnvSAP
                                         FROM Fabricaciones FAB  with (nolock)
                                         WHERE  opMaterial =  @MaterialPlan  AND  opEnmarcha IN (0,1) 
 
  
                                        OPEN vendor_cursor  
  
                                        FETCH NEXT FROM vendor_cursor   
                                        INTO @opOrdenEnvSAP 
  
                                        WHILE @@FETCH_STATUS = 0  
                                        BEGIN  

	                                        set @CantidadFabricadaLOCAL = 0

                                           SELECT @CantidadFabricadaLOCAL = FAB.opCantidadFabricadaSAPV2 
			                                                                        FROM Fabricaciones FAB 
			                                                                        WHERE  opOrdenEnvSAP =  @opOrdenEnvSAP
											                                        group by opOrdenEnvSAP,opCantidadFabricadaSAPV2 
											     
                                             set @CantidadFabricada = @CantidadFabricada + @CantidadFabricadaLOCAL
  
                                                -- Get the next vendor.  
                                            FETCH NEXT FROM vendor_cursor   
                                            INTO @opOrdenEnvSAP
                                        END   
                                        CLOSE vendor_cursor;  
                                        DEALLOCATE vendor_cursor;  

                                        SET @FechaFinFabCurso = (SELECT TOP 1 opFechaFin   FROM Fabricaciones with(nolock) where opMaterial = @MaterialPlan and opEnmarcha IN (0,1)  order by opFechaFin desc )					                                
                                        SET @FabricacionPendiente = ISNULL( @CantidadPlanificada,0)- isnull(@CantidadFabricada,0)
	                                end
	                                else
	                                begin
		                                if  DATEADD(day,@diasControl,GETDATE()) >=  @FechaPrevistaFin
		                                begin 
			                                SELECT 
			                                @CantidadPlanificada= SUM(FAB.opCantidadPlanif)  ,
			                                @CantidadFabricada = SUM(FAB.opCantidadFabricadaSAPV2) --opCantidadFabBuenas
			                                FROM Fabricaciones FAB 
			                                WHERE  opMaterial = @MaterialPlan  AND  opEnmarcha IN (0,1) 

                                               set @CantidadFabricada = 0									
                                               set  @opOrdenEnvSAP  = 0;  
                                            set @CantidadFabricadaLOCAL = 0
  
  
                                        DECLARE vendor_cursor CURSOR FOR   
                                         SELECT distinct FAB.opOrdenEnvSAP
                                         FROM Fabricaciones FAB  with (nolock)
                                         WHERE  opMaterial =  @MaterialPlan  AND  opEnmarcha IN (0,1) 
 
  
                                        OPEN vendor_cursor  
  
                                        FETCH NEXT FROM vendor_cursor   
                                        INTO @opOrdenEnvSAP 
  
                                        WHILE @@FETCH_STATUS = 0  
                                        BEGIN  

	                                        set @CantidadFabricadaLOCAL = 0

                                           SELECT @CantidadFabricadaLOCAL = FAB.opCantidadFabricadaSAPV2 
			                                                                        FROM Fabricaciones FAB 
			                                                                        WHERE  opOrdenEnvSAP =  @opOrdenEnvSAP
											                                        group by opOrdenEnvSAP,opCantidadFabricadaSAPV2 
											     
                                             set @CantidadFabricada = @CantidadFabricada + @CantidadFabricadaLOCAL
  
                                                -- Get the next vendor.  
                                            FETCH NEXT FROM vendor_cursor   
                                            INTO @opOrdenEnvSAP
                                        END   
                                        CLOSE vendor_cursor;  
                                        DEALLOCATE vendor_cursor;  



                                            SET @FechaFinFabCurso = (SELECT TOP 1 opFechaFin   FROM Fabricaciones with(nolock) where opMaterial = @MaterialPlan and opEnmarcha IN (0,1)  order by opFechaFin desc )					
			                                SET @FabricacionPendiente =  ISNULL( @CantidadPlanificada,0)- isnull(@CantidadFabricada,0)
		                                end
	                                end     
                                end
    
                                SET @SituacionActual = @StockActual+@FabricacionPendiente
                                SET @UnidadesFabricar= @CantidadForeCast - @SituacionActual


                                select @cantidadFormulaBase=SUM(LM.dlCantidad )
                                from ListaMateriales LM  with(nolock) 
                                LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
                                where MA.maCod=@MaterialPlan and LM.dlUM='KG'

                                --Si la cantidad base es nula se busca en su lista de materiales para obtener esa cantidad base								
								if @cantidadFormulaBase is null 
								begin
								 select @cantidadFormulaBase= SUM(LM.dlCantidad )
																from ListaMateriales LM  with(nolock) 
																where dlUM = 'KG' and dlLista in ( select maListaMaterial
																from ListaMateriales LM  with(nolock) 
																LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
																where MA.maCod in (select dlMaterial
																from ListaMateriales LM  with(nolock) 
																LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
																where MA.maCod=@MaterialPlan))
								end
    
                                select top 1 @cantidadBase  =clCantidad 
                                from ListaMaterialesCab where clLista = (select top 1 maListaMaterial from Materiales  where maCod = @MaterialPlan)

                                if  @UnidadesFabricar >0
	                            begin
	                            set @KgNuevaFabricacion= (@UnidadesFabricar * @cantidadFormulaBase)/ @cantidadBase
	                            end
	                            else
	                            begin
	                            set @KgNuevaFabricacion=0
	                            end

                                SELECT 
                                MA.maCod as Codigo,
                                MA.maNombre as NombreMat ,
                                MA.maGrupoHR as HojaRuraDefault,
                                MA.maContHR  as ContHojaRutaDefault,
                                isnull(F.fmNombre,'')  as Formato ,
                                isnull(F.fmCod ,'')  as GrupoPlanif ,
                                @CantidadForeCast as CantidadFC,
                                isnull(MA.maDiasPP,0)  as DiasPP,
                                MA.maStokMinPS as StockMin ,
                                MA.maStokMaxPS as StockMax,
                                isnull(@StockActual,0) as StockActual,
                                isnull(@FabricacionPendiente,0) AS  FabriacionPendiente,
                                @FechaFinFabCurso as FechaFinFabricacionActual,
                                isnull(@SituacionActual,0) as  SituacionActual,
                                isnull(@UnidadesFabricar,0) as  NuevaFabricacion,
                                isnull(@KgNuevaFabricacion,0) as KgNuevaFabricacion,
                                isnull(@cantidadFormulaBase,0) as CantidadFormulaBase,
                                isnull(@cantidadBase,0) as CantidadBase,
                                (select PT.ptCod  from PuestosTrabajo PT  
                                where ptCod = 
                                (select top 1 opPuestoTrabajo  from OperacionesHojaRuta  
                                 where opGrupoSAP=  MA.maGrupoHR and opContOperSAP=MA.maContHR)) as CodPuestoTrabajo,
                                (select PT.ptNombre   from PuestosTrabajo PT  
                                where ptCod = 
                                (select top 1 opPuestoTrabajo  from OperacionesHojaRuta  
                                 where opGrupoSAP=  MA.maGrupoHR and opContOperSAP=MA.maContHR)) as NombrePuestoTrabajo,
                                (select PT.ptCentroProd    from PuestosTrabajo PT  
                                where ptCod = 
                                (select top 1 opPuestoTrabajo  from OperacionesHojaRuta  
                                 where opGrupoSAP=  MA.maGrupoHR and opContOperSAP=MA.maContHR)) as CodCentroProd,
                                MA.maRedondeo  as RedondeoLote
                                from materiales MA  with(nolock) 
                                LEFT JOIN HojaRuta HR with(nolock)  on MA.maGrupoHR= HR.hrGrupo and MA.maContHR = HR.hrContGrupo 
                                LEFT JOIN Formatos F with(nolock)  on HR.hrFormato = F.fmCod 
                                where MA.maCod =@MaterialPlan"

            Dim DTDatos As New DataTable
            Dim DTDatosPuestos As New DataTable
            InicializarVariables()








            If Datos.CGPL.DameDatosDT(sSql, DTDatos) Then

                'Dim sSqlPuestos As String = " " &
                '                " select PT.ptNombre as PuestosTrabajo , PT.ptCod as Codigo  from PuestosTrabajo PT  
                '                where ptCod in
                '                (select  opPuestoTrabajo  from OperacionesHojaRuta  
                '                 where opGrupoSAP in  (select maGrupoHR
                '                from ListaMateriales LM  with(nolock) 
                '                LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
                '                where MA.maCod='" & UTrim(DTDatos.Rows(0).Item("Codigo")) & "' ) and opContOperSAP in (select maContHR
                '                from ListaMateriales LM  with(nolock) 
                '                LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
                '                where MA.maCod='" & UTrim(DTDatos.Rows(0).Item("Codigo")) & "' ))"

                Dim sSqlPuestos As String = " " &
                                " select PT.ptNombre as PuestosTrabajo , PT.ptCod as Codigo  from PuestosTrabajo PT  
                                where ptCod in
                                (select  opPuestoTrabajo  from OperacionesHojaRuta  
                                 where opGrupoSAP in  (select maGrupoHR
                                from ListaMateriales LM  with(nolock) 
                                LEFT JOIN materiales MA   with(nolock) on MA.maListaMaterial =LM.dlLista
                                where MA.maCod='" & UTrim(DTDatos.Rows(0).Item("Codigo")) & "' ))"

                If UTrim(DTDatos.Rows(0).Item("Codigo")) = "70901974" Then
                    Dim g = 7
                End If

                Dim listaPuestosTrabajo As New List(Of PuestosTrabajo)
                If Datos.CGPL.DameDatosDT(sSqlPuestos, DTDatosPuestos) Then
                    'If DTDatos.Rows.Count > 1 Then
                    For index = 0 To DTDatosPuestos.Rows.Count - 1
                        Dim pue As New PuestosTrabajo
                        pue.Nombre = UTrim(DTDatosPuestos.Rows(index).Item("PuestosTrabajo"))
                        pue.CodigoPuestoTrabajo = CInt(NoNull(DTDatosPuestos.Rows(index).Item("Codigo"), "N"))
                        listaPuestosTrabajo.Add(pue)
                        'listaPuestosTrabajo.Add(UTrim(DTDatosPuestos.Rows(index).Item("PuestosTrabajo")))
                    Next
                    'End If
                End If
                sDetallePullSystem = DetallePullSystem
                sNombresPuestoTrabajo = listaPuestosTrabajo
                sCodigo = UTrim(DTDatos.Rows(0).Item("Codigo"))
                sNombre = UTrim(DTDatos.Rows(0).Item("NombreMat"))
                sFormato = UTrim(DTDatos.Rows(0).Item("Formato"))
                sGrupoPlanif = UTrim(DTDatos.Rows(0).Item("GrupoPlanif"))
                iCantidad = cantidadFC
                iDiasPP = CInt(NoNull(DTDatos.Rows(0).Item("DiasPP"), "N"))
                iStockMinimo = CInt(NoNull(DTDatos.Rows(0).Item("StockMin"), "N"))
                iStockMaximo = CInt(NoNull(DTDatos.Rows(0).Item("StockMax"), "N"))
                iStockActual = stockActual
                iFabricacionesPendientes = CInt(NoNull(DTDatos.Rows(0).Item("FabriacionPendiente"), "N"))
                iSituacionActual = CInt(NoNull(DTDatos.Rows(0).Item("SituacionActual"), "N"))
                iUnidadesFabricar = CInt(NoNull(DTDatos.Rows(0).Item("NuevaFabricacion"), "N"))
                iKgNuevaFabricacion = CInt(NoNull(DTDatos.Rows(0).Item("KgNuevaFabricacion"), "N"))
                dtFechaFin = fechaPrevistaFin
                iValorRedondeo = CInt(NoNull(DTDatos.Rows(0).Item("RedondeoLote"), "N"))
                iMesPS = MesPS
                iAnioPS = AnioPS

                sHojaRuraDefault = UTrim(DTDatos.Rows(0).Item("HojaRuraDefault"))
                sContHojaRutaDefault = UTrim(DTDatos.Rows(0).Item("ContHojaRutaDefault"))
                iCodPuestoTrabajo = CInt(NoNull(DTDatos.Rows(0).Item("CodPuestoTrabajo"), "N"))
                sNombrePuestoTrabajo = UTrim(DTDatos.Rows(0).Item("NombrePuestoTrabajo"))
                iCodCentroProd = CInt(NoNull(DTDatos.Rows(0).Item("CodCentroProd"), "N"))
                iCantidadBaseFormula = CInt(NoNull(DTDatos.Rows(0).Item("CantidadFormulaBase"), "N"))
                iCantidadBaseMaterias = CInt(NoNull(DTDatos.Rows(0).Item("CantidadBase"), "N"))
                iDiasLaborables = DiasLaborables
                dtFechaFinFanPendientes = CDate(NoNull(DTDatos.Rows(0).Item("FechaFinFabricacionActual"), "DT"))
                iFechaRotura = FechaRotura

                iStockBloqueado = StockBloqueado
                iEstatus = Estatus

                Me.bCreado = True
            End If

        Catch ex As Exception
            bCreado = False
            Throw New Exception(ex.Message & " -- " & Me.GetType().Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "()", ex)
        End Try
    End Sub


#End Region

#Region "Propiedades"
    Public Property DetallePullSystem As List(Of BeanPullSystem)
        Get
            'If sDetalleEnvioProveedor Is Nothing Then
            '    sDetalleEnvioProveedor = DatosCentrosExternos.DameDetalleEnvioProveedor(IdEnvio:=iCodigoEnvio)
            'ElseIf sDetalleEnvioProveedor.Count = 0 Then
            '    sDetalleEnvioProveedor = DatosCentrosExternos.DameDetalleEnvioProveedor(IdEnvio:=iCodigoEnvio)
            'End If

            Return sDetallePullSystem
        End Get
        Set(value As List(Of BeanPullSystem))
            sDetallePullSystem = value
        End Set
    End Property
    Public Property Codigo As String
        Get
            Return sCodigo
        End Get
        Set(value As String)
            sCodigo = value
        End Set
    End Property
    Public Property Nombre As String
        Get
            Return sNombre
        End Get
        Set(value As String)
            sNombre = value
        End Set
    End Property
    Public Property NombreV2 As String
        Get
            Return sNombre & " - " & sNecesidad
        End Get
        Set(value As String)
            sNombre = value
        End Set
    End Property
    Public Property Formato As String
        Get
            Return sFormato
        End Get
        Set(value As String)
            sFormato = value
        End Set
    End Property
    Public Property GrupoPlanif As String
        Get
            Return sGrupoPlanif
        End Get
        Set(value As String)
            sGrupoPlanif = value
        End Set
    End Property
    Public Property Cantidad As Integer
        Get
            Return iCantidad
        End Get
        Set(value As Integer)
            iCantidad = CInt(value)
        End Set
    End Property
    Public Property DiasPP As Integer
        Get
            Return iDiasPP
        End Get
        Set(value As Integer)
            iDiasPP = CInt(value)
        End Set
    End Property
    Public ReadOnly Property CalculoStockMinimo As Integer
        Get
            Return iStockMinimo
        End Get

    End Property
    Public ReadOnly Property StockMaxPS As Integer
        Get
            Return iStockMaximo
        End Get

    End Property

    Public Property StockActual As Double
        Get
            Return iStockActual
        End Get
        Set(value As Double)
            iStockActual = CDbl(value)
        End Set
    End Property
    Public Property FabricacionPendiente As Integer
        Get
            Return iFabricacionesPendientes
        End Get
        Set(value As Integer)
            iFabricacionesPendientes = CInt(value)
        End Set
    End Property
    Public Property FechaFinFanPendientes As Date
        Get
            Return dtFechaFinFanPendientes
        End Get
        Set(value As Date)
            dtFechaFinFanPendientes = CDate(value)
        End Set
    End Property
    Public Property SituacionActual As Integer
        Get
            Return iSituacionActual
        End Get
        Set(value As Integer)
            iSituacionActual = CInt(value)
        End Set
    End Property
    Public Property UnidadesFabricar As Integer
        Get
            Return iUnidadesFabricar
        End Get
        Set(value As Integer)
            iUnidadesFabricar = CInt(value)
        End Set
    End Property

    Public Property KgNuevaFabricacion As Integer
        Get
            Return iKgNuevaFabricacion
        End Get
        Set(value As Integer)
            iKgNuevaFabricacion = CInt(value)
        End Set
    End Property
    Public Property ValorRdoTanque As Integer
        Get
            '2024-05-09 
            If iValorRdoTanque <> 0 Then
                Return iValorRdoTanque
            End If
            ' esto es temporal se tiene que rediseñar con la alta de Tanques y ahi debe tener como
            ' propiedad la cacidad del mismo y al centro producutivo al que pertenece
            Dim TanqueRedondeo As Integer = 0
            If sHojaRuraDefault = "" Then
                Return TanqueRedondeo
            End If

            If iKgNuevaFabricacion <= 0 Then
                Return TanqueRedondeo
            End If


            Select Case iCodCentroProd
                Case 1 ' Cosmetica
                    'Cosmética: 200 Kg, 500 Kg, 3.000 Kg y 6.000 Kg.
                    Select Case iKgNuevaFabricacion
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
                    Select Case iKgNuevaFabricacion
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

            Return TanqueRedondeo
        End Get
        Set(value As Integer)
            iValorRdoTanque = CInt(value)
        End Set
    End Property
    Public Property NuevaFabricacion As Double
        Get
            Dim UNFabricacion As Double = 0

            If ValorRdoTanque > 0 And iNuevaFabricacion = 0 Then
                If iCantidadBaseFormula > 0 Then
                    iNuevaFabricacion = CInt((ValorRdoTanque * iCantidadBaseMaterias) / iCantidadBaseFormula)
                    '11.04.2023 mostrar la nueva fabricacion con multiplos para tener sufucientes a cubrir la "Unidades a Fabricar"                    
                    Dim nuevaUNFabricacion = iNuevaFabricacion
                    While nuevaUNFabricacion < UnidadesFabricar
                        nuevaUNFabricacion = nuevaUNFabricacion + iNuevaFabricacion
                    End While
                    iNuevaFabricacion = nuevaUNFabricacion
                End If
            End If

            Return iNuevaFabricacion
        End Get
        Set(value As Double)
            iNuevaFabricacion = CInt(value)
        End Set
    End Property


    Public Property FechaFin As Date
        Get
            If dtFechaFin = FechaGlobal Then
                If iSituacionActual > 0 AndAlso iCantidad > 0 Then
                    dtFechaFin = Now.Date.AddDays(Math.Round((iSituacionActual * iDiasLaborables) / iCantidad, MidpointRounding.AwayFromZero))
                Else
                    dtFechaFin = Now.Date
                End If
            End If

            Return dtFechaFin
        End Get
        Set(value As Date)
            dtFechaFin = value
        End Set
    End Property


    Public Property ValorRedondeo As Integer
        Get
            Return iValorRedondeo
        End Get
        Set(value As Integer)
            iValorRedondeo = CInt(value)
        End Set
    End Property
    Public Property MesPS As Integer
        Get
            Return iMesPS
        End Get
        Set(value As Integer)
            iMesPS = CInt(value)
        End Set
    End Property
    Public Property AnioPS As Integer
        Get
            Return iAnioPS
        End Get
        Set(value As Integer)
            iAnioPS = CInt(value)
        End Set
    End Property

    Public Property FechaRotura As Date
        Get
            Return iFechaRotura
        End Get
        Set(value As Date)
            iFechaRotura = CDate(value)
        End Set
    End Property
    Public Property StockBloqueado As Double
        Get
            Return iStockBloqueado
        End Get
        Set(value As Double)
            iStockBloqueado = CDbl(value)
        End Set
    End Property

    Public Property Estatus As String
        Get
            Return iEstatus
        End Get
        Set(value As String)
            iEstatus = CStr(value)
        End Set
    End Property

    Public ReadOnly Property FechaPS As String
        Get
            Return iMesPS.ToString().PadLeft(2, CChar("0")) & "/" & iAnioPS.ToString()
        End Get
    End Property

    Public ReadOnly Property StockCritico As Boolean
        Get
            Dim CantidadForeCast = 0
            Dim esStockCritico As Boolean = False
            If iCantidad > 0 Then
                CantidadForeCast = CInt(iCantidad / 2)
                If iStockActual < CantidadForeCast Then
                    esStockCritico = True
                Else
                    esStockCritico = False
                End If

            End If
            Return esStockCritico
        End Get
    End Property

    Public Property CodPuestoTrabajo As Integer
        Get
            Return iCodPuestoTrabajo
        End Get
        Set(value As Integer)
            iCodPuestoTrabajo = CInt(value)
        End Set
    End Property
    Public Property NombrePuestoTrabajo As String
        Get
            Return sNombrePuestoTrabajo.Trim()
        End Get
        Set(value As String)
            sNombrePuestoTrabajo = value
        End Set
    End Property

    Public Property PuestoTrabajo As String
        Get
            Return sPuestoTrabajo
        End Get
        Set(value As String)
            sPuestoTrabajo = value
        End Set
    End Property

    Public Property CodigoMaterialFab As String
        Get
            Return sCodigoMaterialFab
        End Get
        Set(value As String)
            sCodigoMaterialFab = value
        End Set
    End Property

    Public ReadOnly Property NombresPuestoTrabajo As List(Of PuestosTrabajo)
        Get
            Return sNombresPuestoTrabajo
        End Get
        'Set(value As List(Of String))
        '    sNombresPuestoTrabajo = value
        'End Set
    End Property

    Public Property CantidadPendientePedVentas As Integer
        Get
            Return iCantidadPendientePedVentas
        End Get
        Set(value As Integer)
            iCantidadPendientePedVentas = value
        End Set
    End Property

    Public Property Creado As Boolean
        Get
            Return bCreado
        End Get
        Set(value As Boolean)
            bCreado = CBool(value)
        End Set
    End Property

    Public Property CantidadBaseFormula As Integer
        Get
            Return iCantidadBaseFormula
        End Get
        Set(value As Integer)
            iCantidadBaseFormula = CInt(value)
        End Set
    End Property

    Public Property CantidadBaseMaterias As Integer
        Get
            Return iCantidadBaseMaterias
        End Get

        Set(value As Integer)
            iCantidadBaseMaterias = CInt(value)
        End Set
    End Property

    Public Property Necesidad As String
        Get
            Return sNecesidad
        End Get
        Set(value As String)
            sNecesidad = value
        End Set
    End Property

    Public Property FechaCorta As Date
        Get
            Return sFechaCorta
        End Get
        Set(value As Date)
            sFechaCorta = CDate(value)
        End Set
    End Property
#End Region

End Class
