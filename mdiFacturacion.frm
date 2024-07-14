VERSION 5.00
Begin VB.MDIForm mdiFacturacion 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Agua"
   ClientHeight    =   7230
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14280
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mFacturacion 
      Caption         =   "Facturacion"
      Begin VB.Menu mFacturar 
         Caption         =   "Facturación Electrónica"
      End
      Begin VB.Menu lin37373 
         Caption         =   "-"
      End
      Begin VB.Menu mFactura 
         Caption         =   "Liquidación Individual"
      End
      Begin VB.Menu mLiqGeneral 
         Caption         =   "Liquidación General"
      End
      Begin VB.Menu mLiquidarZona 
         Caption         =   "Liquidación por Zona"
      End
      Begin VB.Menu ray77374 
         Caption         =   "-"
      End
      Begin VB.Menu mRecibo 
         Caption         =   "Recibos de Pagos Parciales"
      End
      Begin VB.Menu mNCredito 
         Caption         =   "Notas de Crédito"
      End
      Begin VB.Menu ray375 
         Caption         =   "-"
      End
      Begin VB.Menu mIvaVentas 
         Caption         =   "Libro Iva Ventas"
      End
   End
   Begin VB.Menu mPagos 
      Caption         =   "Pagos"
      Begin VB.Menu mPago 
         Caption         =   "Carga de Pago"
      End
      Begin VB.Menu ray7635 
         Caption         =   "-"
      End
      Begin VB.Menu mImportPF 
         Caption         =   "Importar Pago Fácil"
      End
      Begin VB.Menu mImportRP 
         Caption         =   "Importar Rapipago"
      End
      Begin VB.Menu raya00 
         Caption         =   "-"
      End
      Begin VB.Menu mGenerarPMC 
         Caption         =   "Generar Facturas PagoMisCuentas"
      End
   End
   Begin VB.Menu mDeuda 
      Caption         =   "Deuda"
      Begin VB.Menu mResumen 
         Caption         =   "Resumen"
      End
      Begin VB.Menu mPlan 
         Caption         =   "Plan de Cuotas"
      End
      Begin VB.Menu ray734754 
         Caption         =   "-"
      End
      Begin VB.Menu mInteresDet 
         Caption         =   "Detalle de Intereses"
      End
      Begin VB.Menu ray11 
         Caption         =   "-"
      End
      Begin VB.Menu mArchivoDGE 
         Caption         =   "Generar Archivo DGE"
      End
   End
   Begin VB.Menu mClientesM 
      Caption         =   "Clientes"
      Begin VB.Menu mCliente 
         Caption         =   "Datos Personales"
      End
      Begin VB.Menu mAnotador 
         Caption         =   "Anotador"
      End
      Begin VB.Menu raya4242 
         Caption         =   "-"
      End
      Begin VB.Menu mNovedad 
         Caption         =   "Novedades por Período"
      End
      Begin VB.Menu mLecMedidor 
         Caption         =   "Lectura de Medidor"
      End
      Begin VB.Menu ray36564 
         Caption         =   "-"
      End
      Begin VB.Menu mOrden 
         Caption         =   "Mantenimiento de Zonas y Rutas"
      End
      Begin VB.Menu ray737373 
         Caption         =   "-"
      End
      Begin VB.Menu mNotificacion15 
         Caption         =   "Notificación 15 días"
      End
      Begin VB.Menu mNotificacionOC 
         Caption         =   "Notificación Acta de Corte"
      End
      Begin VB.Menu mNotificacionCorte 
         Caption         =   "Notificación Corte Real"
      End
      Begin VB.Menu mSuspFac 
         Caption         =   "Suspensión / Reanudación Facturación"
      End
      Begin VB.Menu mDesconexion 
         Caption         =   "Desconexión / Reconexión Medidor"
      End
      Begin VB.Menu ray2353 
         Caption         =   "-"
      End
      Begin VB.Menu mLibroSocio 
         Caption         =   "Libro de Socios"
      End
   End
   Begin VB.Menu mConsulta 
      Caption         =   "Consulta"
      Begin VB.Menu mRepClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mClientesSM 
         Caption         =   "Clientes con Servicio Medido"
      End
      Begin VB.Menu mClientesCF 
         Caption         =   "Clientes con Cuota Fija"
      End
      Begin VB.Menu mSocios 
         Caption         =   "Socios"
      End
      Begin VB.Menu mSociosSM 
         Caption         =   "Socios con Servicio Medido"
      End
      Begin VB.Menu mSociosCF 
         Caption         =   "Socios con Cuota Fija"
      End
      Begin VB.Menu mFacPendPer 
         Caption         =   "Facturas Pendientes por Período"
      End
      Begin VB.Menu mFacPagPer 
         Caption         =   "Facturas Pagadas por Período"
      End
      Begin VB.Menu mFacAnuPer 
         Caption         =   "Facturas Anuladas por Período"
      End
      Begin VB.Menu mFacCancPer 
         Caption         =   "Facturas Canceladas por Período"
      End
      Begin VB.Menu mFacPeriodo 
         Caption         =   "Facturas por Período"
      End
      Begin VB.Menu mMedRetirados 
         Caption         =   "Medidores Retirados"
      End
      Begin VB.Menu mRecDiaria 
         Caption         =   "Recaudación Diaria"
      End
      Begin VB.Menu mRecPeriodo 
         Caption         =   "Recaudación por Período"
      End
      Begin VB.Menu mPendientes 
         Caption         =   "Facturas Pendientes"
      End
      Begin VB.Menu mDeudores 
         Caption         =   "Deudores en Plan de Pago"
      End
      Begin VB.Menu mDeudorDet 
         Caption         =   "Deudores en Plan de Pago (Detalle)"
      End
      Begin VB.Menu mLecturasZona 
         Caption         =   "Lecturas por Zona y Ruta"
      End
      Begin VB.Menu mRpFacPagFechas 
         Caption         =   "Facturas Pagadas entre Fechas"
      End
      Begin VB.Menu mRubroFact 
         Caption         =   "Rubros Facturados por Período"
      End
      Begin VB.Menu mRubroPaga 
         Caption         =   "Rubros Pagados por Período"
      End
      Begin VB.Menu mCuotPeriodo 
         Caption         =   "Cuotas Pagadas entre Fechas"
      End
      Begin VB.Menu mVolumen 
         Caption         =   "Volumen Facturado por Período"
      End
      Begin VB.Menu mVolFactuZona 
         Caption         =   "Volumen Facturado por Período y Zona"
      End
      Begin VB.Menu mFactSusp 
         Caption         =   "Clientes con Facturación Suspendida"
      End
      Begin VB.Menu mSaldos 
         Caption         =   "Saldos por Conexión"
      End
      Begin VB.Menu mFacPend 
         Caption         =   "Detalle de Facturas Pendientes"
      End
      Begin VB.Menu mFacCliFec 
         Caption         =   "Facturas por Cliente entre Fechas"
      End
   End
   Begin VB.Menu mImpresion 
      Caption         =   "Impresión"
      Begin VB.Menu mImpLiqIndiv 
         Caption         =   "Impresión Liquidación Individual"
      End
      Begin VB.Menu mImpGeneral 
         Caption         =   "Impresión Liquidación General"
      End
      Begin VB.Menu mImpZona 
         Caption         =   "Impresión Liquidación por Zona"
      End
      Begin VB.Menu mImpRuta 
         Caption         =   "Impresión Liquidación por Ruta"
      End
   End
   Begin VB.Menu mMantenimiento 
      Caption         =   "Mantenimiento"
      Begin VB.Menu mMedidor 
         Caption         =   "Medidores"
      End
      Begin VB.Menu mPeriodo 
         Caption         =   "Períodos de Facturación"
      End
      Begin VB.Menu mRubro 
         Caption         =   "Rubros"
      End
      Begin VB.Menu mRango 
         Caption         =   "Rangos de Consumo"
      End
      Begin VB.Menu ray85858 
         Caption         =   "-"
      End
      Begin VB.Menu mServicioDest 
         Caption         =   "Destinos del Servicio"
      End
      Begin VB.Menu mCategoriaSocio 
         Caption         =   "Categorías de Socio"
      End
      Begin VB.Menu mEstado 
         Caption         =   "Estados"
      End
   End
   Begin VB.Menu mInformacion 
      Caption         =   "Información"
      Begin VB.Menu mGeneral 
         Caption         =   "Datos Generales"
      End
      Begin VB.Menu mOperador 
         Caption         =   "Operador"
      End
   End
   Begin VB.Menu mFin 
      Caption         =   "Fin"
   End
End
Attribute VB_Name = "mdiFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mAnotador_Click()

    frmAnotador.Show
    
End Sub

Private Sub mArchivoDGE_Click()

    frmArchivoDGE.Show
    
End Sub

Private Sub mCategoriaSocio_Click()

    frmCategoria.Show
    
End Sub

Private Sub mCliente_Click()

    frmCliente.Show
    
End Sub

Private Sub mClientesCF_Click()

    frmRepClientesCF.Show

End Sub

Private Sub mClientesSM_Click()

    frmRepClientesSM.Show
    
End Sub

Private Sub mCuotPeriodo_Click()

    frmRepCuoPagFec.Show
    
End Sub

Private Sub mDesconexion_Click()

    frmDesconexion.Show
    
End Sub

Private Sub mDeudorDet_Click()

    frmRepDetDeu.Show
    
End Sub

Private Sub mDeudores_Click()

    frmRepDeudores.Show
    
End Sub

Private Sub MDIForm_Load()
Dim operador As New clsMyAOperador

Dim consumoService As New clsCtlConnect

    modConfRegional.ponerConfiguracionRegional
    modProperties.loadProperties
    
    consumoService.configureDB
    
    cntIVA = Array("Responsable Inscripto", "Responsable No Inscripto", "Consumidor Final", "iva Exento", "iva No Reponsable", "Responsable Monotributo")

    operador.findLast dbapp
    
    Me.Caption = "Sistema de Agua - " & operador.razonSocial & " - " & dbapp.ip & " - " & dbapp.database
    
End Sub

Private Sub MDIForm_Terminate()

    dbapp.closeDB

End Sub

Private Sub mEstado_Click()

    frmEstado.Show
    
End Sub

Private Sub mFacAnuPer_Click()

    frmRepFacAnuPer.Show
    
End Sub

Private Sub mFacCancPer_Click()

    frmRepFacCanPer.Show
    
End Sub

Private Sub mFacCliFec_Click()

    frmRepFacCliFec.Show
    
End Sub

Private Sub mFacPagPer_Click()

    frmRepFacPagPer.Show
    
End Sub

Private Sub mFacPend_Click()

    frmRepFacPend.Show
    
End Sub

Private Sub mFacPendPer_Click()

    frmRepFacPenPer.Show
    
End Sub

Private Sub mFacPeriodo_Click()

    frmRepFacPeriodo.Show
    
End Sub

Private Sub mFactSusp_Click()

    frmRepCliSusp.Show
    
End Sub

Private Sub mFactura_Click()

    frmFactura.Show
    
End Sub

Private Sub mFacturar_Click()

    frmFactElect.Show
    
End Sub

Private Sub mFin_Click()

    MDIForm_Terminate

    End
    
End Sub

Private Sub mGeneral_Click()

    frmGeneral.Show
    
End Sub

Private Sub mGenerarPMC_Click()

    frmGenerarPMC.Show
    
End Sub

Private Sub mImpGeneral_Click()

    frmImpGeneral.Show
    
End Sub

Private Sub mImpLiqIndiv_Click()

    frmImpLiqIndivD.Show
    
End Sub

Private Sub mImportPF_Click()

    frmImportPF.Show
    
End Sub

Private Sub mImportRP_Click()

    frmImportRP.Show
    
End Sub

Private Sub mImpRuta_Click()

    frmImpRuta.Show
    
End Sub

Private Sub mImpZona_Click()

    frmImpZona.Show
    
End Sub

Private Sub mInteresDet_Click()

    frmRepInteresDet.Show
    
End Sub

Private Sub mIvaVentas_Click()

    frmIvaVentas.Show
    
End Sub

Private Sub mLecMedidor_Click()

    frmLectura.Show
    
End Sub

Private Sub mLecturasZona_Click()

    frmRepLecturas.Show
    
End Sub

Private Sub mLibroSocio_Click()

    frmLibroSocio.Show
    
End Sub

Private Sub mLiqGeneral_Click()

    frmLiquidar.Show
    
End Sub

Private Sub mLiquidarZona_Click()

    frmLiquidarZona.Show
    
End Sub

Private Sub mMedidor_Click()

    frmMedidor.Show
    
End Sub

Private Sub mMedRetirados_Click()

    frmRepMedRetirados.Show
    
End Sub

Private Sub mNCredito_Click()

    frmNCredito.Show
    
End Sub

Private Sub mNotificacion15_Click()

    frmNotificacion15.Show
    
End Sub

Private Sub mNotificacionCorte_Click()

    frmNotificacionCorte.Show
    
End Sub

Private Sub mNotificacionOC_Click()

    frmNotificacionOC.Show
    
End Sub

Private Sub mNovedad_Click()

    frmNovedad.Show
    
End Sub

Private Sub mOperador_Click()

    frmOperador.Show
    
End Sub

Private Sub mOrden_Click()

    frmOrden.Show
    
End Sub

Private Sub mPago_Click()

    frmPago.Show
    
End Sub

Private Sub mPendientes_Click()
    
    frmRepPendiente.Show
    
End Sub

Private Sub mPeriodo_Click()

    frmPeriodo.Show
    
End Sub

Private Sub mPlan_Click()

    frmPlan.Show
    
End Sub

Private Sub mRango_Click()

    frmRango.Show
    
End Sub

Private Sub mRecDiaria_Click()

    frmRepRecDia.Show
    
End Sub

Private Sub mRecibo_Click()

    frmRecibo.Show
    
End Sub

Private Sub mRecPeriodo_Click()

    frmRepRecPer.Show
    
End Sub

Private Sub mRepClientes_Click()

    frmRepClientes.Show
    
End Sub

Private Sub mResumen_Click()

    frmDeudaD.Show
    
End Sub

Private Sub mRpFacPagFechas_Click()

    frmRepFacPagFec.Show
    
End Sub

Private Sub mRubro_Click()

    frmRubro.Show
    
End Sub

Private Sub mRubroFact_Click()

    frmRepRubFacPer.Show
    
End Sub

Private Sub mRubroPaga_Click()

    frmRepRubPagPer.Show
    
End Sub

Private Sub mSaldos_Click()

    frmRepSaldos.Show
    
End Sub

Private Sub mServicioDest_Click()

    frmDestino.Show
    
End Sub

Private Sub mSocios_Click()
    
    frmRepSocios.Show
    
End Sub

Private Sub mSociosCF_Click()

    frmRepSociosCF.Show
    
End Sub

Private Sub mSociosSM_Click()

    frmRepSociosSM.Show
    
End Sub

Private Sub mSuspFac_Click()

    frmSuspFac.Show
    
End Sub

Private Sub mVolFactuZona_Click()

    frmRepVolFactuZona.Show
    
End Sub

Private Sub mVolumen_Click()

    frmRepVolFactu.Show
    
End Sub
