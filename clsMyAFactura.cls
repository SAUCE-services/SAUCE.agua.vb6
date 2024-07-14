VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPuntoVta As Integer
Private vNroComprob As Long
Private vFecha As Date
Private vClienteId As Long
Private vPeriodoId As Integer
Private vSituacioniva As Integer
Private vTasa As Currency
Private vDescuento As Currency
Private vPagada As Integer
Private vFechapago As Variant
Private vTipoId As Integer
Private vAnulada As Integer
Private vTotal As Currency
Private vInteres As Currency
Private vLetras As String
Private vPuntoVtaInteres As Integer
Private vNroComprobInteres As Long
Private vIvacf As Currency
Private vIvari As Currency
Private vIvarn As Currency
Private vPeriodoIdfin As Integer
Private vCancelada As Integer
Private vPlanIdcancela As Variant
Private vPfcodigo As String
Private vPfbarras As String
Private vAutoID As Long
Private vUId As String
Private vCreated As Date

Private clon As clsMyAFactura

Public ddl As New clsMyDFactura

Private Sub class_Initialize()

    vPuntoVta = 0
    vNroComprob = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFecha = Date
    vClienteId = 0
    vPeriodoId = 0
    vSituacioniva = 0
    vTasa = 0
    vDescuento = 0
    vPagada = 0
    vFechapago = Null
    vTipoId = 0
    vAnulada = 0
    vTotal = 0
    vInteres = 0
    vLetras = ""
    vPuntoVtaInteres = 0
    vNroComprobInteres = 0
    vIvacf = 0
    vIvari = 0
    vIvarn = 0
    vPeriodoIdfin = 0
    vCancelada = 0
    vPlanIdcancela = Null
    vPfcodigo = ""
    vPfbarras = ""
    vAutoID = 0
    vUId = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get puntoVta() As Variant

    puntoVta = vPuntoVta

End Property

Public Property Let puntoVta(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntoVta = vNewValue

End Property

Public Property Get nroComprob() As Variant

    nroComprob = vNroComprob

End Property

Public Property Let nroComprob(ByVal vNewValue As Variant)

On Error Resume Next

    vNroComprob = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

Public Property Get periodoId() As Variant

    periodoId = vPeriodoId

End Property

Public Property Let periodoId(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoId = vNewValue

End Property

Public Property Get situacionIVA() As Variant

    situacionIVA = vSituacioniva

End Property

Public Property Let situacionIVA(ByVal vNewValue As Variant)

On Error Resume Next

    vSituacioniva = vNewValue

End Property

Public Property Get tasa() As Variant

    tasa = vTasa

End Property

Public Property Let tasa(ByVal vNewValue As Variant)

On Error Resume Next

    vTasa = vNewValue

End Property

Public Property Get descuento() As Variant

    descuento = vDescuento

End Property

Public Property Let descuento(ByVal vNewValue As Variant)

On Error Resume Next

    vDescuento = vNewValue

End Property

Public Property Get pagada() As Variant

    pagada = vPagada

End Property

Public Property Let pagada(ByVal vNewValue As Variant)

On Error Resume Next

    vPagada = vNewValue

End Property

Public Property Get fechapago() As Variant

    fechapago = vFechapago

End Property

Public Property Let fechapago(ByVal vNewValue As Variant)

On Error Resume Next

    vFechapago = vNewValue

End Property

Public Property Get tipoId() As Variant

    tipoId = vTipoId

End Property

Public Property Let tipoId(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoId = vNewValue

End Property

Public Property Get anulada() As Variant

    anulada = vAnulada

End Property

Public Property Let anulada(ByVal vNewValue As Variant)

On Error Resume Next

    vAnulada = vNewValue

End Property

Public Property Get total() As Variant

    total = vTotal

End Property

Public Property Let total(ByVal vNewValue As Variant)

On Error Resume Next

    vTotal = vNewValue

End Property

Public Property Get interes() As Variant

    interes = vInteres

End Property

Public Property Let interes(ByVal vNewValue As Variant)

On Error Resume Next

    vInteres = vNewValue

End Property

Public Property Get letras() As Variant

    letras = vLetras

End Property

Public Property Let letras(ByVal vNewValue As Variant)

On Error Resume Next

    vLetras = vNewValue

End Property

Public Property Get puntoVtaInteres() As Variant

    puntoVtaInteres = vPuntoVtaInteres

End Property

Public Property Let puntoVtaInteres(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntoVtaInteres = vNewValue

End Property

Public Property Get nroComprobInteres() As Variant

    nroComprobInteres = vNroComprobInteres

End Property

Public Property Let nroComprobInteres(ByVal vNewValue As Variant)

On Error Resume Next

    vNroComprobInteres = vNewValue

End Property

Public Property Get ivacf() As Variant

    ivacf = vIvacf

End Property

Public Property Let ivacf(ByVal vNewValue As Variant)

On Error Resume Next

    vIvacf = vNewValue

End Property

Public Property Get ivari() As Variant

    ivari = vIvari

End Property

Public Property Let ivari(ByVal vNewValue As Variant)

On Error Resume Next

    vIvari = vNewValue

End Property

Public Property Get ivarn() As Variant

    ivarn = vIvarn

End Property

Public Property Let ivarn(ByVal vNewValue As Variant)

On Error Resume Next

    vIvarn = vNewValue

End Property

Public Property Get periodoIdfin() As Variant

    periodoIdfin = vPeriodoIdfin

End Property

Public Property Let periodoIdfin(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoIdfin = vNewValue

End Property

Public Property Get cancelada() As Variant

    cancelada = vCancelada

End Property

Public Property Let cancelada(ByVal vNewValue As Variant)

On Error Resume Next

    vCancelada = vNewValue

End Property

Public Property Get planIdcancela() As Variant

    planIdcancela = vPlanIdcancela

End Property

Public Property Let planIdcancela(ByVal vNewValue As Variant)

On Error Resume Next

    vPlanIdcancela = vNewValue

End Property

Public Property Get pfcodigo() As Variant

    pfcodigo = vPfcodigo

End Property

Public Property Let pfcodigo(ByVal vNewValue As Variant)

On Error Resume Next

    vPfcodigo = vNewValue

End Property

Public Property Get pfbarras() As Variant

    pfbarras = vPfbarras

End Property

Public Property Let pfbarras(ByVal vNewValue As Variant)

On Error Resume Next

    vPfbarras = vNewValue

End Property

Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

End Property

Public Property Get uid() As Variant

    uid = vUId

End Property

Public Property Let uid(ByVal vNewValue As Variant)

On Error Resume Next

    vUId = vNewValue

End Property

Public Property Get created() As Variant

    created = vCreated

End Property

Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vFecha

End Property

Public Property Get comboKey() As Variant

    comboKey = vNroComprob

End Property

Public Property Get listText() As Variant

    listText = vFecha

End Property

Public Property Get listKey() As Variant

    listKey = vNroComprob

End Property

Public Function clone() As clsMyAFactura
Dim objObjeto As New clsMyAFactura

    With objObjeto
        .puntoVta = vPuntoVta
        .nroComprob = vNroComprob
        .fecha = vFecha
        .clienteId = vClienteId
        .periodoId = vPeriodoId
        .situacionIVA = vSituacioniva
        .tasa = vTasa
        .descuento = vDescuento
        .pagada = vPagada
        .fechapago = vFechapago
        .tipoId = vTipoId
        .anulada = vAnulada
        .total = vTotal
        .interes = vInteres
        .letras = vLetras
        .puntoVtaInteres = vPuntoVtaInteres
        .nroComprobInteres = vNroComprobInteres
        .ivacf = vIvacf
        .ivari = vIvari
        .ivarn = vIvarn
        .periodoIdfin = vPeriodoIdfin
        .cancelada = vCancelada
        .planIdcancela = vPlanIdcancela
        .pfcodigo = vPfcodigo
        .pfbarras = vPfbarras
        .autoID = vAutoID
        .uid = vUId
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAFactura, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .puntoVta = pQuery.Fields(.ddl.puntoVta)
        .nroComprob = pQuery.Fields(.ddl.nroComprob)
        .fecha = pQuery.Fields(.ddl.fecha)
        .clienteId = pQuery.Fields(.ddl.clienteId)
        .periodoId = pQuery.Fields(.ddl.periodoId)
        .situacionIVA = pQuery.Fields(.ddl.situacionIVA)
        .tasa = pQuery.Fields(.ddl.tasa)
        .descuento = pQuery.Fields(.ddl.descuento)
        .pagada = pQuery.Fields(.ddl.pagada)
        .fechapago = pQuery.Fields(.ddl.fechapago)
        .tipoId = pQuery.Fields(.ddl.tipoId)
        .anulada = pQuery.Fields(.ddl.anulada)
        .total = pQuery.Fields(.ddl.total)
        .interes = pQuery.Fields(.ddl.interes)
        .letras = pQuery.Fields(.ddl.letras)
        .puntoVtaInteres = pQuery.Fields(.ddl.puntoVtaInteres)
        .nroComprobInteres = pQuery.Fields(.ddl.nroComprobInteres)
        .ivacf = pQuery.Fields(.ddl.ivacf)
        .ivari = pQuery.Fields(.ddl.ivari)
        .ivarn = pQuery.Fields(.ddl.ivarn)
        .periodoIdfin = pQuery.Fields(.ddl.periodoIdfin)
        .cancelada = pQuery.Fields(.ddl.cancelada)
        .planIdcancela = pQuery.Fields(.ddl.planIdcancela)
        .pfcodigo = pQuery.Fields(.ddl.pfcodigo)
        .pfbarras = pQuery.Fields(.ddl.pfbarras)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & " AND " & ddl.nroComprob & " = " & vNroComprob

    makeWherePrimaryKey = sql

End Function

Private Sub findSQL(pSQL As String, db As clsDB)
Dim rstQuery As ADODB.Recordset

    Set rstQuery = db.query(pSQL)

    If Not rstQuery.EOF Then
        fillObject Me, rstQuery

        Set clon = Me.clone
    End If

    rstQuery.Close

End Sub

Public Sub findByPrimaryKey(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByClientePeriodo(db As clsDB, Optional incluyeAnteriores As Boolean = False)
Dim sql As String

Dim clienteId As Long

Dim periodoId As Integer

    clienteId = vClienteId
    periodoId = vPeriodoId

    clean

    vClienteId = clienteId
    vPeriodoId = periodoId

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & vClienteId
    If incluyeAnteriores Then
        sql = sql & " AND " & ddl.periodoId & " <= " & vPeriodoId
    Else
        sql = sql & " AND " & ddl.periodoId & " = " & vPeriodoId
    End If
    sql = sql & " AND " & ddl.anulada & " = 0"
    sql = sql & " ORDER BY " & ddl.periodoId & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByClientePeriodoPrev(db As clsDB, Optional verificaFin As Boolean = False)
Dim sql As String

Dim clienteId As Long

Dim periodoId As Integer

    clienteId = vClienteId
    periodoId = vPeriodoId

    clean

    vClienteId = clienteId
    vPeriodoId = periodoId

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & vClienteId
    sql = sql & " AND " & ddl.periodoId & " < " & vPeriodoId
    If verificaFin Then sql = sql & " AND " & ddl.periodoIdfin & " >= " & vPeriodoId
    sql = sql & " AND " & ddl.anulada & " = 0"
    sql = sql & " ORDER BY " & ddl.periodoId & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByNroComprob(db As clsDB)
Dim sql As String

Dim nroComprob As Long

    nroComprob = vNroComprob
    
    clean
    
    vNroComprob = nroComprob

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & " ORDER BY " & ddl.puntoVta & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & " ORDER BY " & ddl.nroComprob & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLastLast(db As clsDB)
Dim sql As String

    clean
    
    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.puntoVta & " DESC, " & ddl.nroComprob & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByAutoID(db As clsDB)
Dim sql As String

Dim lngAutoID As Long

    lngAutoID = vAutoID

    clean

    vAutoID = lngAutoID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.autoID & " = " & vAutoID
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByPFCodigo(db As clsDB)
Dim sql As String
Dim strCod As String

    strCod = vPfcodigo

    clean

    vPfcodigo = strCod

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.pfcodigo & " = '" & vPfcodigo & "'"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Function delete(db As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete = db.execute(sql)

End Function

Public Function save(db As clsDB) As Boolean

    If exist(db) Then
        save = update(db)
    Else
        save = add(db)
    End If

End Function

Public Function exist(db As clsDB) As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = db.query(sql)
    exist = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(db As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.fecha & " = " & fechaDB(vFecha)
    sql = sql & ", " & ddl.clienteId & " = " & vClienteId
    sql = sql & ", " & ddl.periodoId & " = " & vPeriodoId
    sql = sql & ", " & ddl.situacionIVA & " = " & vSituacioniva
    sql = sql & ", " & ddl.tasa & " = '" & Replace(vTasa, ",", ".") & "'"
    sql = sql & ", " & ddl.descuento & " = '" & Replace(vDescuento, ",", ".") & "'"
    sql = sql & ", " & ddl.pagada & " = " & vPagada
    sql = sql & ", " & ddl.fechapago & " = " & fechaDB(vFechapago)
    sql = sql & ", " & ddl.tipoId & " = " & vTipoId
    sql = sql & ", " & ddl.anulada & " = " & vAnulada
    sql = sql & ", " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
    sql = sql & ", " & ddl.interes & " = '" & Replace(vInteres, ",", ".") & "'"
    sql = sql & ", " & ddl.letras & " = '" & replaceSpecialSymbols(vLetras) & "'"
    sql = sql & ", " & ddl.puntoVtaInteres & " = " & vPuntoVtaInteres
    sql = sql & ", " & ddl.nroComprobInteres & " = " & vNroComprobInteres
    sql = sql & ", " & ddl.ivacf & " = " & vIvacf
    sql = sql & ", " & ddl.ivari & " = " & vIvari
    sql = sql & ", " & ddl.ivarn & " = " & vIvarn
    sql = sql & ", " & ddl.periodoIdfin & " = " & vPeriodoIdfin
    sql = sql & ", " & ddl.cancelada & " = " & vCancelada
    sql = sql & ", " & ddl.planIdcancela & " = " & validateNullField(vPlanIdcancela)
    sql = sql & ", " & ddl.pfcodigo & " = '" & replaceSpecialSymbols(vPfcodigo) & "'"
    sql = sql & ", " & ddl.pfbarras & " = '" & replaceSpecialSymbols(vPfbarras) & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUId) & "'"
    sql = sql & ", " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & ", " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & ";"

    add = db.execute(sql)
    
    updateI2of5 db

    findByPrimaryKey db

End Function

Public Function update(db As clsDB) As Boolean
Dim sql As String
Dim strSET As String

Dim blnChange As Boolean

    update = True

    strSET = ""
    blnChange = False
    
    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey db
    End If

    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = " & fechaDB(vFecha)
        blnChange = True
    End If
    If vClienteId <> clon.clienteId Or IsNull(vClienteId) <> IsNull(clon.clienteId) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clienteId & " = " & vClienteId
        blnChange = True
    End If
    If vPeriodoId <> clon.periodoId Or IsNull(vPeriodoId) <> IsNull(clon.periodoId) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.periodoId & " = " & vPeriodoId
        blnChange = True
    End If
    If vSituacioniva <> clon.situacionIVA Or IsNull(vSituacioniva) <> IsNull(clon.situacionIVA) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.situacionIVA & " = " & vSituacioniva
        blnChange = True
    End If
    If vTasa <> clon.tasa Or IsNull(vTasa) <> IsNull(clon.tasa) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tasa & " = '" & Replace(vTasa, ",", ".") & "'"
        blnChange = True
    End If
    If vDescuento <> clon.descuento Or IsNull(vDescuento) <> IsNull(clon.descuento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descuento & " = '" & Replace(vDescuento, ",", ".") & "'"
        blnChange = True
    End If
    If vPagada <> clon.pagada Or IsNull(vPagada) <> IsNull(clon.pagada) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.pagada & " = " & vPagada
        blnChange = True
    End If
    If vFechapago <> clon.fechapago Or IsNull(vFechapago) <> IsNull(clon.fechapago) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechapago & " = " & fechaDB(vFechapago)
        blnChange = True
    End If
    If vTipoId <> clon.tipoId Or IsNull(vTipoId) <> IsNull(clon.tipoId) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipoId & " = " & vTipoId
        blnChange = True
    End If
    If vAnulada <> clon.anulada Or IsNull(vAnulada) <> IsNull(clon.anulada) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.anulada & " = " & vAnulada
        blnChange = True
    End If
    If vTotal <> clon.total Or IsNull(vTotal) <> IsNull(clon.total) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
        blnChange = True
    End If
    If vInteres <> clon.interes Or IsNull(vInteres) <> IsNull(clon.interes) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.interes & " = '" & Replace(vInteres, ",", ".") & "'"
        blnChange = True
    End If
    If vLetras <> clon.letras Or IsNull(vLetras) <> IsNull(clon.letras) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.letras & " = '" & replaceSpecialSymbols(vLetras) & "'"
        blnChange = True
    End If
    If vPuntoVtaInteres <> clon.puntoVtaInteres Or IsNull(vPuntoVtaInteres) <> IsNull(clon.puntoVtaInteres) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.puntoVtaInteres & " = " & vPuntoVtaInteres
        blnChange = True
    End If
    If vNroComprobInteres <> clon.nroComprobInteres Or IsNull(vNroComprobInteres) <> IsNull(clon.nroComprobInteres) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nroComprobInteres & " = " & vNroComprobInteres
        blnChange = True
    End If
    If vIvacf <> clon.ivacf Or IsNull(vIvacf) <> IsNull(clon.ivacf) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ivacf & " = " & vIvacf
        blnChange = True
    End If
    If vIvari <> clon.ivari Or IsNull(vIvari) <> IsNull(clon.ivari) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ivari & " = " & vIvari
        blnChange = True
    End If
    If vIvarn <> clon.ivarn Or IsNull(vIvarn) <> IsNull(clon.ivarn) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ivarn & " = " & vIvarn
        blnChange = True
    End If
    If vPeriodoIdfin <> clon.periodoIdfin Or IsNull(vPeriodoIdfin) <> IsNull(clon.periodoIdfin) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.periodoIdfin & " = " & vPeriodoIdfin
        blnChange = True
    End If
    If vCancelada <> clon.cancelada Or IsNull(vCancelada) <> IsNull(clon.cancelada) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cancelada & " = " & vCancelada
        blnChange = True
    End If
    If vPlanIdcancela <> clon.planIdcancela Or IsNull(vPlanIdcancela) <> IsNull(clon.planIdcancela) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.planIdcancela & " = " & validateNullField(vPlanIdcancela)
        blnChange = True
    End If
    If vPfcodigo <> clon.pfcodigo Or IsNull(vPfcodigo) <> IsNull(clon.pfcodigo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.pfcodigo & " = '" & replaceSpecialSymbols(vPfcodigo) & "'"
        blnChange = True
    End If
    If vPfbarras <> clon.pfbarras Or IsNull(vPfbarras) <> IsNull(clon.pfbarras) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.pfbarras & " = '" & replaceSpecialSymbols(vPfbarras) & "'"
        blnChange = True
    End If
    If vUId <> clon.uid Or IsNull(vUId) <> IsNull(clon.uid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.uid & " = '" & replaceSpecialSymbols(vUId) & "'"
        blnChange = True
    End If

    sql = "UPDATE " & ddl.table & " SET" & strSET
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    If blnChange Then
        update = db.execute(sql)
        updateI2of5 db
    End If

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vNroComprob = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & " ORDER BY " & ddl.nroComprob & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vNroComprob = 1 + rstQuery.Fields(ddl.nroComprob)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionAny(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set collectionAny = collectionSQL(sql, db)

End Function

Public Function collectionByPlanIDCancela(clienteId As Long, planID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & clienteId
    sql = sql & " AND " & ddl.planIdcancela & " = " & planID
    sql = sql & ";"

    Set collectionByPlanIDCancela = collectionSQL(sql, db)

End Function

Public Function collectionByPeriodoID(periodoId As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.periodoId & " = " & periodoId
    sql = sql & " AND " & ddl.pagada & " = 0"
    sql = sql & " AND " & ddl.cancelada & " = 0"
    sql = sql & ";"

    Set collectionByPeriodoID = collectionSQL(sql, db)

End Function

Public Function collectionByPeriodoId_(periodoId As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.periodoId & " = " & periodoId
    sql = sql & " AND " & ddl.anulada & " = 0"
    sql = sql & ";"

    Set collectionByPeriodoId_ = collectionSQL(sql, db)

End Function

Public Function collectionByPeriodoPago(desde As Date, hasta As Date, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.fechaPago_t & " >= " & fechaDB(desde)
    sql = sql & " AND " & ddl.fechaPago_t & " <= " & fechaDB(hasta)
    sql = sql & " AND " & ddl.pagada_t & " <> 0"
    sql = sql & " AND " & ddl.anulada_t & " = 0"
    sql = sql & " AND " & ddl.cancelada_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & c.apellido_t & ", " & c.nombre_t & ", " & ddl.clienteID_t & ", " & ddl.periodoID_t
    sql = sql & ";"

    Set collectionByPeriodoPago = collectionSQL(sql, db)

End Function

Public Function collectionByClienteId(clienteId As Long, db As clsDB, Optional invertido As Boolean = False) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & clienteId
    sql = sql & " AND " & ddl.anulada & " = 0"
    sql = sql & " AND " & ddl.cancelada & " = 0"
    If invertido Then
        sql = sql & " ORDER BY " & ddl.puntoVta & " DESC, " & ddl.nroComprob & " DESC"
    Else
        sql = sql & " ORDER BY " & ddl.puntoVta & ", " & ddl.nroComprob
    End If
    sql = sql & ";"

    Set collectionByClienteId = collectionSQL(sql, db)

End Function

Public Function collectionInteresByClienteID(clienteId As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & clienteId
    sql = sql & " AND (" & ddl.puntoVtaInteres & ", " & ddl.nroComprobInteres & ")"
    sql = sql & " IN (SELECT " & ddl.puntoVtaInteres & ", " & ddl.nroComprobInteres & " FROM factura"
    sql = sql & " WHERE " & ddl.clienteId & " = " & clienteId
    sql = sql & " AND " & ddl.nroComprobInteres & " > 0)"
    sql = sql & " ORDER BY " & ddl.puntoVta & ", " & ddl.nroComprob
    sql = sql & ";"

    Set collectionInteresByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionDeudaByClienteId(clienteId As Long, db As clsDB, Optional periodoId As Integer = 0) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & clienteId
    If periodoId > 0 Then sql = sql & " AND " & ddl.periodoId & " < " & periodoId
    sql = sql & " AND " & ddl.pagada & " = 0"
    sql = sql & " AND " & ddl.anulada & " = 0"
    sql = sql & " AND " & ddl.cancelada & " = 0"
    sql = sql & " ORDER BY " & ddl.periodoId & " DESC"
    sql = sql & ";"

    Set collectionDeudaByClienteId = collectionSQL(sql, db)

End Function

Public Function collectionDeudaDiferidaByClienteID(clienteId As Long, fecha As Date, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & clienteId
    sql = sql & " AND " & ddl.anulada & " = 0"
    sql = sql & " AND " & ddl.cancelada & " = 0"
    sql = sql & " AND " & ddl.fecha & " < " & fechaDB(fecha)
    sql = sql & " AND (" & ddl.pagada & " = 0"
    sql = sql & " OR " & ddl.fechapago & " > " & fechaDB(fecha) & ")"
    sql = sql & " ORDER BY " & ddl.periodoId & " DESC"
    sql = sql & ";"

    Set collectionDeudaDiferidaByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionInteresByLiquidacion(puntoVta As Integer, nroComprob As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVtaInteres & " = " & puntoVta
    sql = sql & " AND " & ddl.nroComprobInteres & " = " & nroComprob
    sql = sql & ";"

    Set collectionInteresByLiquidacion = collectionSQL(sql, db)

End Function

Public Function collectionParaInteresByClienteID(clienteId As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVtaInteres & " = 0"
    sql = sql & " AND " & ddl.nroComprobInteres & " = 0"
    sql = sql & " AND " & ddl.clienteId & " = " & clienteId
    sql = sql & " AND " & ddl.anulada & " = 0"
    sql = sql & " AND " & ddl.cancelada & " = 0"
    sql = sql & " AND NOT " & ddl.fechapago & " IS NULL"
    sql = sql & ";"

    Set collectionParaInteresByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionPendientesByPeriodoID(periodoId As Integer, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.periodoID_t & " = " & periodoId
    sql = sql & " AND " & ddl.pagada_t & " = 0"
    sql = sql & " AND " & ddl.anulada_t & " = 0"
    sql = sql & " AND " & ddl.cancelada_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & c.apellido_t & ", " & c.nombre_t & ", " & c.clienteID_t
    sql = sql & ";"

    Set collectionPendientesByPeriodoID = collectionSQL(sql, db)
            
End Function

Public Function collectionPagadasByPeriodoID(periodoId As Integer, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.periodoID_t & " = " & periodoId
    sql = sql & " AND " & ddl.pagada_t & " <> 0"
    sql = sql & " AND " & ddl.anulada_t & " = 0"
    sql = sql & " AND " & ddl.cancelada_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & c.apellido_t & ", " & c.nombre_t & ", " & c.clienteID_t
    sql = sql & ";"

    Set collectionPagadasByPeriodoID = collectionSQL(sql, db)
            
End Function

Public Function collectionAnuladasByPeriodoID(periodoId As Integer, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.periodoID_t & " = " & periodoId
    sql = sql & " AND " & ddl.pagada_t & " = 0"
    sql = sql & " AND " & ddl.anulada_t & " <> 0"
    sql = sql & " AND " & ddl.cancelada_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & c.apellido_t & ", " & c.nombre_t & ", " & c.clienteID_t
    sql = sql & ";"

    Set collectionAnuladasByPeriodoID = collectionSQL(sql, db)
            
End Function

Public Function collectionCanceladasByPeriodoID(periodoId As Integer, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.periodoID_t & " = " & periodoId
    sql = sql & " AND " & ddl.pagada_t & " = 0"
    sql = sql & " AND " & ddl.anulada_t & " = 0"
    sql = sql & " AND " & ddl.cancelada_t & " <> 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & c.apellido_t & ", " & c.nombre_t & ", " & c.clienteID_t
    sql = sql & ";"

    Set collectionCanceladasByPeriodoID = collectionSQL(sql, db)
            
End Function

Public Function collectionPagadasByPeriodo(desde As Date, hasta As Date, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.fechaPago_t & " BETWEEN " & fechaDB(desde) & " AND " & fechaDB(hasta)
    sql = sql & " AND " & ddl.pagada_t & " <> 0"
    sql = sql & " AND " & ddl.anulada_t & " = 0"
    sql = sql & " AND " & ddl.cancelada_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & ddl.fechaPago_t & ", " & ddl.puntoVta_t & ", " & ddl.nroComprob_t
    sql = sql & ";"

    Set collectionPagadasByPeriodo = collectionSQL(sql, db)

End Function

Public Function collectionByPago(fecha As Date, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.fechaPago_t & " = " & fechaDB(fecha)
    sql = sql & " AND " & ddl.pagada_t & " <> 0"
    sql = sql & " AND " & ddl.anulada_t & " = 0"
    sql = sql & " AND " & ddl.cancelada_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & ddl.fechaPago_t & ", " & ddl.clienteID_t
    sql = sql & ";"

    Set collectionByPago = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAFactura

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAFactura
        fillObject elemento, rstQuery

        registros.add elemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillCombo_ pCombo, collectionAll(db)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyAFactura

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, db As clsDB)

    fillList_ pList, collectionAll(db)

End Sub

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyAFactura

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

Private Sub updateI2of5(db As clsDB)
Dim sql As String

Dim objCommand As New ADODB.Command

Dim pagofacil_service As New clsCtlPagoFacil

    vPfbarras = pagofacil_service.codigoI2of5(vPfcodigo)
    
    sql = "UPDATE " & ddl.table
    sql = sql & " SET " & ddl.pfbarras & " = ?"
    sql = sql & " WHERE " & ddl.puntoVta & " = ?"
    sql = sql & " AND " & ddl.nroComprob & " = ?"
    sql = sql & ";"
    
    With objCommand
        .ActiveConnection = db.activa
        .CommandType = adCmdText
        .CommandText = sql
        .Prepared = True
        
        .parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, vPfbarras)
        .parameters.Append .CreateParameter(, adSmallInt, adParamInput, , vPuntoVta)
        .parameters.Append .CreateParameter(, adInteger, adParamInput, , vNroComprob)
        
        .execute , , adExecuteNoRecords
    End With

End Sub


