VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyACuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vPlanID As Integer
Private vCuotaID As Integer
Private vFechaVencimiento As Date
Private vFechaPago As Variant
Private vImporte As Currency
Private vCancelada As Integer
Private vPlanIDCancela As Integer
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyACuota

Public ddl As New clsMyDCuota

Private Sub class_Initialize()

    vClienteID = 0
    vPlanID = 0
    vCuotaID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFechaVencimiento = Date
    vFechaPago = Null
    vImporte = 0
    vCancelada = 0
    vPlanIDCancela = 0
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

Public Property Get planID() As Variant

    planID = vPlanID

End Property

Public Property Let planID(ByVal vNewValue As Variant)

On Error Resume Next

    vPlanID = vNewValue

End Property

Public Property Get cuotaID() As Variant

    cuotaID = vCuotaID

End Property

Public Property Let cuotaID(ByVal vNewValue As Variant)

On Error Resume Next

    vCuotaID = vNewValue

End Property

Public Property Get fechaVencimiento() As Variant

    fechaVencimiento = vFechaVencimiento

End Property

Public Property Let fechaVencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaVencimiento = vNewValue

End Property

Public Property Get fechaPago() As Variant

    fechaPago = vFechaPago

End Property

Public Property Let fechaPago(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaPago = vNewValue

End Property

Public Property Get importe() As Variant

    importe = vImporte

End Property

Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

End Property

Public Property Get cancelada() As Variant

    cancelada = vCancelada

End Property

Public Property Let cancelada(ByVal vNewValue As Variant)

On Error Resume Next

    vCancelada = vNewValue

End Property

Public Property Get planIDCancela() As Variant

    planIDCancela = vPlanIDCancela

End Property

Public Property Let planIDCancela(ByVal vNewValue As Variant)

On Error Resume Next

    vPlanIDCancela = vNewValue

End Property

Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

End Property

Public Property Get uid() As Variant

    uid = vUid

End Property

Public Property Let uid(ByVal vNewValue As Variant)

On Error Resume Next

    vUid = vNewValue

End Property

Public Property Get created() As Variant

    created = vCreated

End Property

Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vFechaVencimiento

End Property

Public Property Get comboKey() As Variant

    comboKey = vCuotaID

End Property

Public Property Get listText() As Variant

    listText = vFechaVencimiento

End Property

Public Property Get listKey() As Variant

    listKey = vCuotaID

End Property

Public Function clone() As clsMyACuota
Dim objObjeto As New clsMyACuota

    With objObjeto
        .clienteID = vClienteID
        .planID = vPlanID
        .cuotaID = vCuotaID
        .fechaVencimiento = vFechaVencimiento
        .fechaPago = vFechaPago
        .importe = vImporte
        .cancelada = vCancelada
        .planIDCancela = vPlanIDCancela
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyACuota, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .planID = pQuery.Fields(.ddl.planID)
        .cuotaID = pQuery.Fields(.ddl.cuotaID)
        .fechaVencimiento = pQuery.Fields(.ddl.fechaVencimiento)
        .fechaPago = pQuery.Fields(.ddl.fechaPago)
        .importe = pQuery.Fields(.ddl.importe)
        .cancelada = pQuery.Fields(.ddl.cancelada)
        .planIDCancela = pQuery.Fields(.ddl.planIDCancela)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.planID & " = " & vPlanID
    sql = sql & " AND " & ddl.cuotaID & " = " & vCuotaID

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

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.planID & " = " & vPlanID
    sql = sql & " ORDER BY " & ddl.cuotaID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLastByClienteID(db As clsDB, Optional referencia As Date = #1/1/2100#)
Dim sql As String

Dim clienteID As Long

    clienteID = vClienteID
    
    clean
    
    vClienteID = clienteID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    If Not IsMissing(referencia) Then sql = sql & " AND " & ddl.fechaVencimiento & " <= " & fechaDB(referencia)
    sql = sql & " ORDER BY " & ddl.fechaVencimiento & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLastPagada(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.planID & " = " & vPlanID
    sql = sql & " AND NOT (" & ddl.fechaPago & " IS NULL)"
    sql = sql & " ORDER BY " & ddl.cuotaID & " DESC"
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

Public Sub findPagadoByPlanID(db As clsDB)
Dim sql As String

Dim intClienteID As Integer
Dim intPlanID As Integer

    intClienteID = vClienteID
    intPlanID = vPlanID

    clean

    vClienteID = intClienteID
    vPlanID = intPlanID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.planID & " = " & vPlanID
    sql = sql & " AND NOT (" & ddl.fechaPago & " IS NULL)"
    sql = sql & " ORDER BY " & ddl.cuotaID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findPendienteByPlanID(db As clsDB)
Dim sql As String

Dim intClienteID As Integer
Dim intPlanID As Integer

    intClienteID = vClienteID
    intPlanID = vPlanID

    clean

    vClienteID = intClienteID
    vPlanID = intPlanID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.planID & " = " & vPlanID
    sql = sql & " AND " & ddl.fechaPago & " IS NULL"
    sql = sql & " ORDER BY " & ddl.cuotaID
    sql = sql & " LIMIT 1"
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

    save = save_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function save_(db As clsDB) As Boolean

    If exist_(db) Then
        save_ = update_(db)
    Else
        save_ = add_(db)
    End If

End Function

Public Function exist(db As clsDB) As Boolean

    exist = exist_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function exist_(db As clsDB) As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = db.query(sql)
    exist_ = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(db As clsDB) As Boolean

    add = add_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function add_(db As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.fechaVencimiento & " = " & fechaDB(vFechaVencimiento)
    sql = sql & ", " & ddl.fechaPago & " = " & fechaDB(vFechaPago)
    sql = sql & ", " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
    sql = sql & ", " & ddl.cancelada & " = " & vCancelada
    sql = sql & ", " & ddl.planIDCancela & " = " & vPlanIDCancela
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.planID & " = " & vPlanID
    sql = sql & ", " & ddl.cuotaID & " = " & vCuotaID
    sql = sql & ";"

    add_ = db.execute(sql)

    findByPrimaryKey db

End Function

Public Function update(db As clsDB) As Boolean

    update = update_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function update_(db As clsDB) As Boolean
Dim sql As String
Dim strSET As String

Dim blnChange As Boolean

    update_ = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey db
    End If

    If vFechaVencimiento <> clon.fechaVencimiento Or IsNull(vFechaVencimiento) <> IsNull(clon.fechaVencimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaVencimiento & " = " & fechaDB(vFechaVencimiento)
        blnChange = True
    End If
    If vFechaPago <> clon.fechaPago Or IsNull(vFechaPago) <> IsNull(clon.fechaPago) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaPago & " = " & fechaDB(vFechaPago)
        blnChange = True
    End If
    If vImporte <> clon.importe Or IsNull(vImporte) <> IsNull(clon.importe) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
        blnChange = True
    End If
    If vCancelada <> clon.cancelada Or IsNull(vCancelada) <> IsNull(clon.cancelada) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cancelada & " = " & vCancelada
        blnChange = True
    End If
    If vPlanIDCancela <> clon.planIDCancela Or IsNull(vPlanIDCancela) <> IsNull(clon.planIDCancela) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.planIDCancela & " = " & vPlanIDCancela
        blnChange = True
    End If
    If vUid <> clon.uid Or IsNull(vUid) <> IsNull(clon.uid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
        blnChange = True
    End If

    sql = "UPDATE " & ddl.table & " SET" & strSET
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    If blnChange Then update_ = db.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, db As clsDB)

    newID_ pFind, IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub newID_(pFind As Boolean, db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vCuotaID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.planID & " = " & vPlanID
    sql = sql & " ORDER BY " & ddl.cuotaID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vCuotaID = 1 + rstQuery.Fields(ddl.cuotaID)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionPendienteByPlanID(clienteID As Long, planID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.planID & " = " & planID
    sql = sql & " AND " & ddl.fechaPago & " IS NULL"
    sql = sql & ";"

    Set collectionPendienteByPlanID = collectionSQL(sql, db)

End Function

Public Function collectionByPlanIDCancela(clienteID As Long, planID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.planIDCancela & " = " & planID
    sql = sql & ";"

    Set collectionByPlanIDCancela = collectionSQL(sql, db)

End Function

Public Function collectionDeudaByClienteID(clienteID As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.cancelada & " = 0"
    sql = sql & " AND " & ddl.fechaPago & " IS NULL"
    sql = sql & " ORDER BY " & ddl.cuotaID
    sql = sql & ";"

    Set collectionDeudaByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionDeudaByPlanID(clienteID As Long, planID As Integer, db As clsDB, Optional referencia As Date = #1/1/2100#) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.planID & " = " & planID
    sql = sql & " AND (" & ddl.fechaPago & " IS NULL"
    sql = sql & " OR " & ddl.fechaPago & " > " & fechaDB(referencia)
    sql = sql & ")"
    sql = sql & " ORDER BY " & ddl.cuotaID
    sql = sql & ";"

    Set collectionDeudaByPlanID = collectionSQL(sql, db)

End Function

Public Function collectionByPlanID(clienteID As Long, planID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.planID & " = " & planID
    sql = sql & " ORDER BY " & ddl.cuotaID
    sql = sql & ";"

    Set collectionByPlanID = collectionSQL(sql, db)

End Function

Public Function collectionByPago(pFecha As Date, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.fechaPago_t & " = " & fechaDB(pFecha)
    sql = sql & " AND " & ddl.cancelada_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & c.apellido_t & ", " & c.nombre_t & ", " & ddl.clienteID_t & ", " & ddl.cuotaID_t
    sql = sql & ";"

    Set collectionByPago = collectionSQL(sql, db)

End Function

Public Function collectionPagadasByPeriodo(desde As Date, hasta As Date, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.fechaPago_t & " BETWEEN " & fechaDB(desde) & " AND " & fechaDB(hasta)
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & ddl.clienteID_t & ", " & ddl.fechaPago_t
    sql = sql & ";"

    Set collectionPagadasByPeriodo = collectionSQL(sql, db)

End Function

Public Function collectionPagadoByPlanID(clienteID As Long, planID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.planID & " = " & planID
    sql = sql & " AND NOT (" & ddl.fechaPago & " IS NULL)"
    sql = sql & ";"

    Set collectionPagadoByPlanID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyACuota

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyACuota
        fillObject elemento, rstQuery

        registros.add elemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillCombo_ pCombo, collectionAll(db)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyACuota

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
Dim elemento As clsMyACuota

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

