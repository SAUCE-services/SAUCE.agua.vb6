VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyADeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vPlanID As Integer
Private vDeuda As Currency
Private vCuotas As Integer
Private vCuotasPagadas As Integer
Private vTasa As Currency
Private vPagado As Integer
Private vPeriodo As Integer
Private vCancelada As Integer
Private vPlanIDCancela As Integer
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyADeuda

Public ddl As New clsMyDDeuda

Private Sub class_Initialize()

    vClienteID = 0
    vPlanID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vDeuda = 0
    vCuotas = 0
    vCuotasPagadas = 0
    vTasa = 0
    vPagado = 0
    vPeriodo = 0
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

Public Property Get deuda() As Variant

    deuda = vDeuda

End Property

Public Property Let deuda(ByVal vNewValue As Variant)

On Error Resume Next

    vDeuda = vNewValue

End Property

Public Property Get cuotas() As Variant

    cuotas = vCuotas

End Property

Public Property Let cuotas(ByVal vNewValue As Variant)

On Error Resume Next

    vCuotas = vNewValue

End Property

Public Property Get cuotasPagadas() As Variant

    cuotasPagadas = vCuotasPagadas

End Property

Public Property Let cuotasPagadas(ByVal vNewValue As Variant)

On Error Resume Next

    vCuotasPagadas = vNewValue

End Property

Public Property Get tasa() As Variant

    tasa = vTasa

End Property

Public Property Let tasa(ByVal vNewValue As Variant)

On Error Resume Next

    vTasa = vNewValue

End Property

Public Property Get pagado() As Variant

    pagado = vPagado

End Property

Public Property Let pagado(ByVal vNewValue As Variant)

On Error Resume Next

    vPagado = vNewValue

End Property

Public Property Get periodo() As Variant

    periodo = vPeriodo

End Property

Public Property Let periodo(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodo = vNewValue

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

    comboText = vDeuda

End Property

Public Property Get comboKey() As Variant

    comboKey = vPlanID

End Property

Public Property Get listText() As Variant

    listText = vDeuda

End Property

Public Property Get listKey() As Variant

    listKey = vPlanID

End Property

Public Function clone() As clsMyADeuda
Dim objObjeto As New clsMyADeuda

    With objObjeto
        .clienteID = vClienteID
        .planID = vPlanID
        .deuda = vDeuda
        .cuotas = vCuotas
        .cuotasPagadas = vCuotasPagadas
        .tasa = vTasa
        .pagado = vPagado
        .periodo = vPeriodo
        .cancelada = vCancelada
        .planIDCancela = vPlanIDCancela
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyADeuda, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .planID = pQuery.Fields(.ddl.planID)
        .deuda = pQuery.Fields(.ddl.deuda)
        .cuotas = pQuery.Fields(.ddl.cuotas)
        .cuotasPagadas = pQuery.Fields(.ddl.cuotasPagadas)
        .tasa = pQuery.Fields(.ddl.tasa)
        .pagado = pQuery.Fields(.ddl.pagado)
        .periodo = pQuery.Fields(.ddl.periodo)
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

Public Sub findByClienteID(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.planID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.planID & " DESC"
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
    sql = sql & " SET " & ddl.deuda & " = '" & Replace(vDeuda, ",", ".") & "'"
    sql = sql & ", " & ddl.cuotas & " = " & vCuotas
    sql = sql & ", " & ddl.cuotasPagadas & " = " & vCuotasPagadas
    sql = sql & ", " & ddl.tasa & " = '" & Replace(vTasa, ",", ".") & "'"
    sql = sql & ", " & ddl.pagado & " = " & vPagado
    sql = sql & ", " & ddl.periodo & " = " & vPeriodo
    sql = sql & ", " & ddl.cancelada & " = " & vCancelada
    sql = sql & ", " & ddl.planIDCancela & " = " & vPlanIDCancela
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.planID & " = " & vPlanID
    sql = sql & ";"

    add = db.execute(sql)

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

    If vDeuda <> clon.deuda Or IsNull(vDeuda) <> IsNull(clon.deuda) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.deuda & " = '" & Replace(vDeuda, ",", ".") & "'"
        blnChange = True
    End If
    If vCuotas <> clon.cuotas Or IsNull(vCuotas) <> IsNull(clon.cuotas) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuotas & " = " & vCuotas
        blnChange = True
    End If
    If vCuotasPagadas <> clon.cuotasPagadas Or IsNull(vCuotasPagadas) <> IsNull(clon.cuotasPagadas) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuotasPagadas & " = " & vCuotasPagadas
        blnChange = True
    End If
    If vTasa <> clon.tasa Or IsNull(vTasa) <> IsNull(clon.tasa) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tasa & " = '" & Replace(vTasa, ",", ".") & "'"
        blnChange = True
    End If
    If vPagado <> clon.pagado Or IsNull(vPagado) <> IsNull(clon.pagado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.pagado & " = " & vPagado
        blnChange = True
    End If
    If vPeriodo <> clon.periodo Or IsNull(vPeriodo) <> IsNull(clon.periodo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.periodo & " = " & vPeriodo
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

    If blnChange Then update = db.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vPlanID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.planID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vPlanID = 1 + rstQuery.Fields(ddl.planID)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionByClienteID(clienteID As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " ORDER BY " & ddl.planID
    sql = sql & ";"

    Set collectionByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionByPlanIDCancela(clienteID As Long, planID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.planIDCancela & " = " & planID
    sql = sql & ";"

    Set collectionByPlanIDCancela = collectionSQL(sql, db)

End Function

Public Function collectionPendienteByClienteID(clienteID As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.pagado & " = 0"
    sql = sql & " AND " & ddl.cancelada & " = 0"
    sql = sql & " ORDER BY " & ddl.planID
    sql = sql & ";"

    Set collectionPendienteByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionDeudaByClienteID(clienteID As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.pagado & " = 0"
    sql = sql & " ORDER BY " & ddl.planID & " DESC"
    sql = sql & ";"

    Set collectionDeudaByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionDeudores(db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.pagado_t & " = 0"
    sql = sql & " AND " & ddl.cancelada_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & c.apellido_t & ", " & c.nombre_t
    sql = sql & ";"

    Set collectionDeudores = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyADeuda

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyADeuda
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
Dim elemento As clsMyADeuda

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
Dim elemento As clsMyADeuda

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

