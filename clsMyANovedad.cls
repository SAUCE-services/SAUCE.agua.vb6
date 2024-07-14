VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyANovedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vPeriodoID As Integer
Private vRubroID As Integer
Private vFecha As Date
Private vPorcentaje As Currency
Private vCantidad As Currency
Private vImporte As Currency
Private vVeces As Integer
Private vVecesCobradas As Integer
Private vIndefinida As Integer
Private vPeriodoIdSuspension As Variant
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyANovedad

Public ddl As New clsMyDNovedad

Private Sub class_Initialize()

    vClienteID = 0
    vPeriodoID = 0
    vRubroID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFecha = Date
    vPorcentaje = 0
    vCantidad = 0
    vImporte = 0
    vVeces = 0
    vVecesCobradas = 0
    vIndefinida = 0
    vPeriodoIdSuspension = Null
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

Public Property Get periodoID() As Variant

    periodoID = vPeriodoID

End Property

Public Property Let periodoID(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoID = vNewValue

End Property

Public Property Get rubroID() As Variant

    rubroID = vRubroID

End Property

Public Property Let rubroID(ByVal vNewValue As Variant)

On Error Resume Next

    vRubroID = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get porcentaje() As Variant

    porcentaje = vPorcentaje

End Property

Public Property Let porcentaje(ByVal vNewValue As Variant)

On Error Resume Next

    vPorcentaje = vNewValue

End Property

Public Property Get cantidad() As Variant

    cantidad = vCantidad

End Property

Public Property Let cantidad(ByVal vNewValue As Variant)

On Error Resume Next

    vCantidad = vNewValue

End Property

Public Property Get importe() As Variant

    importe = vImporte

End Property

Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

End Property

Public Property Get veces() As Variant

    veces = vVeces

End Property

Public Property Let veces(ByVal vNewValue As Variant)

On Error Resume Next

    vVeces = vNewValue

End Property

Public Property Get vecesCobradas() As Variant

    vecesCobradas = vVecesCobradas

End Property

Public Property Let vecesCobradas(ByVal vNewValue As Variant)

On Error Resume Next

    vVecesCobradas = vNewValue

End Property

Public Property Get indefinida() As Variant

    indefinida = vIndefinida

End Property

Public Property Let indefinida(ByVal vNewValue As Variant)

On Error Resume Next

    vIndefinida = vNewValue

End Property

Public Property Get periodoIdSuspension() As Variant

    periodoIdSuspension = vPeriodoIdSuspension

End Property

Public Property Let periodoIdSuspension(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoIdSuspension = vNewValue

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

    comboText = vFecha

End Property

Public Property Get comboKey() As Variant

    comboKey = vRubroID

End Property

Public Property Get listText() As Variant

    listText = vFecha

End Property

Public Property Get listKey() As Variant

    listKey = vRubroID

End Property

Public Function clone() As clsMyANovedad
Dim objObjeto As New clsMyANovedad

    With objObjeto
        .clienteID = vClienteID
        .periodoID = vPeriodoID
        .rubroID = vRubroID
        .fecha = vFecha
        .porcentaje = vPorcentaje
        .cantidad = vCantidad
        .importe = vImporte
        .veces = vVeces
        .vecesCobradas = vVecesCobradas
        .indefinida = vIndefinida
        .periodoIdSuspension = vPeriodoIdSuspension
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyANovedad, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .periodoID = pQuery.Fields(.ddl.periodoID)
        .rubroID = pQuery.Fields(.ddl.rubroID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .porcentaje = pQuery.Fields(.ddl.porcentaje)
        .cantidad = pQuery.Fields(.ddl.cantidad)
        .importe = pQuery.Fields(.ddl.importe)
        .veces = pQuery.Fields(.ddl.veces)
        .vecesCobradas = pQuery.Fields(.ddl.vecesCobradas)
        .indefinida = pQuery.Fields(.ddl.indefinida)
        .periodoIdSuspension = pQuery.Fields(.ddl.periodoIdSuspension)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.periodoID & " = " & vPeriodoID
    sql = sql & " AND " & ddl.rubroID & " = " & vRubroID

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

Public Sub findVeces(db As clsDB, Optional soloPendientes As Boolean = True)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.rubroID & " = " & vRubroID
    sql = sql & " AND " & ddl.periodoID & " <= " & vPeriodoID
    sql = sql & " AND " & ddl.veces & " > 1"
    If soloPendientes Then sql = sql & " AND " & ddl.vecesCobradas & " < " & ddl.veces
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findRango(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.rubroID & " = " & vRubroID
    sql = sql & " AND (" & ddl.periodoID & " = " & vPeriodoID
    sql = sql & " OR (" & ddl.indefinida & " = 1"
    sql = sql & " AND " & ddl.periodoID & " < " & vPeriodoID
    sql = sql & " AND (" & ddl.periodoIdSuspension & " IS NULL"
    sql = sql & " OR " & ddl.periodoIdSuspension & " >= " & vPeriodoID & ")))"
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.periodoID & " = " & vPeriodoID
    sql = sql & " ORDER BY " & ddl.rubroID & " DESC"
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
    sql = sql & " SET " & ddl.fecha & " = " & fechaDB(vFecha)
    sql = sql & ", " & ddl.porcentaje & " = '" & Replace(vPorcentaje, ",", ".") & "'"
    sql = sql & ", " & ddl.cantidad & " = '" & Replace(vCantidad, ",", ".") & "'"
    sql = sql & ", " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
    sql = sql & ", " & ddl.veces & " = " & vVeces
    sql = sql & ", " & ddl.vecesCobradas & " = " & vVecesCobradas
    sql = sql & ", " & ddl.indefinida & " = " & vIndefinida
    sql = sql & ", " & ddl.periodoIdSuspension & " = " & validateNullField(vPeriodoIdSuspension)
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.periodoID & " = " & vPeriodoID
    sql = sql & ", " & ddl.rubroID & " = " & vRubroID
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

    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = " & fechaDB(vFecha)
        blnChange = True
    End If
    If vPorcentaje <> clon.porcentaje Or IsNull(vPorcentaje) <> IsNull(clon.porcentaje) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.porcentaje & " = '" & Replace(vPorcentaje, ",", ".") & "'"
        blnChange = True
    End If
    If vCantidad <> clon.cantidad Or IsNull(vCantidad) <> IsNull(clon.cantidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cantidad & " = '" & Replace(vCantidad, ",", ".") & "'"
        blnChange = True
    End If
    If vImporte <> clon.importe Or IsNull(vImporte) <> IsNull(clon.importe) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
        blnChange = True
    End If
    If vVeces <> clon.veces Or IsNull(vVeces) <> IsNull(clon.veces) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.veces & " = " & vVeces
        blnChange = True
    End If
    If vVecesCobradas <> clon.vecesCobradas Or IsNull(vVecesCobradas) <> IsNull(clon.vecesCobradas) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.vecesCobradas & " = " & vVecesCobradas
        blnChange = True
    End If
    If vIndefinida <> clon.indefinida Or IsNull(vIndefinida) <> IsNull(clon.indefinida) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.indefinida & " = " & vIndefinida
        blnChange = True
    End If
    If vPeriodoIdSuspension <> clon.periodoIdSuspension Or IsNull(vPeriodoIdSuspension) <> IsNull(clon.periodoIdSuspension) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.periodoIdSuspension & " = " & validateNullField(vPeriodoIdSuspension)
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

    vRubroID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.periodoID & " = " & vPeriodoID
    sql = sql & " ORDER BY " & ddl.rubroID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vRubroID = 1 + rstQuery.Fields(ddl.rubroID)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionByPeriodoID(periodoID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.periodoID & " = " & periodoID
    sql = sql & ";"

    Set collectionByPeriodoID = collectionSQL(sql, db)

End Function

Public Function collectionByRubroID(rubroID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.rubroID & " = " & rubroID
    sql = sql & ";"

    Set collectionByRubroID = collectionSQL(sql, db)

End Function

Public Function collectionByClienteID(clienteID As Long, db As clsDB, Optional periodoID As Integer = 0) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    If periodoID > 0 Then sql = sql & " AND " & ddl.periodoID & " = " & periodoID
    sql = sql & " ORDER BY " & ddl.rubroID
    sql = sql & ";"

    Set collectionByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionUnicasByClienteID(clienteID As Long, periodoID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.periodoID & " = " & periodoID
    sql = sql & " AND " & ddl.veces & " = 1"
    sql = sql & " AND " & ddl.indefinida & " = 0"
    sql = sql & " ORDER BY " & ddl.rubroID
    sql = sql & ";"

    Set collectionUnicasByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionVecesByClienteID(clienteID As Long, periodoID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.periodoID & " <= " & periodoID
    sql = sql & " AND " & ddl.veces & " > 1"
    sql = sql & " AND " & ddl.vecesCobradas & " < " & ddl.veces
    sql = sql & " ORDER BY " & ddl.rubroID
    sql = sql & ";"

    Set collectionVecesByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionIndefinidasByClienteID(clienteID As Long, periodoID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.periodoID & " <= " & periodoID
    sql = sql & " AND (" & ddl.periodoIdSuspension & " IS NULL"
    sql = sql & " OR " & ddl.periodoIdSuspension & " >= " & vPeriodoID & ")"
    sql = sql & " AND " & ddl.indefinida & " = 1"
    sql = sql & " ORDER BY " & ddl.rubroID
    sql = sql & ";"

    Set collectionIndefinidasByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyANovedad

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyANovedad
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
Dim elemento As clsMyANovedad

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
Dim elemento As clsMyANovedad

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

