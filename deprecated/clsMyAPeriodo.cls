VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPeriodoID As Integer
Private vDescripcion As String
Private vFechaInicio As Date
Private vFechaFin As Date
Private vFechaPrimero As Date
Private vFechaSegundo As Date
Private vTasa As Currency
Private vLeyenda As String
Private vLiquidado As Currency
Private vUid As String
Private vCreated As Date

Private clon As clsMyAPeriodo

Public ddl As New clsMyDPeriodo

Private Sub class_Initialize()

    vPeriodoID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vDescripcion = ""
    vFechaInicio = Date
    vFechaFin = Date
    vFechaPrimero = Date
    vFechaSegundo = Date
    vTasa = 0
    vLeyenda = ""
    vLiquidado = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get periodoID() As Variant

    periodoID = vPeriodoID

End Property

Public Property Get key() As String

    key = "k." & periodoID
    
End Property

Public Property Let periodoID(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoID = vNewValue

End Property

Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

Public Property Get fechaInicio() As Variant

    fechaInicio = vFechaInicio

End Property

Public Property Let fechaInicio(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaInicio = vNewValue

End Property

Public Property Get fechaFin() As Variant

    fechaFin = vFechaFin

End Property

Public Property Let fechaFin(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaFin = vNewValue

End Property

Public Property Get fechaPrimero() As Variant

    fechaPrimero = vFechaPrimero

End Property

Public Property Let fechaPrimero(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaPrimero = vNewValue

End Property

Public Property Get fechaSegundo() As Variant

    fechaSegundo = vFechaSegundo

End Property

Public Property Let fechaSegundo(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaSegundo = vNewValue

End Property

Public Property Get tasa() As Variant

    tasa = vTasa

End Property

Public Property Let tasa(ByVal vNewValue As Variant)

On Error Resume Next

    vTasa = vNewValue

End Property

Public Property Get leyenda() As Variant

    leyenda = vLeyenda

End Property

Public Property Let leyenda(ByVal vNewValue As Variant)

On Error Resume Next

    vLeyenda = vNewValue

End Property

Public Property Get liquidado() As Variant

    liquidado = vLiquidado

End Property

Public Property Let liquidado(ByVal vNewValue As Variant)

On Error Resume Next

    vLiquidado = vNewValue

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

    comboText = vDescripcion

End Property

Public Property Get comboKey() As Variant

    comboKey = vPeriodoID

End Property

Public Property Get listText() As Variant

    listText = vDescripcion

End Property

Public Property Get listKey() As Variant

    listKey = vPeriodoID

End Property

Public Function clone() As clsMyAPeriodo
Dim objObjeto As New clsMyAPeriodo

    With objObjeto
        .periodoID = vPeriodoID
        .descripcion = vDescripcion
        .fechaInicio = vFechaInicio
        .fechaFin = vFechaFin
        .fechaPrimero = vFechaPrimero
        .fechaSegundo = vFechaSegundo
        .tasa = vTasa
        .leyenda = vLeyenda
        .liquidado = vLiquidado
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAPeriodo, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .periodoID = pQuery.Fields(.ddl.periodoID)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .fechaInicio = pQuery.Fields(.ddl.fechaInicio)
        .fechaFin = pQuery.Fields(.ddl.fechaFin)
        .fechaPrimero = pQuery.Fields(.ddl.fechaPrimero)
        .fechaSegundo = pQuery.Fields(.ddl.fechaSegundo)
        .tasa = pQuery.Fields(.ddl.tasa)
        .leyenda = pQuery.Fields(.ddl.leyenda)
        .liquidado = pQuery.Fields(.ddl.liquidado)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.periodoID & " = " & vPeriodoID

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

Public Sub findToday(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaInicio & " <= " & fechaDB(Date)
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByFecha(fecha As Date, db As clsDB)
Dim sql As String

    clean

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaInicio & " <= " & fechaDB(fecha)
    sql = sql & " AND " & ddl.fechaFin & " >= " & fechaDB(fecha)
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Function delete(db As clsDB) As Boolean

    delete = delete_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function delete_(db As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete_ = db.execute(sql)

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
    sql = sql & " SET " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
    sql = sql & ", " & ddl.fechaInicio & " = " & fechaDB(vFechaInicio)
    sql = sql & ", " & ddl.fechaFin & " = " & fechaDB(vFechaFin)
    sql = sql & ", " & ddl.fechaPrimero & " = " & fechaDB(vFechaPrimero)
    sql = sql & ", " & ddl.fechaSegundo & " = " & fechaDB(vFechaSegundo)
    sql = sql & ", " & ddl.tasa & " = '" & Replace(vTasa, ",", ".") & "'"
    sql = sql & ", " & ddl.leyenda & " = '" & replaceSpecialSymbols(vLeyenda) & "'"
    sql = sql & ", " & ddl.liquidado & " = '" & Replace(vLiquidado, ",", ".") & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.periodoID & " = " & vPeriodoID
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

    If vDescripcion <> clon.descripcion Or IsNull(vDescripcion) <> IsNull(clon.descripcion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
        blnChange = True
    End If
    If vFechaInicio <> clon.fechaInicio Or IsNull(vFechaInicio) <> IsNull(clon.fechaInicio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaInicio & " = " & fechaDB(vFechaInicio)
        blnChange = True
    End If
    If vFechaFin <> clon.fechaFin Or IsNull(vFechaFin) <> IsNull(clon.fechaFin) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaFin & " = " & fechaDB(vFechaFin)
        blnChange = True
    End If
    If vFechaPrimero <> clon.fechaPrimero Or IsNull(vFechaPrimero) <> IsNull(clon.fechaPrimero) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaPrimero & " = " & fechaDB(vFechaPrimero)
        blnChange = True
    End If
    If vFechaSegundo <> clon.fechaSegundo Or IsNull(vFechaSegundo) <> IsNull(clon.fechaSegundo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaSegundo & " = " & fechaDB(vFechaSegundo)
        blnChange = True
    End If
    If vTasa <> clon.tasa Or IsNull(vTasa) <> IsNull(clon.tasa) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tasa & " = '" & Replace(vTasa, ",", ".") & "'"
        blnChange = True
    End If
    If vLeyenda <> clon.leyenda Or IsNull(vLeyenda) <> IsNull(clon.leyenda) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.leyenda & " = '" & replaceSpecialSymbols(vLeyenda) & "'"
        blnChange = True
    End If
    If vLiquidado <> clon.liquidado Or IsNull(vLiquidado) <> IsNull(clon.liquidado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.liquidado & " = '" & Replace(vLiquidado, ",", ".") & "'"
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

    newID_ pFind, IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub newID_(pFind As Boolean, db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vPeriodoID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vPeriodoID = 1 + rstQuery.Fields(ddl.periodoID)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionRecaudadoByPeriodo(desde As Date, hasta As Date, db As clsDB) As Collection
Dim sql As String

Dim f As New clsMyDFactura

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & f.table
    sql = sql & " ON " & f.periodoID_t & " = " & ddl.periodoID_t
    sql = sql & " WHERE " & f.fechaPago_t & " BETWEEN " & fechaDB(desde) & " AND " & fechaDB(hasta)
    sql = sql & " AND " & f.pagada & " = 1"
    sql = sql & " AND " & f.anulada & " = 0"
    sql = sql & " AND " & f.cancelada & " = 0"
    sql = sql & " ORDER BY " & ddl.periodoID
    sql = sql & ";"

    Set collectionRecaudadoByPeriodo = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAPeriodo

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAPeriodo
        fillObject elemento, rstQuery
        
        registros.add elemento, elemento.key

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillCombo_ pCombo, collectionAll(db), db

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection, db As clsDB)
Dim elemento As clsMyAPeriodo

Dim objPer As New clsMyAPeriodo

On Error Resume Next

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next
    
    If pCombo.ListCount = 0 Then
        pCombo.ListIndex = -1
        Exit Sub
    End If

    objPer.findToday db

    If objPer.periodoID = 0 Then
        pCombo.ListIndex = 0
        Exit Sub
    End If
    
    pCombo.Text = objPer.comboText

End Sub

Public Sub fillList(pList As ListBox, db As clsDB)

    fillList_ pList, collectionAll(db)

End Sub

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyAPeriodo

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

