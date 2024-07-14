VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAMedidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vMedidorID As String
Private vFechaAlta As Date
Private vClienteId As Long
Private vFechaColocacion As Variant
Private vFechaRetiro As Variant
Private vMotivoRetiro As Integer
Private vEstadoInicio As Long
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyAMedidor

Public ddl As New clsMyDMedidor

Private Sub class_Initialize()

    vMedidorID = ""
    vFechaAlta = Date
    vClienteId = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFechaColocacion = Null
    vFechaRetiro = Null
    vMotivoRetiro = 0
    vEstadoInicio = 0
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get medidorID() As Variant

    medidorID = vMedidorID

End Property

Public Property Let medidorID(ByVal vNewValue As Variant)

On Error Resume Next

    vMedidorID = vNewValue

End Property

Public Property Get fechaAlta() As Variant

    fechaAlta = vFechaAlta

End Property

Public Property Let fechaAlta(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaAlta = vNewValue

End Property

Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

Public Property Get fechaColocacion() As Variant

    fechaColocacion = vFechaColocacion

End Property

Public Property Let fechaColocacion(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaColocacion = vNewValue

End Property

Public Property Get fechaRetiro() As Variant

    fechaRetiro = vFechaRetiro

End Property

Public Property Let fechaRetiro(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaRetiro = vNewValue

End Property

Public Property Get motivoRetiro() As Variant

    motivoRetiro = vMotivoRetiro

End Property

Public Property Let motivoRetiro(ByVal vNewValue As Variant)

On Error Resume Next

    vMotivoRetiro = vNewValue

End Property

Public Property Get estadoInicio() As Variant

    estadoInicio = vEstadoInicio

End Property

Public Property Let estadoInicio(ByVal vNewValue As Variant)

On Error Resume Next

    vEstadoInicio = vNewValue

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

    comboText = vFechaColocacion

End Property

Public Property Get comboKey() As Variant

    comboKey = vClienteId

End Property

Public Property Get listText() As Variant

    listText = vFechaColocacion

End Property

Public Property Get listKey() As Variant

    listKey = vClienteId

End Property

Public Function clone() As clsMyAMedidor
Dim objObjeto As New clsMyAMedidor

    With objObjeto
        .medidorID = vMedidorID
        .fechaAlta = vFechaAlta
        .clienteId = vClienteId
        .fechaColocacion = vFechaColocacion
        .fechaRetiro = vFechaRetiro
        .motivoRetiro = vMotivoRetiro
        .estadoInicio = vEstadoInicio
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAMedidor, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .medidorID = pQuery.Fields(.ddl.medidorID)
        .fechaAlta = pQuery.Fields(.ddl.fechaAlta)
        .clienteId = pQuery.Fields(.ddl.clienteId)
        .fechaColocacion = pQuery.Fields(.ddl.fechaColocacion)
        .fechaRetiro = pQuery.Fields(.ddl.fechaRetiro)
        .motivoRetiro = pQuery.Fields(.ddl.motivoRetiro)
        .estadoInicio = pQuery.Fields(.ddl.estadoInicio)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.medidorID & " = '" & replaceSpecialSymbols(vMedidorID) & "'"
    sql = sql & " AND " & ddl.fechaAlta & " = " & fechaHoraDB(vFechaAlta)

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
    sql = sql & " WHERE " & ddl.medidorID & " = '" & replaceSpecialSymbols(vMedidorID) & "'"
    sql = sql & " ORDER BY " & ddl.fechaAlta & " = " & fechaHoraDB(vFechaAlta)
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLastByClienteID(db As clsDB)
Dim sql As String

Dim clienteId As Long

    clienteId = vClienteId

    clean
    
    vClienteId = clienteId
    
    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & vClienteId
    sql = sql & " ORDER BY " & ddl.fechaColocacion & " DESC"
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

Public Sub findByClienteID(db As clsDB, Optional colocado As Boolean = True)
Dim sql As String

Dim clienteId As Long

    clienteId = vClienteId

    clean

    vClienteId = clienteId

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & vClienteId
    sql = sql & " ORDER BY " & ddl.fechaAlta & " DESC, " & ddl.fechaColocacion & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db
    
    If colocado And vFechaRetiro <= Date Then
        clean
        Exit Sub
    End If
    
    If colocado And Not IsNull(vFechaRetiro) Then
        findColocadoByClienteID db
        
        If vAutoID = 0 Then findSQL sql, db
    End If
    
End Sub

Public Sub findColocadoByClienteID(db As clsDB)
Dim sql As String

Dim clienteId As Long

    clienteId = vClienteId

    clean

    vClienteId = clienteId

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteId & " = " & vClienteId
    sql = sql & " AND " & ddl.fechaRetiro & " IS NULL"
    sql = sql & " ORDER BY " & ddl.fechaAlta & " DESC, " & ddl.fechaColocacion & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByMedidorID(db As clsDB)
Dim sql As String

Dim strMedidorID As String

    strMedidorID = vMedidorID

    clean

    vMedidorID = strMedidorID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.medidorID & " = '" & vMedidorID & "'"
    sql = sql & " ORDER BY " & ddl.fechaAlta & " DESC, " & ddl.autoID & " DESC"
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
    sql = sql & " SET " & ddl.fechaColocacion & " = " & fechaDB(vFechaColocacion)
    sql = sql & ", " & ddl.fechaRetiro & " = " & fechaDB(vFechaRetiro)
    sql = sql & ", " & ddl.motivoRetiro & " = " & vMotivoRetiro
    sql = sql & ", " & ddl.estadoInicio & " = " & vEstadoInicio
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.medidorID & " = '" & replaceSpecialSymbols(vMedidorID) & "'"
    sql = sql & ", " & ddl.fechaAlta & " = " & fechaHoraDB(vFechaAlta)
    sql = sql & ", " & ddl.clienteId & " = " & vClienteId
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

    If vFechaColocacion <> clon.fechaColocacion Or IsNull(vFechaColocacion) <> IsNull(clon.fechaColocacion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaColocacion & " = " & fechaDB(vFechaColocacion)
        blnChange = True
    End If
    If vFechaRetiro <> clon.fechaRetiro Or IsNull(vFechaRetiro) <> IsNull(clon.fechaRetiro) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaRetiro & " = " & fechaDB(vFechaRetiro)
        blnChange = True
    End If
    If vMotivoRetiro <> clon.motivoRetiro Or IsNull(vMotivoRetiro) <> IsNull(clon.motivoRetiro) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.motivoRetiro & " = " & vMotivoRetiro
        blnChange = True
    End If
    If vEstadoInicio <> clon.estadoInicio Or IsNull(vEstadoInicio) <> IsNull(clon.estadoInicio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.estadoInicio & " = " & vEstadoInicio
        blnChange = True
    End If
    If vClienteId <> clon.clienteId Or IsNull(vClienteId) <> IsNull(clon.clienteId) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clienteId & " = " & vClienteId
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

    vClienteId = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.medidorID & " = '" & replaceSpecialSymbols(vMedidorID) & "'"
    sql = sql & " AND " & ddl.fechaAlta & " = " & fechaHoraDB(vFechaAlta)
    sql = sql & " ORDER BY " & ddl.clienteId & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vClienteId = 1 + rstQuery.Fields(ddl.clienteId)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.medidorID & ", " & ddl.fechaAlta & ", " & ddl.autoID
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

Public Function collectionByMedidorID(pMedidorID As String, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.medidorID & " = '" & pMedidorID & "'"
    sql = sql & ";"

    Set collectionByMedidorID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAMedidor

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAMedidor
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
Dim elemento As clsMyAMedidor

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
Dim elemento As clsMyAMedidor

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

