VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAImputado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vSerieID As Integer
Private vNumeroID As Long
Private vTipoID As Integer
Private vClienteID As Long
Private vCompSerieID As Integer
Private vCompNumeroID As Long
Private vFecha As Date
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyAImputado

Public ddl As New clsMyDImputado

Private Sub class_Initialize()

    vSerieID = 0
    vNumeroID = 0
    vTipoID = 0
    vClienteID = 0
    vCompSerieID = 0
    vCompNumeroID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFecha = Date
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get serieID() As Variant

    serieID = vSerieID

End Property

Public Property Let serieID(ByVal vNewValue As Variant)

On Error Resume Next

    vSerieID = vNewValue

End Property

Public Property Get numeroID() As Variant

    numeroID = vNumeroID

End Property

Public Property Let numeroID(ByVal vNewValue As Variant)

On Error Resume Next

    vNumeroID = vNewValue

End Property

Public Property Get tipoID() As Variant

    tipoID = vTipoID

End Property

Public Property Let tipoID(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoID = vNewValue

End Property

Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

Public Property Get compSerieID() As Variant

    compSerieID = vCompSerieID

End Property

Public Property Let compSerieID(ByVal vNewValue As Variant)

On Error Resume Next

    vCompSerieID = vNewValue

End Property

Public Property Get compNumeroID() As Variant

    compNumeroID = vCompNumeroID

End Property

Public Property Let compNumeroID(ByVal vNewValue As Variant)

On Error Resume Next

    vCompNumeroID = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

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

    comboKey = vCompNumeroID

End Property

Public Property Get listText() As Variant

    listText = vFecha

End Property

Public Property Get listKey() As Variant

    listKey = vCompNumeroID

End Property

Public Function clone() As clsMyAImputado
Dim objObjeto As New clsMyAImputado

    With objObjeto
        .serieID = vSerieID
        .numeroID = vNumeroID
        .tipoID = vTipoID
        .clienteID = vClienteID
        .compSerieID = vCompSerieID
        .compNumeroID = vCompNumeroID
        .fecha = vFecha
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAImputado, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .serieID = pQuery.Fields(.ddl.serieID)
        .numeroID = pQuery.Fields(.ddl.numeroID)
        .tipoID = pQuery.Fields(.ddl.tipoID)
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .compSerieID = pQuery.Fields(.ddl.compSerieID)
        .compNumeroID = pQuery.Fields(.ddl.compNumeroID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.serieID & " = " & vSerieID
    sql = sql & " AND " & ddl.numeroID & " = " & vNumeroID
    sql = sql & " AND " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.compSerieID & " = " & vCompSerieID
    sql = sql & " AND " & ddl.compNumeroID & " = " & vCompNumeroID

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
    sql = sql & " WHERE " & ddl.serieID & " = " & vSerieID
    sql = sql & " AND " & ddl.numeroID & " = " & vNumeroID
    sql = sql & " AND " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.compSerieID & " = " & vCompSerieID
    sql = sql & " ORDER BY " & ddl.compNumeroID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByPlanIDClienteID(db As clsDB)
Dim sql As String

Dim intClienteID As Integer
Dim intPlanID As Integer

    intClienteID = vClienteID
    intPlanID = vCompSerieID
    
    clean
    
    vClienteID = intClienteID
    vCompSerieID = intPlanID
    
    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.tipoID & " = 2"
    sql = sql & " AND " & ddl.compSerieID & " = " & vCompSerieID
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByReciboComprobante(db As clsDB)
Dim sql As String

Dim serieID As Integer
Dim tipoID As Integer
Dim comprobanteSerieID As Integer

Dim numeroID As Long
Dim comprobanteNumeroID As Long

    serieID = vSerieID
    tipoID = vTipoID
    comprobanteSerieID = vCompSerieID
    numeroID = vNumeroID
    comprobanteNumeroID = vCompNumeroID

    clean

    vSerieID = serieID
    vTipoID = tipoID
    vCompSerieID = comprobanteSerieID
    vNumeroID = numeroID
    vCompNumeroID = comprobanteNumeroID
    
    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.serieID & " = " & vSerieID
    sql = sql & " AND " & ddl.numeroID & " = " & vNumeroID
    sql = sql & " AND " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.compSerieID & " = " & vCompSerieID
    sql = sql & " AND " & ddl.compNumeroID & " = " & vCompNumeroID
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findImputadoByCuotaID(db As clsDB)
Dim sql As String

Dim intClienteID As Integer
Dim intPlanID As Integer
Dim intCuotaID As Integer

    intClienteID = vClienteID
    intPlanID = vCompSerieID
    intCuotaID = vCompNumeroID
    
    clean
    
    vClienteID = intClienteID
    vCompSerieID = intPlanID
    vCompNumeroID = intCuotaID
    
    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.tipoID & " = 2"
    sql = sql & " AND " & ddl.compSerieID & " = " & vCompSerieID
    sql = sql & " AND " & ddl.compNumeroID & " = " & vCompNumeroID
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByComprobante(db As clsDB)
Dim sql As String

Dim tipoID As Integer
Dim comprobanteSerieID As Integer

Dim comprobanteNumeroID As Long

    tipoID = vTipoID
    comprobanteSerieID = vCompSerieID
    comprobanteNumeroID = vCompNumeroID

    clean

    vTipoID = tipoID
    vCompSerieID = comprobanteSerieID
    vCompNumeroID = comprobanteNumeroID
    
    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.compSerieID & " = " & vCompSerieID
    sql = sql & " AND " & ddl.compNumeroID & " = " & vCompNumeroID
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
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.serieID & " = " & vSerieID
    sql = sql & ", " & ddl.numeroID & " = " & vNumeroID
    sql = sql & ", " & ddl.tipoID & " = " & vTipoID
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.compSerieID & " = " & vCompSerieID
    sql = sql & ", " & ddl.compNumeroID & " = " & vCompNumeroID
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

    vCompNumeroID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.serieID & " = " & vSerieID
    sql = sql & " AND " & ddl.numeroID & " = " & vNumeroID
    sql = sql & " AND " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.compSerieID & " = " & vCompSerieID
    sql = sql & " ORDER BY " & ddl.compNumeroID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vCompNumeroID = 1 + rstQuery.Fields(ddl.compNumeroID)

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

Public Function collectionByRecibo(serieID As Integer, numeroID As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.serieID & " = " & serieID
    sql = sql & " AND " & ddl.numeroID & " = " & numeroID
    sql = sql & ";"

    Set collectionByRecibo = collectionSQL(sql, db)

End Function

Public Function collectionByComprobante(tipoID As Integer, comprobanteSerie As Integer, comprobanteNumero As Long, db As clsDB, Optional clienteID As Integer = 0) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tipoID & " = " & tipoID
    sql = sql & " AND " & ddl.compSerieID & " = " & comprobanteSerie
    sql = sql & " AND " & ddl.compNumeroID & " = " & comprobanteNumero
    If clienteID > 0 Then sql = sql & " AND " & ddl.clienteID & " = " & clienteID
    sql = sql & ";"

    Set collectionByComprobante = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAImputado

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAImputado
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
Dim elemento As clsMyAImputado

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
Dim elemento As clsMyAImputado

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

