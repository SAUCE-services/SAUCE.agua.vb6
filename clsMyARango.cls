VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyARango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vCategoria As Integer
Private vRangoID As Integer
Private vFecha As Date
Private vLimiteInferior As Currency
Private vLimiteSuperior As Currency
Private vTarifa As Currency
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyARango

Public ddl As New clsMyDRango

Private Sub class_Initialize()

    vCategoria = 0
    vRangoID = 0
    vFecha = Date

    initializeOther

End Sub

Private Sub initializeOther()

    vLimiteInferior = 0
    vLimiteSuperior = 0
    vTarifa = 0
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get categoria() As Variant

    categoria = vCategoria

End Property

Public Property Let categoria(ByVal vNewValue As Variant)

On Error Resume Next

    vCategoria = vNewValue

End Property

Public Property Get rangoID() As Variant

    rangoID = vRangoID

End Property

Public Property Let rangoID(ByVal vNewValue As Variant)

On Error Resume Next

    vRangoID = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get limiteInferior() As Variant

    limiteInferior = vLimiteInferior

End Property

Public Property Let limiteInferior(ByVal vNewValue As Variant)

On Error Resume Next

    vLimiteInferior = vNewValue

End Property

Public Property Get limiteSuperior() As Variant

    limiteSuperior = vLimiteSuperior

End Property

Public Property Let limiteSuperior(ByVal vNewValue As Variant)

On Error Resume Next

    vLimiteSuperior = vNewValue

End Property

Public Property Get tarifa() As Variant

    tarifa = vTarifa

End Property

Public Property Let tarifa(ByVal vNewValue As Variant)

On Error Resume Next

    vTarifa = vNewValue

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

    comboText = vLimiteInferior

End Property

Public Property Get comboKey() As Variant

    comboKey = vFecha

End Property

Public Property Get listText() As Variant

    listText = vLimiteInferior

End Property

Public Property Get listKey() As Variant

    listKey = vFecha

End Property

Public Function clone() As clsMyARango
Dim objObjeto As New clsMyARango

    With objObjeto
        .categoria = vCategoria
        .rangoID = vRangoID
        .fecha = vFecha
        .limiteInferior = vLimiteInferior
        .limiteSuperior = vLimiteSuperior
        .tarifa = vTarifa
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyARango, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .categoria = pQuery.Fields(.ddl.categoria)
        .rangoID = pQuery.Fields(.ddl.rangoID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .limiteInferior = pQuery.Fields(.ddl.limiteInferior)
        .limiteSuperior = pQuery.Fields(.ddl.limiteSuperior)
        .tarifa = pQuery.Fields(.ddl.tarifa)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.categoria & " = " & vCategoria
    sql = sql & " AND " & ddl.rangoID & " = " & vRangoID
    sql = sql & " AND " & ddl.fecha & " = " & fechaDB(vFecha)

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
    sql = sql & " WHERE " & ddl.categoria & " = " & vCategoria
    sql = sql & " AND " & ddl.rangoID & " = " & vRangoID
    sql = sql & " ORDER BY " & ddl.fecha & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByFechaSinCategoria(fecha As Date, db As clsDB)
Dim sql As String

Dim rangoID As Integer

    rangoID = vRangoID
    
    clean
    
    vRangoID = rangoID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.rangoID & " = " & vRangoID
    sql = sql & " AND " & ddl.fecha & " >= " & fechaDB(fecha)
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
    sql = sql & " SET " & ddl.limiteInferior & " = '" & Replace(vLimiteInferior, ",", ".") & "'"
    sql = sql & ", " & ddl.limiteSuperior & " = '" & Replace(vLimiteSuperior, ",", ".") & "'"
    sql = sql & ", " & ddl.tarifa & " = '" & Replace(vTarifa, ",", ".") & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.categoria & " = " & vCategoria
    sql = sql & ", " & ddl.rangoID & " = " & vRangoID
    sql = sql & ", " & ddl.fecha & " = " & fechaDB(vFecha)
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

    If vLimiteInferior <> clon.limiteInferior Or IsNull(vLimiteInferior) <> IsNull(clon.limiteInferior) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.limiteInferior & " = '" & Replace(vLimiteInferior, ",", ".") & "'"
        blnChange = True
    End If
    If vLimiteSuperior <> clon.limiteSuperior Or IsNull(vLimiteSuperior) <> IsNull(clon.limiteSuperior) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.limiteSuperior & " = '" & Replace(vLimiteSuperior, ",", ".") & "'"
        blnChange = True
    End If
    If vTarifa <> clon.tarifa Or IsNull(vTarifa) <> IsNull(clon.tarifa) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tarifa & " = '" & Replace(vTarifa, ",", ".") & "'"
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

    vFecha = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.categoria & " = " & vCategoria
    sql = sql & " AND " & ddl.rangoID & " = " & vRangoID
    sql = sql & " ORDER BY " & ddl.fecha & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vFecha = 1 + rstQuery.Fields(ddl.fecha)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionByRangoID(categoriaID As Integer, rangoID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.categoria & " = " & categoriaID
    sql = sql & " AND " & ddl.rangoID & " = " & rangoID
    sql = sql & " ORDER BY " & ddl.fecha
    sql = sql & ";"

    Set collectionByRangoID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyARango

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyARango
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
Dim elemento As clsMyARango

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
Dim elemento As clsMyARango

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

