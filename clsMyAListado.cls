VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vC1 As String
Private vC2 As String
Private vC3 As String
Private vC4 As String
Private vC5 As String
Private vC6 As String
Private vN1 As Currency
Private vN2 As Currency
Private vN3 As Currency
Private vN4 As Currency
Private vN5 As Currency
Private vN6 As Currency
Private vAutoID As Long
Private vCreated As Date

Private clon As clsMyAListado

Public ddl As New clsMyDListado

Private Sub class_Initialize()

    vAutoID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vC1 = ""
    vC2 = ""
    vC3 = ""
    vC4 = ""
    vC5 = ""
    vC6 = ""
    vN1 = 0
    vN2 = 0
    vN3 = 0
    vN4 = 0
    vN5 = 0
    vN6 = 0
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get c1() As Variant

    c1 = vC1

End Property

Public Property Let c1(ByVal vNewValue As Variant)

On Error Resume Next

    vC1 = vNewValue

End Property

Public Property Get c2() As Variant

    c2 = vC2

End Property

Public Property Let c2(ByVal vNewValue As Variant)

On Error Resume Next

    vC2 = vNewValue

End Property

Public Property Get c3() As Variant

    c3 = vC3

End Property

Public Property Let c3(ByVal vNewValue As Variant)

On Error Resume Next

    vC3 = vNewValue

End Property

Public Property Get c4() As Variant

    c4 = vC4

End Property

Public Property Let c4(ByVal vNewValue As Variant)

On Error Resume Next

    vC4 = vNewValue

End Property

Public Property Get c5() As Variant

    c5 = vC5

End Property

Public Property Let c5(ByVal vNewValue As Variant)

On Error Resume Next

    vC5 = vNewValue

End Property

Public Property Get c6() As Variant

    c6 = vC6

End Property

Public Property Let c6(ByVal vNewValue As Variant)

On Error Resume Next

    vC6 = vNewValue

End Property

Public Property Get n1() As Variant

    n1 = vN1

End Property

Public Property Let n1(ByVal vNewValue As Variant)

On Error Resume Next

    vN1 = vNewValue

End Property

Public Property Get n2() As Variant

    n2 = vN2

End Property

Public Property Let n2(ByVal vNewValue As Variant)

On Error Resume Next

    vN2 = vNewValue

End Property

Public Property Get n3() As Variant

    n3 = vN3

End Property

Public Property Let n3(ByVal vNewValue As Variant)

On Error Resume Next

    vN3 = vNewValue

End Property

Public Property Get n4() As Variant

    n4 = vN4

End Property

Public Property Let n4(ByVal vNewValue As Variant)

On Error Resume Next

    vN4 = vNewValue

End Property

Public Property Get n5() As Variant

    n5 = vN5

End Property

Public Property Let n5(ByVal vNewValue As Variant)

On Error Resume Next

    vN5 = vNewValue

End Property

Public Property Get n6() As Variant

    n6 = vN6

End Property

Public Property Let n6(ByVal vNewValue As Variant)

On Error Resume Next

    vN6 = vNewValue

End Property

Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

End Property

Public Property Get created() As Variant

    created = vCreated

End Property

Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vC1

End Property

Public Property Get comboKey() As Variant

    comboKey = vAutoID

End Property

Public Property Get listText() As Variant

    listText = vC1

End Property

Public Property Get listKey() As Variant

    listKey = vAutoID

End Property

Public Function clone() As clsMyAListado
Dim objObjeto As New clsMyAListado

    With objObjeto
        .c1 = vC1
        .c2 = vC2
        .c3 = vC3
        .c4 = vC4
        .c5 = vC5
        .c6 = vC6
        .n1 = vN1
        .n2 = vN2
        .n3 = vN3
        .n4 = vN4
        .n5 = vN5
        .n6 = vN6
        .autoID = vAutoID
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAListado, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .c1 = pQuery.Fields(.ddl.c1)
        .c2 = pQuery.Fields(.ddl.c2)
        .c3 = pQuery.Fields(.ddl.c3)
        .c4 = pQuery.Fields(.ddl.c4)
        .c5 = pQuery.Fields(.ddl.c5)
        .c6 = pQuery.Fields(.ddl.c6)
        .n1 = pQuery.Fields(.ddl.n1)
        .n2 = pQuery.Fields(.ddl.n2)
        .n3 = pQuery.Fields(.ddl.n3)
        .n4 = pQuery.Fields(.ddl.n4)
        .n5 = pQuery.Fields(.ddl.n5)
        .n6 = pQuery.Fields(.ddl.n6)
        .autoID = pQuery.Fields(.ddl.autoID)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.autoID & " = " & vAutoID

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

    findByPrimaryKey_ IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub findByPrimaryKey_(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLastID(db As clsDB)

    findLastID_ IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub findLastID_(db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    clean

    sql = "SELECT LAST_INSERT_ID() AS " & ddl.autoID
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vAutoID = rstQuery.Fields(ddl.autoID)

    rstQuery.Close

    findByPrimaryKey_ db

End Sub

Public Sub findLast(db As clsDB)

    findLast_ IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub findLast_(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.autoID & " DESC"
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

Public Function truncate(db As clsDB) As Boolean
Dim sql As String

    sql = "TRUNCATE " & ddl.table
    sql = sql & ";"

    truncate = db.execute(sql)

End Function

Public Function save(db As clsDB) As Boolean

    If exist_(db) Then
        save = update(db)
    Else
        save = add(db)
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
    sql = sql & " SET " & ddl.c1 & " = '" & replaceSpecialSymbols(vC1) & "'"
    sql = sql & ", " & ddl.c2 & " = '" & replaceSpecialSymbols(vC2) & "'"
    sql = sql & ", " & ddl.c3 & " = '" & replaceSpecialSymbols(vC3) & "'"
    sql = sql & ", " & ddl.c4 & " = '" & replaceSpecialSymbols(vC4) & "'"
    sql = sql & ", " & ddl.c5 & " = '" & replaceSpecialSymbols(vC5) & "'"
    sql = sql & ", " & ddl.c6 & " = '" & replaceSpecialSymbols(vC6) & "'"
    sql = sql & ", " & ddl.n1 & " = '" & Replace(vN1, ",", ".") & "'"
    sql = sql & ", " & ddl.n2 & " = '" & Replace(vN2, ",", ".") & "'"
    sql = sql & ", " & ddl.n3 & " = '" & Replace(vN3, ",", ".") & "'"
    sql = sql & ", " & ddl.n4 & " = '" & Replace(vN4, ",", ".") & "'"
    sql = sql & ", " & ddl.n5 & " = '" & Replace(vN5, ",", ".") & "'"
    sql = sql & ", " & ddl.n6 & " = '" & Replace(vN6, ",", ".") & "'"
    sql = sql & ";"

    add_ = db.execute(sql)

    findLastID_ db

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

    If vC1 <> clon.c1 Or IsNull(vC1) <> IsNull(clon.c1) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.c1 & " = '" & replaceSpecialSymbols(vC1) & "'"
        blnChange = True
    End If
    If vC2 <> clon.c2 Or IsNull(vC2) <> IsNull(clon.c2) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.c2 & " = '" & replaceSpecialSymbols(vC2) & "'"
        blnChange = True
    End If
    If vC3 <> clon.c3 Or IsNull(vC3) <> IsNull(clon.c3) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.c3 & " = '" & replaceSpecialSymbols(vC3) & "'"
        blnChange = True
    End If
    If vC4 <> clon.c4 Or IsNull(vC4) <> IsNull(clon.c4) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.c4 & " = '" & replaceSpecialSymbols(vC4) & "'"
        blnChange = True
    End If
    If vC5 <> clon.c5 Or IsNull(vC5) <> IsNull(clon.c5) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.c5 & " = '" & replaceSpecialSymbols(vC5) & "'"
        blnChange = True
    End If
    If vC6 <> clon.c6 Or IsNull(vC6) <> IsNull(clon.c6) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.c6 & " = '" & replaceSpecialSymbols(vC6) & "'"
        blnChange = True
    End If
    If vN1 <> clon.n1 Or IsNull(vN1) <> IsNull(clon.n1) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.n1 & " = '" & Replace(vN1, ",", ".") & "'"
        blnChange = True
    End If
    If vN2 <> clon.n2 Or IsNull(vN2) <> IsNull(clon.n2) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.n2 & " = '" & Replace(vN2, ",", ".") & "'"
        blnChange = True
    End If
    If vN3 <> clon.n3 Or IsNull(vN3) <> IsNull(clon.n3) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.n3 & " = '" & Replace(vN3, ",", ".") & "'"
        blnChange = True
    End If
    If vN4 <> clon.n4 Or IsNull(vN4) <> IsNull(clon.n4) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.n4 & " = '" & Replace(vN4, ",", ".") & "'"
        blnChange = True
    End If
    If vN5 <> clon.n5 Or IsNull(vN5) <> IsNull(clon.n5) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.n5 & " = '" & Replace(vN5, ",", ".") & "'"
        blnChange = True
    End If
    If vN6 <> clon.n6 Or IsNull(vN6) <> IsNull(clon.n6) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.n6 & " = '" & Replace(vN6, ",", ".") & "'"
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

    vAutoID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.autoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vAutoID = 1 + rstQuery.Fields(ddl.autoID)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ db

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

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAListado

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAListado
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
Dim elemento As clsMyAListado

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
Dim elemento As clsMyAListado

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

