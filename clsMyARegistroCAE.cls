VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyARegistroCAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTipoID As Integer
Private vPrefijo As Integer
Private vNumero As Long
Private vClienteID As Long
Private vTotal As Currency
Private vExento As Currency
Private vNeto27 As Currency
Private vNeto As Currency
Private vNeto105 As Currency
Private vIva27 As Currency
Private vIva As Currency
Private vIva105 As Currency
Private vCae As String
Private vFecha As String
Private vCaeVencimiento As String
Private vBarras As String
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyARegistroCAE

Public ddl As New clsMyDRegistroCAE

Private Sub class_Initialize()

    vTipoID = 0
    vPrefijo = 0
    vNumero = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vClienteID = 0
    vTotal = 0
    vExento = 0
    vNeto27 = 0
    vNeto = 0
    vNeto105 = 0
    vIva27 = 0
    vIva = 0
    vIva105 = 0
    vCae = ""
    vFecha = ""
    vCaeVencimiento = ""
    vBarras = ""
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get tipoID() As Variant

    tipoID = vTipoID

End Property

Public Property Let tipoID(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoID = vNewValue

End Property

Public Property Get prefijo() As Variant

    prefijo = vPrefijo

End Property

Public Property Let prefijo(ByVal vNewValue As Variant)

On Error Resume Next

    vPrefijo = vNewValue

End Property

Public Property Get numero() As Variant

    numero = vNumero

End Property

Public Property Let numero(ByVal vNewValue As Variant)

On Error Resume Next

    vNumero = vNewValue

End Property

Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

Public Property Get total() As Variant

    total = vTotal

End Property

Public Property Let total(ByVal vNewValue As Variant)

On Error Resume Next

    vTotal = vNewValue

End Property

Public Property Get exento() As Variant

    exento = vExento

End Property

Public Property Let exento(ByVal vNewValue As Variant)

On Error Resume Next

    vExento = vNewValue

End Property

Public Property Get neto27() As Variant

    neto27 = vNeto27

End Property

Public Property Let neto27(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto27 = vNewValue

End Property

Public Property Get neto() As Variant

    neto = vNeto

End Property

Public Property Let neto(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto = vNewValue

End Property

Public Property Get neto105() As Variant

    neto105 = vNeto105

End Property

Public Property Let neto105(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto105 = vNewValue

End Property

Public Property Get iva27() As Variant

    iva27 = vIva27

End Property

Public Property Let iva27(ByVal vNewValue As Variant)

On Error Resume Next

    vIva27 = vNewValue

End Property

Public Property Get IVA() As Variant

    IVA = vIva

End Property

Public Property Let IVA(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

End Property

Public Property Get iva105() As Variant

    iva105 = vIva105

End Property

Public Property Let iva105(ByVal vNewValue As Variant)

On Error Resume Next

    vIva105 = vNewValue

End Property

Public Property Get cae() As Variant

    cae = vCae

End Property

Public Property Let cae(ByVal vNewValue As Variant)

On Error Resume Next

    vCae = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get caeVencimiento() As Variant

    caeVencimiento = vCaeVencimiento

End Property

Public Property Let caeVencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vCaeVencimiento = vNewValue

End Property

Public Property Get barras() As Variant

    barras = vBarras

End Property

Public Property Let barras(ByVal vNewValue As Variant)

On Error Resume Next

    vBarras = vNewValue

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

    comboText = vClienteID

End Property

Public Property Get comboKey() As Variant

    comboKey = vNumero

End Property

Public Property Get listText() As Variant

    listText = vClienteID

End Property

Public Property Get listKey() As Variant

    listKey = vNumero

End Property

Public Function clone() As clsMyARegistroCAE
Dim objObjeto As New clsMyARegistroCAE

    With objObjeto
        .tipoID = vTipoID
        .prefijo = vPrefijo
        .numero = vNumero
        .clienteID = vClienteID
        .total = vTotal
        .exento = vExento
        .neto27 = vNeto27
        .neto = vNeto
        .neto105 = vNeto105
        .iva27 = vIva27
        .IVA = vIva
        .iva105 = vIva105
        .cae = vCae
        .fecha = vFecha
        .caeVencimiento = vCaeVencimiento
        .barras = vBarras
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyARegistroCAE, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .tipoID = pQuery.Fields(.ddl.tipoID)
        .prefijo = pQuery.Fields(.ddl.prefijo)
        .numero = pQuery.Fields(.ddl.numero)
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .total = pQuery.Fields(.ddl.total)
        .exento = pQuery.Fields(.ddl.exento)
        .neto27 = pQuery.Fields(.ddl.neto27)
        .neto = pQuery.Fields(.ddl.neto)
        .neto105 = pQuery.Fields(.ddl.neto105)
        .iva27 = pQuery.Fields(.ddl.iva27)
        .IVA = pQuery.Fields(.ddl.IVA)
        .iva105 = pQuery.Fields(.ddl.iva105)
        .cae = pQuery.Fields(.ddl.cae)
        .fecha = pQuery.Fields(.ddl.fecha)
        .caeVencimiento = pQuery.Fields(.ddl.caeVencimiento)
        .barras = pQuery.Fields(.ddl.barras)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " AND " & ddl.numero & " = " & vNumero

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
    sql = sql & " WHERE " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " ORDER BY " & ddl.numero & " DESC"
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
    sql = sql & " SET " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
    sql = sql & ", " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
    sql = sql & ", " & ddl.neto27 & " = '" & Replace(vNeto27, ",", ".") & "'"
    sql = sql & ", " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
    sql = sql & ", " & ddl.neto105 & " = '" & Replace(vNeto105, ",", ".") & "'"
    sql = sql & ", " & ddl.iva27 & " = '" & Replace(vIva27, ",", ".") & "'"
    sql = sql & ", " & ddl.IVA & " = '" & Replace(vIva, ",", ".") & "'"
    sql = sql & ", " & ddl.iva105 & " = '" & Replace(vIva105, ",", ".") & "'"
    sql = sql & ", " & ddl.cae & " = '" & replaceSpecialSymbols(vCae) & "'"
    sql = sql & ", " & ddl.fecha & " = '" & replaceSpecialSymbols(vFecha) & "'"
    sql = sql & ", " & ddl.caeVencimiento & " = '" & replaceSpecialSymbols(vCaeVencimiento) & "'"
    sql = sql & ", " & ddl.barras & " = '" & replaceSpecialSymbols(vBarras) & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.tipoID & " = " & vTipoID
    sql = sql & ", " & ddl.prefijo & " = " & vPrefijo
    sql = sql & ", " & ddl.numero & " = " & vNumero
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

    If vClienteID <> clon.clienteID Or IsNull(vClienteID) <> IsNull(clon.clienteID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clienteID & " = " & vClienteID
        blnChange = True
    End If
    If vTotal <> clon.total Or IsNull(vTotal) <> IsNull(clon.total) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
        blnChange = True
    End If
    If vExento <> clon.exento Or IsNull(vExento) <> IsNull(clon.exento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
        blnChange = True
    End If
    If vNeto27 <> clon.neto27 Or IsNull(vNeto27) <> IsNull(clon.neto27) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.neto27 & " = '" & Replace(vNeto27, ",", ".") & "'"
        blnChange = True
    End If
    If vNeto <> clon.neto Or IsNull(vNeto) <> IsNull(clon.neto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
        blnChange = True
    End If
    If vNeto105 <> clon.neto105 Or IsNull(vNeto105) <> IsNull(clon.neto105) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.neto105 & " = '" & Replace(vNeto105, ",", ".") & "'"
        blnChange = True
    End If
    If vIva27 <> clon.iva27 Or IsNull(vIva27) <> IsNull(clon.iva27) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva27 & " = '" & Replace(vIva27, ",", ".") & "'"
        blnChange = True
    End If
    If vIva <> clon.IVA Or IsNull(vIva) <> IsNull(clon.IVA) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.IVA & " = '" & Replace(vIva, ",", ".") & "'"
        blnChange = True
    End If
    If vIva105 <> clon.iva105 Or IsNull(vIva105) <> IsNull(clon.iva105) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva105 & " = '" & Replace(vIva105, ",", ".") & "'"
        blnChange = True
    End If
    If vCae <> clon.cae Or IsNull(vCae) <> IsNull(clon.cae) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cae & " = '" & replaceSpecialSymbols(vCae) & "'"
        blnChange = True
    End If
    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = '" & replaceSpecialSymbols(vFecha) & "'"
        blnChange = True
    End If
    If vCaeVencimiento <> clon.caeVencimiento Or IsNull(vCaeVencimiento) <> IsNull(clon.caeVencimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.caeVencimiento & " = '" & replaceSpecialSymbols(vCaeVencimiento) & "'"
        blnChange = True
    End If
    If vBarras <> clon.barras Or IsNull(vBarras) <> IsNull(clon.barras) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.barras & " = '" & replaceSpecialSymbols(vBarras) & "'"
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

    vNumero = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " ORDER BY " & ddl.numero & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vNumero = 1 + rstQuery.Fields(ddl.numero)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyARegistroCAE

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyARegistroCAE
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
Dim elemento As clsMyARegistroCAE

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
Dim elemento As clsMyARegistroCAE

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

