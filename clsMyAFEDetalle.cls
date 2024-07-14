VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAFEDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTipoID As Integer
Private vPrefijo As Integer
Private vNumero As Long
Private vItem As Integer
Private vRubroID As Integer
Private vCantidad As Currency
Private vUnitarioSinIva As Currency
Private vUnitarioConIva As Currency
Private vIva105 As Integer
Private vExento As Integer
Private vFecha As Date
Private vConcepto As String
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyAFEDetalle

Public ddl As New clsMyDFEDetalle

Private Sub class_Initialize()

    vTipoID = 0
    vPrefijo = 0
    vNumero = 0
    vItem = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vRubroID = 0
    vCantidad = 0
    vUnitarioSinIva = 0
    vUnitarioConIva = 0
    vIva105 = 0
    vExento = 0
    vFecha = Date
    vConcepto = ""
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

Public Property Get item() As Variant

    item = vItem

End Property

Public Property Let item(ByVal vNewValue As Variant)

On Error Resume Next

    vItem = vNewValue

End Property

Public Property Get rubroID() As Variant

    rubroID = vRubroID

End Property

Public Property Let rubroID(ByVal vNewValue As Variant)

On Error Resume Next

    vRubroID = vNewValue

End Property

Public Property Get cantidad() As Variant

    cantidad = vCantidad

End Property

Public Property Let cantidad(ByVal vNewValue As Variant)

On Error Resume Next

    vCantidad = vNewValue

End Property

Public Property Get unitarioSinIva() As Variant

    unitarioSinIva = vUnitarioSinIva

End Property

Public Property Let unitarioSinIva(ByVal vNewValue As Variant)

On Error Resume Next

    vUnitarioSinIva = vNewValue

End Property

Public Property Get unitarioConIva() As Variant

    unitarioConIva = vUnitarioConIva

End Property

Public Property Let unitarioConIva(ByVal vNewValue As Variant)

On Error Resume Next

    vUnitarioConIva = vNewValue

End Property

Public Property Get iva105() As Variant

    iva105 = vIva105

End Property

Public Property Let iva105(ByVal vNewValue As Variant)

On Error Resume Next

    vIva105 = vNewValue

End Property

Public Property Get exento() As Variant

    exento = vExento

End Property

Public Property Let exento(ByVal vNewValue As Variant)

On Error Resume Next

    vExento = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get concepto() As Variant

    concepto = vConcepto

End Property

Public Property Let concepto(ByVal vNewValue As Variant)

On Error Resume Next

    vConcepto = vNewValue

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

    comboText = vRubroID

End Property

Public Property Get comboKey() As Variant

    comboKey = vItem

End Property

Public Property Get listText() As Variant

    listText = vRubroID

End Property

Public Property Get listKey() As Variant

    listKey = vItem

End Property

Public Function clone() As clsMyAFEDetalle
Dim objObjeto As New clsMyAFEDetalle

    With objObjeto
        .tipoID = vTipoID
        .prefijo = vPrefijo
        .numero = vNumero
        .item = vItem
        .rubroID = vRubroID
        .cantidad = vCantidad
        .unitarioSinIva = vUnitarioSinIva
        .unitarioConIva = vUnitarioConIva
        .iva105 = vIva105
        .exento = vExento
        .fecha = vFecha
        .concepto = vConcepto
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAFEDetalle, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .tipoID = pQuery.Fields(.ddl.tipoID)
        .prefijo = pQuery.Fields(.ddl.prefijo)
        .numero = pQuery.Fields(.ddl.numero)
        .item = pQuery.Fields(.ddl.item)
        .rubroID = pQuery.Fields(.ddl.rubroID)
        .cantidad = pQuery.Fields(.ddl.cantidad)
        .unitarioSinIva = pQuery.Fields(.ddl.unitarioSinIva)
        .unitarioConIva = pQuery.Fields(.ddl.unitarioConIva)
        .iva105 = pQuery.Fields(.ddl.iva105)
        .exento = pQuery.Fields(.ddl.exento)
        .fecha = pQuery.Fields(.ddl.fecha)
        .concepto = pQuery.Fields(.ddl.concepto)
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
    sql = sql & " AND " & ddl.item & " = " & vItem

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
    sql = sql & " AND " & ddl.numero & " = " & vNumero
    sql = sql & " ORDER BY " & ddl.item & " DESC"
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
    sql = sql & " SET " & ddl.rubroID & " = " & vRubroID
    sql = sql & ", " & ddl.cantidad & " = " & vCantidad
    sql = sql & ", " & ddl.unitarioSinIva & " = " & vUnitarioSinIva
    sql = sql & ", " & ddl.unitarioConIva & " = " & vUnitarioConIva
    sql = sql & ", " & ddl.iva105 & " = " & vIva105
    sql = sql & ", " & ddl.exento & " = " & vExento
    sql = sql & ", " & ddl.fecha & " = " & fechaDB(vFecha)
    sql = sql & ", " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.tipoID & " = " & vTipoID
    sql = sql & ", " & ddl.prefijo & " = " & vPrefijo
    sql = sql & ", " & ddl.numero & " = " & vNumero
    sql = sql & ", " & ddl.item & " = " & vItem
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

    If vRubroID <> clon.rubroID Or IsNull(vRubroID) <> IsNull(clon.rubroID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.rubroID & " = " & vRubroID
        blnChange = True
    End If
    If vCantidad <> clon.cantidad Or IsNull(vCantidad) <> IsNull(clon.cantidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cantidad & " = " & vCantidad
        blnChange = True
    End If
    If vUnitarioSinIva <> clon.unitarioSinIva Or IsNull(vUnitarioSinIva) <> IsNull(clon.unitarioSinIva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.unitarioSinIva & " = " & vUnitarioSinIva
        blnChange = True
    End If
    If vUnitarioConIva <> clon.unitarioConIva Or IsNull(vUnitarioConIva) <> IsNull(clon.unitarioConIva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.unitarioConIva & " = " & vUnitarioConIva
        blnChange = True
    End If
    If vIva105 <> clon.iva105 Or IsNull(vIva105) <> IsNull(clon.iva105) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva105 & " = " & vIva105
        blnChange = True
    End If
    If vExento <> clon.exento Or IsNull(vExento) <> IsNull(clon.exento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.exento & " = " & vExento
        blnChange = True
    End If
    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = " & fechaDB(vFecha)
        blnChange = True
    End If
    If vConcepto <> clon.concepto Or IsNull(vConcepto) <> IsNull(clon.concepto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
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

    vItem = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tipoID & " = " & vTipoID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " AND " & ddl.numero & " = " & vNumero
    sql = sql & " ORDER BY " & ddl.item & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vItem = 1 + rstQuery.Fields(ddl.item)

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

Dim elemento As clsMyAFEDetalle

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAFEDetalle
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
Dim elemento As clsMyAFEDetalle

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
Dim elemento As clsMyAFEDetalle

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

