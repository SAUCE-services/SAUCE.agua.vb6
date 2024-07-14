VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyATipoComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTipoID As Integer
Private vDescripcion As String
Private vModulo As Integer
Private vAplicaPendiente As Integer
Private vCuentaCorriente As Integer
Private vDebita As Integer
Private vIva As Integer
Private vAplicable As Integer
Private vLibroIva As Integer
Private vTipoComprobante As String
Private vRecibo As Integer
Private vContado As Integer
Private vPuntoVenta As Integer
Private vComprobanteID As Integer
Private vFacturaElectronica As Integer
Private vUid As String
Private vCreated As Date

Private clon As clsMyATipoComprobante

Public ddl As New clsMyDTipoComprobante

Private Sub class_Initialize()

    vTipoID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vDescripcion = ""
    vModulo = 0
    vAplicaPendiente = 0
    vCuentaCorriente = 0
    vDebita = 0
    vIva = 0
    vAplicable = 0
    vLibroIva = 0
    vTipoComprobante = ""
    vRecibo = 0
    vContado = 0
    vPuntoVenta = 0
    vComprobanteID = 0
    vFacturaElectronica = 0
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

Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

Public Property Get modulo() As Variant

    modulo = vModulo

End Property

Public Property Let modulo(ByVal vNewValue As Variant)

On Error Resume Next

    vModulo = vNewValue

End Property

Public Property Get aplicaPendiente() As Variant

    aplicaPendiente = vAplicaPendiente

End Property

Public Property Let aplicaPendiente(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicaPendiente = vNewValue

End Property

Public Property Get cuentaCorriente() As Variant

    cuentaCorriente = vCuentaCorriente

End Property

Public Property Let cuentaCorriente(ByVal vNewValue As Variant)

On Error Resume Next

    vCuentaCorriente = vNewValue

End Property

Public Property Get debita() As Variant

    debita = vDebita

End Property

Public Property Let debita(ByVal vNewValue As Variant)

On Error Resume Next

    vDebita = vNewValue

End Property

Public Property Get IVA() As Variant

    IVA = vIva

End Property

Public Property Let IVA(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

End Property

Public Property Get aplicable() As Variant

    aplicable = vAplicable

End Property

Public Property Let aplicable(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicable = vNewValue

End Property

Public Property Get libroIva() As Variant

    libroIva = vLibroIva

End Property

Public Property Let libroIva(ByVal vNewValue As Variant)

On Error Resume Next

    vLibroIva = vNewValue

End Property

Public Property Get tipocomprobante() As Variant

    tipocomprobante = vTipoComprobante

End Property

Public Property Let tipocomprobante(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoComprobante = vNewValue

End Property

Public Property Get recibo() As Variant

    recibo = vRecibo

End Property

Public Property Let recibo(ByVal vNewValue As Variant)

On Error Resume Next

    vRecibo = vNewValue

End Property

Public Property Get contado() As Variant

    contado = vContado

End Property

Public Property Let contado(ByVal vNewValue As Variant)

On Error Resume Next

    vContado = vNewValue

End Property

Public Property Get puntoVenta() As Variant

    puntoVenta = vPuntoVenta

End Property

Public Property Let puntoVenta(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntoVenta = vNewValue

End Property

Public Property Get comprobanteId() As Variant

    comprobanteId = vComprobanteID

End Property

Public Property Let comprobanteId(ByVal vNewValue As Variant)

On Error Resume Next

    vComprobanteID = vNewValue

End Property

Public Property Get facturaElectronica() As Variant

    facturaElectronica = vFacturaElectronica

End Property

Public Property Let facturaElectronica(ByVal vNewValue As Variant)

On Error Resume Next

    vFacturaElectronica = vNewValue

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

    comboKey = vTipoID

End Property

Public Property Get listText() As Variant

    listText = vDescripcion

End Property

Public Property Get listKey() As Variant

    listKey = vTipoID

End Property

Public Function clone() As clsMyATipoComprobante
Dim objObjeto As New clsMyATipoComprobante

    With objObjeto
        .tipoID = vTipoID
        .descripcion = vDescripcion
        .modulo = vModulo
        .aplicaPendiente = vAplicaPendiente
        .cuentaCorriente = vCuentaCorriente
        .debita = vDebita
        .IVA = vIva
        .aplicable = vAplicable
        .libroIva = vLibroIva
        .tipocomprobante = vTipoComprobante
        .recibo = vRecibo
        .contado = vContado
        .puntoVenta = vPuntoVenta
        .comprobanteId = vComprobanteID
        .facturaElectronica = vFacturaElectronica
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyATipoComprobante, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .tipoID = pQuery.Fields(.ddl.tipoID)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .modulo = pQuery.Fields(.ddl.modulo)
        .aplicaPendiente = pQuery.Fields(.ddl.aplicaPendiente)
        .cuentaCorriente = pQuery.Fields(.ddl.cuentaCorriente)
        .debita = pQuery.Fields(.ddl.debita)
        .IVA = pQuery.Fields(.ddl.IVA)
        .aplicable = pQuery.Fields(.ddl.aplicable)
        .libroIva = pQuery.Fields(.ddl.libroIva)
        .tipocomprobante = pQuery.Fields(.ddl.tipocomprobante)
        .recibo = pQuery.Fields(.ddl.recibo)
        .contado = pQuery.Fields(.ddl.contado)
        .puntoVenta = pQuery.Fields(.ddl.puntoVenta)
        .comprobanteId = pQuery.Fields(.ddl.comprobanteId)
        .facturaElectronica = pQuery.Fields(.ddl.facturaElectronica)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.tipoID & " = " & vTipoID

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

Public Sub findLastID(db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    clean

    sql = "SELECT LAST_INSERT_ID() AS " & ddl.tipoID
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vTipoID = rstQuery.Fields(ddl.tipoID)

    rstQuery.Close

    findByPrimaryKey db

End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.tipoID & " DESC"
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
    sql = sql & ", " & ddl.modulo & " = " & vModulo
    sql = sql & ", " & ddl.aplicaPendiente & " = " & vAplicaPendiente
    sql = sql & ", " & ddl.cuentaCorriente & " = " & vCuentaCorriente
    sql = sql & ", " & ddl.debita & " = " & vDebita
    sql = sql & ", " & ddl.IVA & " = " & vIva
    sql = sql & ", " & ddl.aplicable & " = " & vAplicable
    sql = sql & ", " & ddl.libroIva & " = " & vLibroIva
    sql = sql & ", " & ddl.tipocomprobante & " = '" & replaceSpecialSymbols(vTipoComprobante) & "'"
    sql = sql & ", " & ddl.recibo & " = " & vRecibo
    sql = sql & ", " & ddl.contado & " = " & vContado
    sql = sql & ", " & ddl.puntoVenta & " = " & vPuntoVenta
    sql = sql & ", " & ddl.comprobanteId & " = " & vComprobanteID
    sql = sql & ", " & ddl.facturaElectronica & " = " & vFacturaElectronica
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ";"

    add = db.execute(sql)

    findLastID db

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
    If vModulo <> clon.modulo Or IsNull(vModulo) <> IsNull(clon.modulo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.modulo & " = " & vModulo
        blnChange = True
    End If
    If vAplicaPendiente <> clon.aplicaPendiente Or IsNull(vAplicaPendiente) <> IsNull(clon.aplicaPendiente) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.aplicaPendiente & " = " & vAplicaPendiente
        blnChange = True
    End If
    If vCuentaCorriente <> clon.cuentaCorriente Or IsNull(vCuentaCorriente) <> IsNull(clon.cuentaCorriente) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuentaCorriente & " = " & vCuentaCorriente
        blnChange = True
    End If
    If vDebita <> clon.debita Or IsNull(vDebita) <> IsNull(clon.debita) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.debita & " = " & vDebita
        blnChange = True
    End If
    If vIva <> clon.IVA Or IsNull(vIva) <> IsNull(clon.IVA) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.IVA & " = " & vIva
        blnChange = True
    End If
    If vAplicable <> clon.aplicable Or IsNull(vAplicable) <> IsNull(clon.aplicable) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.aplicable & " = " & vAplicable
        blnChange = True
    End If
    If vLibroIva <> clon.libroIva Or IsNull(vLibroIva) <> IsNull(clon.libroIva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.libroIva & " = " & vLibroIva
        blnChange = True
    End If
    If vTipoComprobante <> clon.tipocomprobante Or IsNull(vTipoComprobante) <> IsNull(clon.tipocomprobante) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipocomprobante & " = '" & replaceSpecialSymbols(vTipoComprobante) & "'"
        blnChange = True
    End If
    If vRecibo <> clon.recibo Or IsNull(vRecibo) <> IsNull(clon.recibo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.recibo & " = " & vRecibo
        blnChange = True
    End If
    If vContado <> clon.contado Or IsNull(vContado) <> IsNull(clon.contado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.contado & " = " & vContado
        blnChange = True
    End If
    If vPuntoVenta <> clon.puntoVenta Or IsNull(vPuntoVenta) <> IsNull(clon.puntoVenta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.puntoVenta & " = " & vPuntoVenta
        blnChange = True
    End If
    If vComprobanteID <> clon.comprobanteId Or IsNull(vComprobanteID) <> IsNull(clon.comprobanteId) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.comprobanteId & " = " & vComprobanteID
        blnChange = True
    End If
    If vFacturaElectronica <> clon.facturaElectronica Or IsNull(vFacturaElectronica) <> IsNull(clon.facturaElectronica) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.facturaElectronica & " = " & vFacturaElectronica
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

    vTipoID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.tipoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vTipoID = 1 + rstQuery.Fields(ddl.tipoID)

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

Dim elemento As clsMyATipoComprobante

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyATipoComprobante
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
Dim elemento As clsMyATipoComprobante

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
Dim elemento As clsMyATipoComprobante

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

