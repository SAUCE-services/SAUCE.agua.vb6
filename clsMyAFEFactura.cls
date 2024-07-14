VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAFEFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTipoID As Integer
Private vPrefijo As Integer
Private vNumero As Long
Private vClienteID As Long
Private vFecha As Date
Private vImporte As Currency
Private vNeto27 As Currency
Private vNeto As Currency
Private vIva27 As Currency
Private vIva As Currency
Private vExento As Currency
Private vRecibo As Integer
Private vAnulada As Integer
Private vTipoCompro As String
Private vLetras As String
Private vObservaciones As String
Private vCae As String
Private vCaeVencimiento As String
Private vCaeBarras As String
Private vPuntoVta As Integer
Private vNroComprob As Long
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyAFEFactura

Public ddl As New clsMyDFEFactura

Private Sub class_Initialize()

    vTipoID = 0
    vPrefijo = 0
    vNumero = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vClienteID = 0
    vFecha = Date
    vImporte = 0
    vNeto27 = 0
    vNeto = 0
    vIva27 = 0
    vIva = 0
    vExento = 0
    vRecibo = 0
    vAnulada = 0
    vTipoCompro = ""
    vLetras = ""
    vObservaciones = ""
    vCae = ""
    vCaeVencimiento = ""
    vCaeBarras = ""
    vPuntoVta = 0
    vNroComprob = 0
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

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get importe() As Variant

    importe = vImporte

End Property

Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

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

Public Property Get exento() As Variant

    exento = vExento

End Property

Public Property Let exento(ByVal vNewValue As Variant)

On Error Resume Next

    vExento = vNewValue

End Property

Public Property Get recibo() As Variant

    recibo = vRecibo

End Property

Public Property Let recibo(ByVal vNewValue As Variant)

On Error Resume Next

    vRecibo = vNewValue

End Property

Public Property Get anulada() As Variant

    anulada = vAnulada

End Property

Public Property Let anulada(ByVal vNewValue As Variant)

On Error Resume Next

    vAnulada = vNewValue

End Property

Public Property Get tipoCompro() As Variant

    tipoCompro = vTipoCompro

End Property

Public Property Let tipoCompro(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoCompro = vNewValue

End Property

Public Property Get letras() As Variant

    letras = vLetras

End Property

Public Property Let letras(ByVal vNewValue As Variant)

On Error Resume Next

    vLetras = vNewValue

End Property

Public Property Get observaciones() As Variant

    observaciones = vObservaciones

End Property

Public Property Let observaciones(ByVal vNewValue As Variant)

On Error Resume Next

    vObservaciones = vNewValue

End Property

Public Property Get cae() As Variant

    cae = vCae

End Property

Public Property Let cae(ByVal vNewValue As Variant)

On Error Resume Next

    vCae = vNewValue

End Property

Public Property Get caeVencimiento() As Variant

    caeVencimiento = vCaeVencimiento

End Property

Public Property Let caeVencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vCaeVencimiento = vNewValue

End Property

Public Property Get caeBarras() As Variant

    caeBarras = vCaeBarras

End Property

Public Property Let caeBarras(ByVal vNewValue As Variant)

On Error Resume Next

    vCaeBarras = vNewValue

End Property

Public Property Get puntoVta() As Variant

    puntoVta = vPuntoVta

End Property

Public Property Let puntoVta(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntoVta = vNewValue

End Property

Public Property Get nroComprob() As Variant

    nroComprob = vNroComprob

End Property

Public Property Let nroComprob(ByVal vNewValue As Variant)

On Error Resume Next

    vNroComprob = vNewValue

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

Public Function clone() As clsMyAFEFactura
Dim objObjeto As New clsMyAFEFactura

    With objObjeto
        .tipoID = vTipoID
        .prefijo = vPrefijo
        .numero = vNumero
        .clienteID = vClienteID
        .fecha = vFecha
        .importe = vImporte
        .neto27 = vNeto27
        .neto = vNeto
        .iva27 = vIva27
        .IVA = vIva
        .exento = vExento
        .recibo = vRecibo
        .anulada = vAnulada
        .tipoCompro = vTipoCompro
        .letras = vLetras
        .observaciones = vObservaciones
        .cae = vCae
        .caeVencimiento = vCaeVencimiento
        .caeBarras = vCaeBarras
        .puntoVta = vPuntoVta
        .nroComprob = vNroComprob
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAFEFactura, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .tipoID = pQuery.Fields(.ddl.tipoID)
        .prefijo = pQuery.Fields(.ddl.prefijo)
        .numero = pQuery.Fields(.ddl.numero)
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .importe = pQuery.Fields(.ddl.importe)
        .neto27 = pQuery.Fields(.ddl.neto27)
        .neto = pQuery.Fields(.ddl.neto)
        .iva27 = pQuery.Fields(.ddl.iva27)
        .IVA = pQuery.Fields(.ddl.IVA)
        .exento = pQuery.Fields(.ddl.exento)
        .recibo = pQuery.Fields(.ddl.recibo)
        .anulada = pQuery.Fields(.ddl.anulada)
        .tipoCompro = pQuery.Fields(.ddl.tipoCompro)
        .letras = pQuery.Fields(.ddl.letras)
        .observaciones = pQuery.Fields(.ddl.observaciones)
        .cae = pQuery.Fields(.ddl.cae)
        .caeVencimiento = pQuery.Fields(.ddl.caeVencimiento)
        .caeBarras = pQuery.Fields(.ddl.caeBarras)
        .puntoVta = pQuery.Fields(.ddl.puntoVta)
        .nroComprob = pQuery.Fields(.ddl.nroComprob)
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

Public Sub findByLiquidacion(db As clsDB)
Dim sql As String

Dim intPuntoVta As Integer
Dim lngNroComprob As Long

    intPuntoVta = vPuntoVta
    lngNroComprob = vNroComprob

    clean

    vPuntoVta = intPuntoVta
    vNroComprob = lngNroComprob

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & " AND " & ddl.nroComprob & " = " & vNroComprob
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
    sql = sql & ", " & ddl.fecha & " = " & fechaDB(vFecha)
    sql = sql & ", " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
    sql = sql & ", " & ddl.neto27 & " = '" & Replace(vNeto27, ",", ".") & "'"
    sql = sql & ", " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
    sql = sql & ", " & ddl.iva27 & " = '" & Replace(vIva27, ",", ".") & "'"
    sql = sql & ", " & ddl.IVA & " = '" & Replace(vIva, ",", ".") & "'"
    sql = sql & ", " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
    sql = sql & ", " & ddl.recibo & " = " & vRecibo
    sql = sql & ", " & ddl.anulada & " = " & vAnulada
    sql = sql & ", " & ddl.tipoCompro & " = '" & replaceSpecialSymbols(vTipoCompro) & "'"
    sql = sql & ", " & ddl.letras & " = '" & replaceSpecialSymbols(vLetras) & "'"
    sql = sql & ", " & ddl.observaciones & " = '" & replaceSpecialSymbols(vObservaciones) & "'"
    sql = sql & ", " & ddl.cae & " = '" & replaceSpecialSymbols(vCae) & "'"
    sql = sql & ", " & ddl.caeVencimiento & " = '" & replaceSpecialSymbols(vCaeVencimiento) & "'"
    sql = sql & ", " & ddl.caeBarras & " = '" & replaceSpecialSymbols(vCaeBarras) & "'"
    sql = sql & ", " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & ", " & ddl.nroComprob & " = " & vNroComprob
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
    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = " & fechaDB(vFecha)
        blnChange = True
    End If
    If vImporte <> clon.importe Or IsNull(vImporte) <> IsNull(clon.importe) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
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
    If vExento <> clon.exento Or IsNull(vExento) <> IsNull(clon.exento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
        blnChange = True
    End If
    If vRecibo <> clon.recibo Or IsNull(vRecibo) <> IsNull(clon.recibo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.recibo & " = " & vRecibo
        blnChange = True
    End If
    If vAnulada <> clon.anulada Or IsNull(vAnulada) <> IsNull(clon.anulada) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.anulada & " = " & vAnulada
        blnChange = True
    End If
    If vTipoCompro <> clon.tipoCompro Or IsNull(vTipoCompro) <> IsNull(clon.tipoCompro) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipoCompro & " = '" & replaceSpecialSymbols(vTipoCompro) & "'"
        blnChange = True
    End If
    If vLetras <> clon.letras Or IsNull(vLetras) <> IsNull(clon.letras) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.letras & " = '" & replaceSpecialSymbols(vLetras) & "'"
        blnChange = True
    End If
    If vObservaciones <> clon.observaciones Or IsNull(vObservaciones) <> IsNull(clon.observaciones) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.observaciones & " = '" & replaceSpecialSymbols(vObservaciones) & "'"
        blnChange = True
    End If
    If vCae <> clon.cae Or IsNull(vCae) <> IsNull(clon.cae) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cae & " = '" & replaceSpecialSymbols(vCae) & "'"
        blnChange = True
    End If
    If vCaeVencimiento <> clon.caeVencimiento Or IsNull(vCaeVencimiento) <> IsNull(clon.caeVencimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.caeVencimiento & " = '" & replaceSpecialSymbols(vCaeVencimiento) & "'"
        blnChange = True
    End If
    If vCaeBarras <> clon.caeBarras Or IsNull(vCaeBarras) <> IsNull(clon.caeBarras) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.caeBarras & " = '" & replaceSpecialSymbols(vCaeBarras) & "'"
        blnChange = True
    End If
    If vPuntoVta <> clon.puntoVta Or IsNull(vPuntoVta) <> IsNull(clon.puntoVta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.puntoVta & " = " & vPuntoVta
        blnChange = True
    End If
    If vNroComprob <> clon.nroComprob Or IsNull(vNroComprob) <> IsNull(clon.nroComprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nroComprob & " = " & vNroComprob
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

Dim elemento As clsMyAFEFactura

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAFEFactura
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
Dim elemento As clsMyAFEFactura

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
Dim elemento As clsMyAFEFactura

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

