VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyANCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vSerieID As Integer
Private vNumero As Long
Private vFecha As Date
Private vClienteID As Long
Private vSituacionIva As Integer
Private vAnulado As Integer
Private vTotal As Currency
Private vPuntoVta As Integer
Private vNroComprob As Long
Private vIvaCF As Currency
Private vIvaRI As Currency
Private vIvaRN As Currency
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyANCredito

Public ddl As New clsMyDNCredito

Private Sub class_Initialize()

    vSerieID = 0
    vNumero = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFecha = Date
    vClienteID = 0
    vSituacionIva = 0
    vAnulado = 0
    vTotal = 0
    vPuntoVta = 0
    vNroComprob = 0
    vIvaCF = 0
    vIvaRI = 0
    vIvaRN = 0
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

Public Property Get numero() As Variant

    numero = vNumero

End Property

Public Property Let numero(ByVal vNewValue As Variant)

On Error Resume Next

    vNumero = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

Public Property Get situacionIVA() As Variant

    situacionIVA = vSituacionIva

End Property

Public Property Let situacionIVA(ByVal vNewValue As Variant)

On Error Resume Next

    vSituacionIva = vNewValue

End Property

Public Property Get anulado() As Variant

    anulado = vAnulado

End Property

Public Property Let anulado(ByVal vNewValue As Variant)

On Error Resume Next

    vAnulado = vNewValue

End Property

Public Property Get total() As Variant

    total = vTotal

End Property

Public Property Let total(ByVal vNewValue As Variant)

On Error Resume Next

    vTotal = vNewValue

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

Public Property Get ivaCF() As Variant

    ivaCF = vIvaCF

End Property

Public Property Let ivaCF(ByVal vNewValue As Variant)

On Error Resume Next

    vIvaCF = vNewValue

End Property

Public Property Get ivaRI() As Variant

    ivaRI = vIvaRI

End Property

Public Property Let ivaRI(ByVal vNewValue As Variant)

On Error Resume Next

    vIvaRI = vNewValue

End Property

Public Property Get ivaRN() As Variant

    ivaRN = vIvaRN

End Property

Public Property Let ivaRN(ByVal vNewValue As Variant)

On Error Resume Next

    vIvaRN = vNewValue

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

    comboKey = vNumero

End Property

Public Property Get listText() As Variant

    listText = vFecha

End Property

Public Property Get listKey() As Variant

    listKey = vNumero

End Property

Public Function clone() As clsMyANCredito
Dim objObjeto As New clsMyANCredito

    With objObjeto
        .serieID = vSerieID
        .numero = vNumero
        .fecha = vFecha
        .clienteID = vClienteID
        .situacionIVA = vSituacionIva
        .anulado = vAnulado
        .total = vTotal
        .puntoVta = vPuntoVta
        .nroComprob = vNroComprob
        .ivaCF = vIvaCF
        .ivaRI = vIvaRI
        .ivaRN = vIvaRN
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyANCredito, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .serieID = pQuery.Fields(.ddl.serieID)
        .numero = pQuery.Fields(.ddl.numero)
        .fecha = pQuery.Fields(.ddl.fecha)
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .situacionIVA = pQuery.Fields(.ddl.situacionIVA)
        .anulado = pQuery.Fields(.ddl.anulado)
        .total = pQuery.Fields(.ddl.total)
        .puntoVta = pQuery.Fields(.ddl.puntoVta)
        .nroComprob = pQuery.Fields(.ddl.nroComprob)
        .ivaCF = pQuery.Fields(.ddl.ivaCF)
        .ivaRI = pQuery.Fields(.ddl.ivaRI)
        .ivaRN = pQuery.Fields(.ddl.ivaRN)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.serieID & " = " & vSerieID
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

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.serieID & " = " & vSerieID
    sql = sql & " ORDER BY " & ddl.numero & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLastLast(db As clsDB)
Dim sql As String

    clean
    
    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.serieID & " DESC, " & ddl.numero & " DESC"
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
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.situacionIVA & " = " & vSituacionIva
    sql = sql & ", " & ddl.anulado & " = " & vAnulado
    sql = sql & ", " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
    sql = sql & ", " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & ", " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & ", " & ddl.ivaCF & " = '" & Replace(vIvaCF, ",", ".") & "'"
    sql = sql & ", " & ddl.ivaRI & " = '" & Replace(vIvaRI, ",", ".") & "'"
    sql = sql & ", " & ddl.ivaRN & " = '" & Replace(vIvaRN, ",", ".") & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.serieID & " = " & vSerieID
    sql = sql & ", " & ddl.numero & " = " & vNumero
    sql = sql & ";"

    add = db.execute(sql)

    findByPrimaryKey_ db

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
    If vClienteID <> clon.clienteID Or IsNull(vClienteID) <> IsNull(clon.clienteID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clienteID & " = " & vClienteID
        blnChange = True
    End If
    If vSituacionIva <> clon.situacionIVA Or IsNull(vSituacionIva) <> IsNull(clon.situacionIVA) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.situacionIVA & " = " & vSituacionIva
        blnChange = True
    End If
    If vAnulado <> clon.anulado Or IsNull(vAnulado) <> IsNull(clon.anulado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.anulado & " = " & vAnulado
        blnChange = True
    End If
    If vTotal <> clon.total Or IsNull(vTotal) <> IsNull(clon.total) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
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
    If vIvaCF <> clon.ivaCF Or IsNull(vIvaCF) <> IsNull(clon.ivaCF) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ivaCF & " = '" & Replace(vIvaCF, ",", ".") & "'"
        blnChange = True
    End If
    If vIvaRI <> clon.ivaRI Or IsNull(vIvaRI) <> IsNull(clon.ivaRI) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ivaRI & " = '" & Replace(vIvaRI, ",", ".") & "'"
        blnChange = True
    End If
    If vIvaRN <> clon.ivaRN Or IsNull(vIvaRN) <> IsNull(clon.ivaRN) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ivaRN & " = '" & Replace(vIvaRN, ",", ".") & "'"
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
    sql = sql & " WHERE " & ddl.serieID & " = " & vSerieID
    sql = sql & " ORDER BY " & ddl.numero & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vNumero = 1 + rstQuery.Fields(ddl.numero)

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

Public Function collectionByLiquidacion(puntoVta As Integer, nroComprob As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVta & " = " & puntoVta
    sql = sql & " AND " & ddl.nroComprob & " = " & nroComprob
    sql = sql & " AND " & ddl.anulado & " = 0"
    sql = sql & ";"

    Set collectionByLiquidacion = collectionSQL(sql, db)

End Function

Public Function collectionActivasByClienteID(clienteID As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " AND " & ddl.anulado & " = 0"
    sql = sql & ";"

    Set collectionActivasByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionByPago(pFecha As Date, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " INNER JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.fecha_t & " = " & fechaDB(pFecha)
    sql = sql & " AND " & ddl.anulado_t & " = 0"
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & c.apellido_t & ", " & c.nombre_t & ", " & ddl.clienteID_t
    sql = sql & ";"

    Set collectionByPago = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyANCredito

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyANCredito
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
Dim elemento As clsMyANCredito

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
Dim elemento As clsMyANCredito

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

