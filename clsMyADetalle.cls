VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyADetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPuntoVta As Integer
Private vNroComprob As Long
Private vRubroID As Integer
Private vConcepto As String
Private vCantidad As Currency
Private vPrecioUnitario As Currency
Private vIva As Integer
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyADetalle

Public ddl As New clsMyDDetalle

Private Sub class_Initialize()

    vPuntoVta = 0
    vNroComprob = 0
    vRubroID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vConcepto = ""
    vCantidad = 0
    vPrecioUnitario = 0
    vIva = 0
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

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

Public Property Get rubroID() As Variant

    rubroID = vRubroID

End Property

Public Property Let rubroID(ByVal vNewValue As Variant)

On Error Resume Next

    vRubroID = vNewValue

End Property

Public Property Get concepto() As Variant

    concepto = vConcepto

End Property

Public Property Let concepto(ByVal vNewValue As Variant)

On Error Resume Next

    vConcepto = vNewValue

End Property

Public Property Get cantidad() As Variant

    cantidad = vCantidad

End Property

Public Property Let cantidad(ByVal vNewValue As Variant)

On Error Resume Next

    vCantidad = vNewValue

End Property

Public Property Get precioUnitario() As Variant

    precioUnitario = vPrecioUnitario

End Property

Public Property Let precioUnitario(ByVal vNewValue As Variant)

On Error Resume Next

    vPrecioUnitario = vNewValue

End Property

Public Property Get IVA() As Variant

    IVA = vIva

End Property

Public Property Let IVA(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

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

    comboText = vConcepto

End Property

Public Property Get comboKey() As Variant

    comboKey = vRubroID

End Property

Public Property Get listText() As Variant

    listText = vConcepto

End Property

Public Property Get listKey() As Variant

    listKey = vRubroID

End Property

Public Function clone() As clsMyADetalle
Dim objObjeto As New clsMyADetalle

    With objObjeto
        .puntoVta = vPuntoVta
        .nroComprob = vNroComprob
        .rubroID = vRubroID
        .concepto = vConcepto
        .cantidad = vCantidad
        .precioUnitario = vPrecioUnitario
        .IVA = vIva
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyADetalle, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .puntoVta = pQuery.Fields(.ddl.puntoVta)
        .nroComprob = pQuery.Fields(.ddl.nroComprob)
        .rubroID = pQuery.Fields(.ddl.rubroID)
        .concepto = pQuery.Fields(.ddl.concepto)
        .cantidad = pQuery.Fields(.ddl.cantidad)
        .precioUnitario = pQuery.Fields(.ddl.precioUnitario)
        .IVA = pQuery.Fields(.ddl.IVA)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & " AND " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & " AND " & ddl.rubroID & " = " & vRubroID

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

    findLast_ IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub findLast_(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & " AND " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & " ORDER BY " & ddl.rubroID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByAutoID(db As clsDB)

    findByAutoID_ IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub findByAutoID_(db As clsDB)
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
    sql = sql & " SET " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
    sql = sql & ", " & ddl.cantidad & " = '" & Replace(vCantidad, ",", ".") & "'"
    sql = sql & ", " & ddl.precioUnitario & " = '" & Replace(vPrecioUnitario, ",", ".") & "'"
    sql = sql & ", " & ddl.IVA & " = " & vIva
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & ", " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & ", " & ddl.rubroID & " = " & vRubroID
    sql = sql & ";"

    add_ = db.execute(sql)

    findByPrimaryKey_ db

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

    If vConcepto <> clon.concepto Or IsNull(vConcepto) <> IsNull(clon.concepto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
        blnChange = True
    End If
    If vCantidad <> clon.cantidad Or IsNull(vCantidad) <> IsNull(clon.cantidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cantidad & " = '" & Replace(vCantidad, ",", ".") & "'"
        blnChange = True
    End If
    If vPrecioUnitario <> clon.precioUnitario Or IsNull(vPrecioUnitario) <> IsNull(clon.precioUnitario) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.precioUnitario & " = '" & Replace(vPrecioUnitario, ",", ".") & "'"
        blnChange = True
    End If
    If vIva <> clon.IVA Or IsNull(vIva) <> IsNull(clon.IVA) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.IVA & " = " & vIva
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

    vRubroID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & " AND " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & " ORDER BY " & ddl.rubroID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vRubroID = 1 + rstQuery.Fields(ddl.rubroID)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionByRubroID(rubroID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.rubroID & " = " & rubroID
    sql = sql & ";"

    Set collectionByRubroID = collectionSQL(sql, db)

End Function

Public Function collectionByLiquidacion(pPuntoVta As Integer, pnroComprob As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.puntoVta & " = " & pPuntoVta
    sql = sql & " AND " & ddl.nroComprob & " = " & pnroComprob
    sql = sql & " ORDER BY " & ddl.rubroID
    sql = sql & ";"

    Set collectionByLiquidacion = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyADetalle

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyADetalle
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
Dim elemento As clsMyADetalle

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
Dim elemento As clsMyADetalle

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

