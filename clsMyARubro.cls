VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyARubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRubroID As Integer
Private vFecha As Date
Private vRangoID As Integer
Private vConcepto As String
Private vPrecioUnitario As Currency
Private vIva As Integer
Private vComun As Integer
Private vComunSocio As Integer
Private vCobro As Integer
Private vDesconectado As Integer
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyARubro

Public ddl As New clsMyDRubro

Private Sub class_Initialize()

    vRubroID = 0
    vFecha = Date

    initializeOther

End Sub

Private Sub initializeOther()

    vRangoID = 0
    vConcepto = ""
    vPrecioUnitario = 0
    vIva = 0
    vComun = 0
    vComunSocio = 0
    vCobro = 0
    vDesconectado = 0
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get rubroID() As Variant

    rubroID = vRubroID

End Property

Public Property Let rubroID(ByVal vNewValue As Variant)

On Error Resume Next

    vRubroID = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get rangoID() As Variant

    rangoID = vRangoID

End Property

Public Property Let rangoID(ByVal vNewValue As Variant)

On Error Resume Next

    vRangoID = vNewValue

End Property

Public Property Get concepto() As Variant

    concepto = vConcepto

End Property

Public Property Let concepto(ByVal vNewValue As Variant)

On Error Resume Next

    vConcepto = vNewValue

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

Public Property Get comun() As Variant

    comun = vComun

End Property

Public Property Let comun(ByVal vNewValue As Variant)

On Error Resume Next

    vComun = vNewValue

End Property

Public Property Get comunSocio() As Variant

    comunSocio = vComunSocio

End Property

Public Property Let comunSocio(ByVal vNewValue As Variant)

On Error Resume Next

    vComunSocio = vNewValue

End Property

Public Property Get cobro() As Variant

    cobro = vCobro

End Property

Public Property Let cobro(ByVal vNewValue As Variant)

On Error Resume Next

    vCobro = vNewValue

End Property

Public Property Get desconectado() As Variant

    desconectado = vDesconectado

End Property

Public Property Let desconectado(ByVal vNewValue As Variant)

On Error Resume Next

    vDesconectado = vNewValue

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

    comboText = vRangoID

End Property

Public Property Get comboKey() As Variant

    comboKey = vFecha

End Property

Public Property Get listText() As Variant

    listText = vRangoID

End Property

Public Property Get listKey() As Variant

    listKey = vFecha

End Property

Public Function clone() As clsMyARubro
Dim objObjeto As New clsMyARubro

    With objObjeto
        .rubroID = vRubroID
        .fecha = vFecha
        .rangoID = vRangoID
        .concepto = vConcepto
        .precioUnitario = vPrecioUnitario
        .IVA = vIva
        .comun = vComun
        .comunSocio = vComunSocio
        .cobro = vCobro
        .desconectado = vDesconectado
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyARubro, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .rubroID = pQuery.Fields(.ddl.rubroID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .rangoID = pQuery.Fields(.ddl.rangoID)
        .concepto = pQuery.Fields(.ddl.concepto)
        .precioUnitario = pQuery.Fields(.ddl.precioUnitario)
        .IVA = pQuery.Fields(.ddl.IVA)
        .comun = pQuery.Fields(.ddl.comun)
        .comunSocio = pQuery.Fields(.ddl.comunSocio)
        .cobro = pQuery.Fields(.ddl.cobro)
        .desconectado = pQuery.Fields(.ddl.desconectado)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.rubroID & " = " & vRubroID
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
    sql = sql & " WHERE " & ddl.rubroID & " = " & vRubroID
    sql = sql & " ORDER BY " & ddl.fecha & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findFirst(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.rubroID & " = " & vRubroID
    sql = sql & " ORDER BY " & ddl.fecha
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByFecha(fecha As Date, db As clsDB)
Dim sql As String

Dim rubroID As Integer

    rubroID = vRubroID
    
    clean
    
    vRubroID = rubroID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.rubroID & " = " & vRubroID
    sql = sql & " AND " & ddl.fecha & " <= " & fechaDB(fecha)
    sql = sql & " ORDER BY " & ddl.fecha & " DESC"
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
    sql = sql & " SET " & ddl.rangoID & " = " & vRangoID
    sql = sql & ", " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
    sql = sql & ", " & ddl.precioUnitario & " = '" & Replace(vPrecioUnitario, ",", ".") & "'"
    sql = sql & ", " & ddl.IVA & " = " & vIva
    sql = sql & ", " & ddl.comun & " = " & vComun
    sql = sql & ", " & ddl.comunSocio & " = " & vComunSocio
    sql = sql & ", " & ddl.cobro & " = " & vCobro
    sql = sql & ", " & ddl.desconectado & " = " & vDesconectado
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.rubroID & " = " & vRubroID
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

    If vRangoID <> clon.rangoID Or IsNull(vRangoID) <> IsNull(clon.rangoID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.rangoID & " = " & vRangoID
        blnChange = True
    End If
    If vConcepto <> clon.concepto Or IsNull(vConcepto) <> IsNull(clon.concepto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
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
    If vComun <> clon.comun Or IsNull(vComun) <> IsNull(clon.comun) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.comun & " = " & vComun
        blnChange = True
    End If
    If vComunSocio <> clon.comunSocio Or IsNull(vComunSocio) <> IsNull(clon.comunSocio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.comunSocio & " = " & vComunSocio
        blnChange = True
    End If
    If vCobro <> clon.cobro Or IsNull(vCobro) <> IsNull(clon.cobro) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cobro & " = " & vCobro
        blnChange = True
    End If
    If vDesconectado <> clon.desconectado Or IsNull(vDesconectado) <> IsNull(clon.desconectado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.desconectado & " = " & vDesconectado
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
    sql = sql & " WHERE " & ddl.rubroID & " = " & vRubroID
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

Public Function collectionAny(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set collectionAny = collectionSQL(sql, db)

End Function

Public Function collectionSinRepeticion(db As clsDB) As Collection
Dim colUnicos As New Collection

Dim objRub As New clsMyARubro
Dim objRubAnt As clsMyARubro

Dim rubroID As Integer

Dim blnPrimero As Boolean

    rubroID = 1
    blnPrimero = True
    For Each objRub In objRub.collectionAll(db)
        If blnPrimero Then
            Set objRubAnt = objRub.clone
            blnPrimero = False
        End If
        If rubroID < objRub.rubroID Then
            colUnicos.add objRubAnt
            rubroID = objRub.rubroID
        End If
        Set objRubAnt = objRub.clone
    Next
    colUnicos.add objRubAnt
    
    Set collectionSinRepeticion = colUnicos

End Function

Public Function collectionByRubroID(rubroID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.rubroID & " = " & rubroID
    sql = sql & " ORDER BY " & ddl.fecha
    sql = sql & ";"

    Set collectionByRubroID = collectionSQL(sql, db)

End Function

Public Function collectionByRangoID(rangoID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.rangoID & " = " & rangoID
    sql = sql & ";"

    Set collectionByRangoID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyARubro

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyARubro
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
Dim elemento As clsMyARubro

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
Dim elemento As clsMyARubro

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

