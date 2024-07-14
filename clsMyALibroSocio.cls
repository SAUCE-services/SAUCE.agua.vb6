VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyALibroSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vNumeroSocio As Integer
Private vAnho As Integer
Private vNombreApellido As String
Private vDomicilio As String
Private vDocumento As Currency
Private vEstado As String
Private vEdad As Integer
Private vCategoria As String
Private vIngreso As Date
Private vEnero As Variant
Private vFebrero As Variant
Private vMarzo As Variant
Private vAbril As Variant
Private vMayo As Variant
Private vJunio As Variant
Private vJulio As Variant
Private vAgosto As Variant
Private vSetiembre As Variant
Private vOctubre As Variant
Private vNoviembre As Variant
Private vDiciembre As Variant
Private vAutoID As Long
Private vCreated As Date

Private clon As clsMyALibroSocio

Public ddl As New clsMyDLibroSocio

Private Sub class_Initialize()

    vNumeroSocio = 0
    vAnho = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNombreApellido = ""
    vDomicilio = ""
    vDocumento = 0
    vEstado = ""
    vEdad = 0
    vCategoria = ""
    vIngreso = Date
    vEnero = Null
    vFebrero = Null
    vMarzo = Null
    vAbril = Null
    vMayo = Null
    vJunio = Null
    vJulio = Null
    vAgosto = Null
    vSetiembre = Null
    vOctubre = Null
    vNoviembre = Null
    vDiciembre = Null
    vAutoID = 0
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get numeroSocio() As Variant

    numeroSocio = vNumeroSocio

End Property

Public Property Let numeroSocio(ByVal vNewValue As Variant)

On Error Resume Next

    vNumeroSocio = vNewValue

End Property

Public Property Get anho() As Variant

    anho = vAnho

End Property

Public Property Let anho(ByVal vNewValue As Variant)

On Error Resume Next

    vAnho = vNewValue

End Property

Public Property Get nombreApellido() As Variant

    nombreApellido = vNombreApellido

End Property

Public Property Let nombreApellido(ByVal vNewValue As Variant)

On Error Resume Next

    vNombreApellido = vNewValue

End Property

Public Property Get domicilio() As Variant

    domicilio = vDomicilio

End Property

Public Property Let domicilio(ByVal vNewValue As Variant)

On Error Resume Next

    vDomicilio = vNewValue

End Property

Public Property Get documento() As Variant

    documento = vDocumento

End Property

Public Property Let documento(ByVal vNewValue As Variant)

On Error Resume Next

    vDocumento = vNewValue

End Property

Public Property Get estado() As Variant

    estado = vEstado

End Property

Public Property Let estado(ByVal vNewValue As Variant)

On Error Resume Next

    vEstado = vNewValue

End Property

Public Property Get edad() As Variant

    edad = vEdad

End Property

Public Property Let edad(ByVal vNewValue As Variant)

On Error Resume Next

    vEdad = vNewValue

End Property

Public Property Get categoria() As Variant

    categoria = vCategoria

End Property

Public Property Let categoria(ByVal vNewValue As Variant)

On Error Resume Next

    vCategoria = vNewValue

End Property

Public Property Get ingreso() As Variant

    ingreso = vIngreso

End Property

Public Property Let ingreso(ByVal vNewValue As Variant)

On Error Resume Next

    vIngreso = vNewValue

End Property

Public Property Get enero() As Variant

    enero = vEnero

End Property

Public Property Let enero(ByVal vNewValue As Variant)

On Error Resume Next

    vEnero = vNewValue

End Property

Public Property Get febrero() As Variant

    febrero = vFebrero

End Property

Public Property Let febrero(ByVal vNewValue As Variant)

On Error Resume Next

    vFebrero = vNewValue

End Property

Public Property Get marzo() As Variant

    marzo = vMarzo

End Property

Public Property Let marzo(ByVal vNewValue As Variant)

On Error Resume Next

    vMarzo = vNewValue

End Property

Public Property Get abril() As Variant

    abril = vAbril

End Property

Public Property Let abril(ByVal vNewValue As Variant)

On Error Resume Next

    vAbril = vNewValue

End Property

Public Property Get mayo() As Variant

    mayo = vMayo

End Property

Public Property Let mayo(ByVal vNewValue As Variant)

On Error Resume Next

    vMayo = vNewValue

End Property

Public Property Get junio() As Variant

    junio = vJunio

End Property

Public Property Let junio(ByVal vNewValue As Variant)

On Error Resume Next

    vJunio = vNewValue

End Property

Public Property Get julio() As Variant

    julio = vJulio

End Property

Public Property Let julio(ByVal vNewValue As Variant)

On Error Resume Next

    vJulio = vNewValue

End Property

Public Property Get agosto() As Variant

    agosto = vAgosto

End Property

Public Property Let agosto(ByVal vNewValue As Variant)

On Error Resume Next

    vAgosto = vNewValue

End Property

Public Property Get setiembre() As Variant

    setiembre = vSetiembre

End Property

Public Property Let setiembre(ByVal vNewValue As Variant)

On Error Resume Next

    vSetiembre = vNewValue

End Property

Public Property Get octubre() As Variant

    octubre = vOctubre

End Property

Public Property Let octubre(ByVal vNewValue As Variant)

On Error Resume Next

    vOctubre = vNewValue

End Property

Public Property Get noviembre() As Variant

    noviembre = vNoviembre

End Property

Public Property Let noviembre(ByVal vNewValue As Variant)

On Error Resume Next

    vNoviembre = vNewValue

End Property

Public Property Get diciembre() As Variant

    diciembre = vDiciembre

End Property

Public Property Let diciembre(ByVal vNewValue As Variant)

On Error Resume Next

    vDiciembre = vNewValue

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

    comboText = vNombreApellido

End Property

Public Property Get comboKey() As Variant

    comboKey = vAnho

End Property

Public Property Get listText() As Variant

    listText = vNombreApellido

End Property

Public Property Get listKey() As Variant

    listKey = vAnho

End Property

Public Function clone() As clsMyALibroSocio
Dim objObjeto As New clsMyALibroSocio

    With objObjeto
        .numeroSocio = vNumeroSocio
        .anho = vAnho
        .nombreApellido = vNombreApellido
        .domicilio = vDomicilio
        .documento = vDocumento
        .estado = vEstado
        .edad = vEdad
        .categoria = vCategoria
        .ingreso = vIngreso
        .enero = vEnero
        .febrero = vFebrero
        .marzo = vMarzo
        .abril = vAbril
        .mayo = vMayo
        .junio = vJunio
        .julio = vJulio
        .agosto = vAgosto
        .setiembre = vSetiembre
        .octubre = vOctubre
        .noviembre = vNoviembre
        .diciembre = vDiciembre
        .autoID = vAutoID
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyALibroSocio, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .numeroSocio = pQuery.Fields(.ddl.numeroSocio)
        .anho = pQuery.Fields(.ddl.anho)
        .nombreApellido = pQuery.Fields(.ddl.nombreApellido)
        .domicilio = pQuery.Fields(.ddl.domicilio)
        .documento = pQuery.Fields(.ddl.documento)
        .estado = pQuery.Fields(.ddl.estado)
        .edad = pQuery.Fields(.ddl.edad)
        .categoria = pQuery.Fields(.ddl.categoria)
        .ingreso = pQuery.Fields(.ddl.ingreso)
        .enero = pQuery.Fields(.ddl.enero)
        .febrero = pQuery.Fields(.ddl.febrero)
        .marzo = pQuery.Fields(.ddl.marzo)
        .abril = pQuery.Fields(.ddl.abril)
        .mayo = pQuery.Fields(.ddl.mayo)
        .junio = pQuery.Fields(.ddl.junio)
        .julio = pQuery.Fields(.ddl.julio)
        .agosto = pQuery.Fields(.ddl.agosto)
        .setiembre = pQuery.Fields(.ddl.setiembre)
        .octubre = pQuery.Fields(.ddl.octubre)
        .noviembre = pQuery.Fields(.ddl.noviembre)
        .diciembre = pQuery.Fields(.ddl.diciembre)
        .autoID = pQuery.Fields(.ddl.autoID)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.numeroSocio & " = " & vNumeroSocio
    sql = sql & " AND " & ddl.anho & " = " & vAnho

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
    sql = sql & " WHERE " & ddl.numeroSocio & " = " & vNumeroSocio
    sql = sql & " ORDER BY " & ddl.anho & " DESC"
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
    sql = sql & " SET " & ddl.nombreApellido & " = '" & replaceSpecialSymbols(vNombreApellido) & "'"
    sql = sql & ", " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
    sql = sql & ", " & ddl.documento & " = '" & Replace(vDocumento, ",", ".") & "'"
    sql = sql & ", " & ddl.estado & " = '" & replaceSpecialSymbols(vEstado) & "'"
    sql = sql & ", " & ddl.edad & " = " & vEdad
    sql = sql & ", " & ddl.categoria & " = '" & replaceSpecialSymbols(vCategoria) & "'"
    sql = sql & ", " & ddl.ingreso & " = " & fechaDB(vIngreso)
    sql = sql & ", " & ddl.enero & " = " & fechaDB(vEnero)
    sql = sql & ", " & ddl.febrero & " = " & fechaDB(vFebrero)
    sql = sql & ", " & ddl.marzo & " = " & fechaDB(vMarzo)
    sql = sql & ", " & ddl.abril & " = " & fechaDB(vAbril)
    sql = sql & ", " & ddl.mayo & " = " & fechaDB(vMayo)
    sql = sql & ", " & ddl.junio & " = " & fechaDB(vJunio)
    sql = sql & ", " & ddl.julio & " = " & fechaDB(vJulio)
    sql = sql & ", " & ddl.agosto & " = " & fechaDB(vAgosto)
    sql = sql & ", " & ddl.setiembre & " = " & fechaDB(vSetiembre)
    sql = sql & ", " & ddl.octubre & " = " & fechaDB(vOctubre)
    sql = sql & ", " & ddl.noviembre & " = " & fechaDB(vNoviembre)
    sql = sql & ", " & ddl.diciembre & " = " & fechaDB(vDiciembre)
    sql = sql & ", " & ddl.numeroSocio & " = " & vNumeroSocio
    sql = sql & ", " & ddl.anho & " = " & vAnho
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

    If vNombreApellido <> clon.nombreApellido Or IsNull(vNombreApellido) <> IsNull(clon.nombreApellido) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nombreApellido & " = '" & replaceSpecialSymbols(vNombreApellido) & "'"
        blnChange = True
    End If
    If vDomicilio <> clon.domicilio Or IsNull(vDomicilio) <> IsNull(clon.domicilio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
        blnChange = True
    End If
    If vDocumento <> clon.documento Or IsNull(vDocumento) <> IsNull(clon.documento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.documento & " = '" & Replace(vDocumento, ",", ".") & "'"
        blnChange = True
    End If
    If vEstado <> clon.estado Or IsNull(vEstado) <> IsNull(clon.estado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.estado & " = '" & replaceSpecialSymbols(vEstado) & "'"
        blnChange = True
    End If
    If vEdad <> clon.edad Or IsNull(vEdad) <> IsNull(clon.edad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.edad & " = " & vEdad
        blnChange = True
    End If
    If vCategoria <> clon.categoria Or IsNull(vCategoria) <> IsNull(clon.categoria) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.categoria & " = '" & replaceSpecialSymbols(vCategoria) & "'"
        blnChange = True
    End If
    If vIngreso <> clon.ingreso Or IsNull(vIngreso) <> IsNull(clon.ingreso) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ingreso & " = " & fechaDB(vIngreso)
        blnChange = True
    End If
    If vEnero <> clon.enero Or IsNull(vEnero) <> IsNull(clon.enero) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.enero & " = " & fechaDB(vEnero)
        blnChange = True
    End If
    If vFebrero <> clon.febrero Or IsNull(vFebrero) <> IsNull(clon.febrero) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.febrero & " = " & fechaDB(vFebrero)
        blnChange = True
    End If
    If vMarzo <> clon.marzo Or IsNull(vMarzo) <> IsNull(clon.marzo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.marzo & " = " & fechaDB(vMarzo)
        blnChange = True
    End If
    If vAbril <> clon.abril Or IsNull(vAbril) <> IsNull(clon.abril) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.abril & " = " & fechaDB(vAbril)
        blnChange = True
    End If
    If vMayo <> clon.mayo Or IsNull(vMayo) <> IsNull(clon.mayo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.mayo & " = " & fechaDB(vMayo)
        blnChange = True
    End If
    If vJunio <> clon.junio Or IsNull(vJunio) <> IsNull(clon.junio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.junio & " = " & fechaDB(vJunio)
        blnChange = True
    End If
    If vJulio <> clon.julio Or IsNull(vJulio) <> IsNull(clon.julio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.julio & " = " & fechaDB(vJulio)
        blnChange = True
    End If
    If vAgosto <> clon.agosto Or IsNull(vAgosto) <> IsNull(clon.agosto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.agosto & " = " & fechaDB(vAgosto)
        blnChange = True
    End If
    If vSetiembre <> clon.setiembre Or IsNull(vSetiembre) <> IsNull(clon.setiembre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.setiembre & " = " & fechaDB(vSetiembre)
        blnChange = True
    End If
    If vOctubre <> clon.octubre Or IsNull(vOctubre) <> IsNull(clon.octubre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.octubre & " = " & fechaDB(vOctubre)
        blnChange = True
    End If
    If vNoviembre <> clon.noviembre Or IsNull(vNoviembre) <> IsNull(clon.noviembre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.noviembre & " = " & fechaDB(vNoviembre)
        blnChange = True
    End If
    If vDiciembre <> clon.diciembre Or IsNull(vDiciembre) <> IsNull(clon.diciembre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.diciembre & " = " & fechaDB(vDiciembre)
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

    vAnho = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.numeroSocio & " = " & vNumeroSocio
    sql = sql & " ORDER BY " & ddl.anho & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vAnho = 1 + rstQuery.Fields(ddl.anho)

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

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyALibroSocio

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyALibroSocio
        fillObject elemento, rstQuery

        registros.add elemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillComboCollection pCombo, collectionAll(db)

End Sub

Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyALibroSocio

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, db As clsDB)

    fillListCollection pList, collectionAll(db)

End Sub

Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyALibroSocio

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

