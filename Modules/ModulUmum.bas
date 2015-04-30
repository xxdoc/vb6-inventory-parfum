Attribute VB_Name = "modulGencil"
'=========================================================
'           Module Pemrograman Database (CRUD)
'               Author : Eric Ariyanto
'                   @ericariyanto
'       Anda bebas menggunakan module ini, tapi mohon
'       tidak menghapus creditnya. Thanks
'              (c) 2013 - Yogyakarta
'=========================================================



Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal h As Long, ByVal lpoOP As String, ByVal lpFile As String, ByVal lpParam As String, ByVal lpDir As String, ByVal nShowCmd As Long) As Long


Public Conn As New ADODB.Connection 'var utk koneksi

Option Explicit
Dim eachField As Control
Dim button As Control

Public Strconn As String 'utk lokasi database
Public sql As String 'utk meyiimpan query sql
Public Rs As New ADODB.Recordset 'utk membaca isi table

Public appTitlte As String

Public usrID As String
Public usrName As String
Public usrLevel As String

Sub main()
    usrID = "1"
    usrName = "admin"
    usrLevel = "1"
    'frmSplash.Show
    FrmLaporan.Show
End Sub

'sub untuk menghubungkan dengan database
Sub Koneksi()
On Error GoTo error_koneksi
    appTitlte = "Aplikasi Inventory Rumah Parfum"
    
    'Strconn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=penggajian"
'    Strconn = "DRIVER={MYSQL ODBC 5.3w Driver};" _
'        & "SERVER=localhost;" _
'        & "DATABASE=gencil_perpus_bopkri;" _
'        & "UID=root;" _
'        & "PWD=910904;" _
'        & "OPTION=3"

'    Strconn = "DRIVER={MYSQL ODBC 5.3 ANSI Driver};" _
'        & "SERVER=localhost;" _
'        & "DATABASE=gencil_perpus_bopkri;" _
'        & "User=root;" _
'        & "Password=910904;" _
'        & "OPTION=3"
    Strconn = "Provider=MSDASQL.1;Password=910904;Persist Security Info=True;User ID=root;Data Source=ODBC 35;Initial Catalog=gencil_parfum"

    'Conn.CursorLocation = adUseClient
    'If Conn.State = adStateOpen Then
    'Conn.Close
    Set Conn = New ADODB.Connection
    'End If
    Conn.Open (Strconn)
Exit Sub
error_koneksi:
Beep
MsgBox "Tidak dapat terkoneksi ke database", vbCritical, "PERHATIAN"
End
End Sub

Sub ordering_listview(lv As ListView, Optional intStart As Integer = 1, Optional intKolom As Integer = 1)
    Dim i As Integer
    Dim no As Integer
    no = intStart
    For i = 1 To lv.ListItems.Count
        lv.ListItems(i).SubItems(intKolom) = no
        no = no + 1
    Next
End Sub

Function is_lv_item_valid(lv As ListView, strCari As String, Optional intCari As Integer = 3)
is_lv_item_valid = True
    Dim i As Integer
    For i = 1 To lv.ListItems.Count
        If lv.ListItems(i).SubItems(intCari) = strCari Then
            is_lv_item_valid = False
            Exit Function
        End If
    Next
End Function

Function get_total_data(strTbl As String, Optional strWhere As String = "")
    sql = "SELECT COUNT(*) FROM " & strTbl & " " & strWhere
    Set Rs = Conn.Execute(sql)
    get_total_data = Rs.Fields(0)
End Function

Function pagingValid(cbo As ComboBox, nilai As Integer, strTable As String, Optional strWhere As String = "")
    Dim total As Integer
    Dim perpage As Integer
        perpage = Val(cbo.Text)
    
    sql = "SELECT COUNT(*) FROM " & strTable & " " & strWhere
    Set Rs = Conn.Execute(sql)
    total = Rs.Fields(0)
    
    If (nilai * perpage) > (total + perpage) Then
        pagingValid = False
    Else
        pagingValid = True
    End If
End Function

Sub IMK_Paging(cbo As ComboBox, fr As Frame)
    If LCase(cbo.Text) = "semua" Then
        fr.Visible = False
    Else
        fr.Visible = True
    End If
End Sub

Sub setPerPage(cbo As ComboBox, txt As TextBox)
    cbo.Clear
    cbo.AddItem "10", 0
    cbo.AddItem "20", 1
    cbo.AddItem "50", 2
    cbo.AddItem "100", 3
    cbo.AddItem "Semua", 4
    cbo.ListIndex = 0
    cbo.Enabled = True
    
    txt.Text = "1"
    txt.Enabled = True
End Sub


Sub getFocused(txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt)
End Sub

Sub clearAllText(a As Form)
    On Error Resume Next
    For Each eachField In a.Controls
        If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Then
            eachField.Text = ""
        End If
    Next
End Sub

Sub enableAllText(nil As Boolean, a As Form)
    On Error Resume Next
    For Each eachField In a.Controls
        If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Or TypeOf eachField Is DTPicker Then
            eachField.Enabled = nil
        End If
    Next
End Sub

Sub tombol(cmdT As CommandButton, cmdS As CommandButton, _
        cmdU As CommandButton, cmdD As CommandButton, cmdB As CommandButton, _
        Optional tbh As Boolean = True, Optional spn As Boolean = False, _
        Optional ubh As Boolean = False, Optional del As Boolean = False, _
        Optional b As Boolean = False)
    cmdT.Enabled = tbh
    cmdS.Enabled = spn
    cmdU.Enabled = ubh
    cmdD.Enabled = del
    cmdB.Enabled = b
End Sub

Function AntiSQLi(sumber As String)
Dim hasil As Variant
    hasil = Replace(sumber, "\", "\\")
    hasil = Replace(hasil, "'", "\'")
    hasil = Replace(hasil, ";", "")
    hasil = Replace(hasil, "--", "")
    AntiSQLi = hasil
End Function

Function AntiSQLiWithQuotes(sumber As String)
Dim hasil As Variant
    hasil = AntiSQLi(sumber)
    hasil = "'" & hasil & "'"
    AntiSQLiWithQuotes = hasil
End Function

Sub jmlRecord(lbl As Label, lv As ListView)
    Dim jml%
    jml = lv.ListItems.Count
    lbl.Caption = "Jumlah Record : " & jml
End Sub

Function isDuplicate(tbl$, kol$, qWhere$) As Boolean
On Error GoTo err
isDuplicate = False
    sql = "select " & kol & " from " & tbl & " WHERE " & qWhere
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        If IsNull(Rs(0)) = False Then
            isDuplicate = True
        End If
    End If
    Rs.Close
    Exit Function
err:
End Function

Function autoNumbering(tbl As String, kol As String, _
        strLen As Integer, optFormat As String) As String
On Local Error GoTo err
Dim justNumber As Integer
justNumber = strLen - Len(optFormat)
autoNumbering = optFormat & Right(String(justNumber - 1, "0") & "1", strLen)
    sql = "select max(right(" & kol & "," & CStr(justNumber) & "))+1 from " & tbl
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        If IsNull(Rs(0)) = False Then
            'autoNumbering = optFormat & Right(String(justNumber - 1, "0") & CStr(Rs(0)), strLen)
            autoNumbering = Right(optFormat & Right(String(justNumber - 1, "0") & CStr(Rs(0)), justNumber), strLen)
        End If
    End If
    Exit Function
err:
    autoNumbering = optFormat & Right(String(justNumber - 1, "0") & "1", strLen)
End Function

'untuk menghitung panjang karakter
Function lenString(str As String) As Integer
lenString = Len(Trim(str))
End Function

Function isLogin(user As String, pass As String, Optional tbl As String = "user") As Boolean
isLogin = False
    user = AntiSQLi(user)
    pass = AntiSQLi(pass)
    
    sql = "select * from " & tbl & " WHERE " & _
          "user_nama = '" & user & "' " & _
          "AND " & _
          "user_password='" & pass & "'"
    
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        isLogin = True
        usrID = Rs(0)
        usrName = Rs(1)
        usrLevel = Rs(3)
    End If
    Rs.Close
Exit Function
err:
   isLogin = False
End Function

Public Sub sortingListView(lv As ListView, Optional intLst As Integer = 1)
    Dim i As Integer
    For i = 1 To lv.ListItems.Count
        lv.ListItems(i).SubItems(1) = intLst
        intLst = intLst + 1
    Next
End Sub

Public Sub tampilData(query As String, lv As ListView, Optional intStart As Integer = 1)
'On Error GoTo err
On Error Resume Next
Dim i As Integer
Dim j As Integer
    lv.ListItems.Clear
    Set Rs = Conn.Execute(query)
    j = intStart
    While Not Rs.EOF
        With lv.ListItems.Add
            .Text = ""
            .SubItems(1) = j
            For i = 1 To Rs.Fields.Count
                .SubItems(i + 1) = IIf(IsNull(Rs(i - 1)), "", Rs(i - 1))
            Next
        End With
        Rs.MoveNext
        j = j + 1
    Wend
    Rs.Close
    
    If lv.ListItems.Count > 0 Then
        'default sort
        lv.SortOrder = lvwAscending
        lv.SortKey = 0
        'lv.ColumnHeaders(1).Icon = 1
        lv.ListItems(1).Selected = True
    End If
Exit Sub
err:
    MsgBox "Terjadi kesalahan, data tidak dapat diload sempurna", vbExclamation, "Warning"
End Sub


Public Function saveData(tbl As String, arr() As String) As Boolean
On Error GoTo hell
Dim i As Integer
Dim c As Integer
saveData = True
    sql = "insert into " & tbl & " values ("
    c = UBound(arr)
    For i = 0 To c
        If arr(i) = "null" Then
            sql = sql & "null"
            If i < (c) Then sql = sql & ","
        ElseIf arr(i) = "" Then
            sql = sql & "'" & arr(i) & "'"
            If i < (c) Then sql = sql & ","
        ElseIf arr(i) <> "" Then
            sql = sql & "'" & arr(i) & "'"
            If i < (c) Then sql = sql & ","
        End If
    Next
    sql = sql & ")"
    Conn.Execute sql
    Exit Function
hell:
    saveData = False
End Function

'not support null
'Public Function updateData(tbl As String, col() As String, _
'        arr() As String, strWhere As String) As Boolean
'On Error GoTo hell
'Dim i As Integer
'Dim c As Integer
'updateData = True
'    sql = "update " & tbl & " set "
'    c = UBound(arr)
'    For i = 0 To c
'        If arr(i) <> "" Then
'            sql = sql & col(i) & "='" & arr(i) & "'"
'            If i < (c) Then sql = sql & ","
'        End If
'    Next
'    sql = sql & " Where " & strWhere
'    Conn.Execute sql
'    Exit Function
'hell:
'    updateData = False
'End Function


'support null
Public Function updateData(tbl As String, col() As String, _
        arr() As String, strWhere As String) As Boolean
On Error GoTo hell
Dim i As Integer
Dim c As Integer
updateData = True
    sql = "update " & tbl & " set "
    c = UBound(arr)
    For i = 0 To c
        If arr(i) = "_NOT_CHANGE_" Then
        ElseIf arr(i) = "null" Then
            sql = sql & col(i) & "=" & arr(i)
            If i < (c) Then sql = sql & ","
        ElseIf arr(i) = "" Then
            sql = sql & col(i) & "='" & arr(i) & "'"
            If i < (c) Then sql = sql & ","
        ElseIf arr(i) <> "" Then
            sql = sql & col(i) & "='" & arr(i) & "'"
            If i < (c) Then sql = sql & ","
        End If
    Next
    sql = sql & " Where " & strWhere
    Conn.Execute sql
    Exit Function
hell:
    updateData = False
End Function


Public Function delData(tbl As String, strWhere As String) As Boolean
On Error GoTo hell
delData = True
    sql = "Delete from " & tbl
    sql = sql & " WHERE " & strWhere
    Conn.Execute sql
    Exit Function
hell:
    delData = False
End Function

Public Function strNota(tbl As String, kol As String, _
    tglnya As Date, Optional tgl As String = "tgl_beli") As String
Dim yy As String
Dim mm As String
Dim dd As String
    yy = Right(Year(tglnya), 2)
    mm = Format(tglnya, "mm")
    dd = Format(tglnya, "dd")
strNota = yy & mm & dd & "0001"
On Error GoTo hell
    sql = "select max(right(" & kol & ",4))+1 from " & tbl & _
          " WHERE day(" & tgl & ") = '" & Format(tglnya, "dd") & "' " & _
          "AND month(" & tgl & ")='" & Format(tglnya, "mm") & "' " & _
          "AND year(" & tgl & ")='" & Year(tglnya) & "'"
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        If IsNull(Rs(0)) = False Then
            strNota = yy & mm & dd & Right("000" & CStr(Rs(0)), 4)
        End If
    End If
    Rs.Close
Exit Function
hell:
strNota = yy & mm & dd & "01"
End Function

Function set_tglHarusKembali(tgl As Date, _
    Optional tipe$ = "Biasa", Optional intHari As Integer = 7) As Date
set_tglHarusKembali = tgl
On Error GoTo err
    If tipe = "Biasa" Then
        set_tglHarusKembali = DateAdd("d", intHari, tgl)
    Else
        set_tglHarusKembali = DateAdd("YYYY", 1, tgl)
    End If
Exit Function
err:
End Function

Public Function intForIR(nil As Variant) As String
    intForIR = Format(nil, "##,##0")
    intForIR = Replace(intForIR, ",", ".")
End Function

Public Function intPolos(nil As String) As Variant
    intPolos = Replace(nil, ".", "")
    intPolos = Replace(intPolos, ",", "")
End Function

Public Function getStok(kode As String) As Integer
On Error GoTo err
getStok = 0
    sql = "select stok from katalog where id_katalog='" & AntiSQLi(kode) & "'"
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        getStok = Rs(0)
    End If
    Rs.Close
    Exit Function
err:
End Function

Public Function getValues(tbl As String, kol As String, qWhere As String) As Variant
On Error GoTo err
    getValues = ""
    sql = "select " & kol & " from " & tbl & " WHere " & qWhere
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        getValues = Rs(0)
    End If
    Rs.Close
    Exit Function
err:
    
End Function
'listview warna warni
'Public Sub SetListViewLedger(fr As Form, ImageList1 As ImageList, _
'                              Picture1_WARNA_LISVIEW As PictureBox, _
'                              lv As ListView, _
'                              Bar1Color As LedgerColours, _
'                              Bar2Color As LedgerColours, _
'                              nSizingType As ImageSizingTypes)
'   Dim iBarHeight  As Long
'   Dim lBarWidth   As Long
'   Dim diff        As Long
'   Dim twipsy      As Long
'
'   iBarHeight = 0
'   lBarWidth = 0
'   diff = 0
'
'    On Local Error GoTo SetListViewColor_Error
'
'   twipsy = Screen.TwipsPerPixelY
'   If lv.View = lvwReport Then
'      With lv
'        .Picture = Nothing
'        .Refresh
'        .Visible = 1
'        .PictureAlignment = lvwTile
'        lBarWidth = .Width
'      End With  ' lv
'
'      With Picture1_WARNA_LISVIEW
'         .AutoRedraw = False
'         .Picture = Nothing
'         .BackColor = vbWhite
'         .Height = 1
'         .AutoRedraw = True
'         .BorderStyle = vbBSNone
'         .ScaleMode = vbTwips
'         .Top = fr.Top - 10000
'         .Width = Screen.Width
'         .Visible = False
'         .Font = lv.Font
'
'         With .Font
'            .Bold = lv.Font.Bold
'            .Charset = lv.Font.Charset
'            .Italic = lv.Font.Italic
'            .Name = lv.Font.Name
'            .Strikethrough = lv.Font.Strikethrough
'            .Underline = lv.Font.Underline
'            .Weight = lv.Font.Weight
'            .Size = lv.Font.Size
'         End With
'
'         iBarHeight = .TextHeight("W")
'
'         Select Case nSizingType
'            Case sizeNone:
'               iBarHeight = iBarHeight + twipsy
'            Case sizeCheckBox:
'               If (iBarHeight \ twipsy) > 18 Then
'                  iBarHeight = iBarHeight + twipsy
'               Else
'                  diff = 18 - (iBarHeight \ twipsy)
'                  iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)
'               End If
'            Case sizeIcon:
'               diff = ImageList1.ImageHeight - (iBarHeight \ twipsy)
'               iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)
'         End Select
'         .Height = iBarHeight * 2
'         .Width = lBarWidth
'         Picture1_WARNA_LISVIEW.Line (0, 0)-(lBarWidth, iBarHeight), Bar1Color, BF
'         Picture1_WARNA_LISVIEW.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), Bar2Color, BF
'         .AutoSize = True
'         .Refresh
'      End With  'Picture1
'      lv.Refresh
'      lv.Picture = Picture1_WARNA_LISVIEW.Image
'   Else
'      lv.Picture = Nothing
'   End If  'lv.View = lvwReport
'SetListViewColor_Exit:
'On Local Error GoTo 0
'Exit Sub
'SetListViewColor_Error:
'   With lv
'      .Picture = Nothing
'      .Refresh
'   End With
'   Resume SetListViewColor_Exit
'End Sub



'Public Sub deleteFile(fname As String)
'On Error GoTo err
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If FExists(fname) Then
'        fso.deleteFile (fname)
'    End If
'    Exit Sub
'err:
'    MsgBox "Gagal menghapus file", vbInformation, "informasi"
'End Sub
'
'Public Function getImage(id As String, Optional tipe As String = "siswa") As String
'getImage = ""
'    If tipe = "siswa" Then
'        sql = "select nm_panggilan from siswa where nis='" & id & "'"
'        getImage = App.Path & "\foto\siswa\"
'    Else
'        sql = "select nik from pegawai where nik='" & id & "'"
'        getImage = App.Path & "\foto\pegawai\"
'    End If
'
'    Set Rs = Conn.Execute(sql)
'    If Not Rs.EOF Then
'        getImage = getImage & Rs(0) & ".jpg"
'    End If
'End Function
'
'
'
'Public Sub openFile(fname As String)
'On Error GoTo err
'    Call ShellExecute(hWnd, "open", fname, vbNullString, vbNullString, 3)
'
'    Exit Sub
'err:
'    MsgBox "File tidak ditemukan"
'End Sub
'
'Public Sub simpanImage(lokasi As String, fname As String, _
'        Optional tipe As String = "siswa")
'On Error Resume Next
'    FileCopy lokasi, App.Path & "\foto\" & tipe & "\" & fname
'    Exit Sub
'End Sub
'
'
''untuk cek data tersebut valid atau tidak
'Function isDataValid(tbl As String, kolCheck As String, _
'        kolSyarat As String, nilai As String) As Boolean
'isDataValid = False
'On Error GoTo err
'    sql = "select " & kolCheck & " from " & tbl & _
'          " WHERE " & kolSyarat & " = '" & AntiSQLi(nilai) & "'"
'    Set Rs = Conn.Execute(sql)
'    If Not Rs.EOF Then
'        isDataValid = True
'    End If
'Exit Function
'err:
'End Function
'
'Function unSQLi(sumber As String)
'Dim hasil As Variant
'If sumber = Null Then
'    hasil = ""
'Else
'    hasil = Replace(sumber, "\'", "'")
'    hasil = Replace(hasil, "\\", "\")
'End If
'    unSQLi = hasil
'End Function
'
'Public Function FExists(OrigFile As String) As Boolean
'    Dim fs As Object
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    FExists = fs.fileexists(OrigFile)
'End Function
'



'
''untuk format lVdetail
'Sub formatFind(Lv As ListView, arrHeader() As String, arrHWidth() As Integer)
'    Lv.ColumnHeaders.Clear
'    Dim i As Integer
'    For i = 0 To UBound(arrHeader)
'        Lv.ColumnHeaders.Add , , arrHeader(i), arrHWidth(i)
'    Next
'End Sub
'
''untuk cari data di LVDetail
'Sub findData(Lv As ListView, kolom() As String, tbl As String, _
'            qWhere As String)
'Lv.ListItems.Clear
'On Error GoTo err
'    Dim i As Integer
'    sql = "select "
'    For i = 0 To UBound(kolom)
'        'If kolom(i) <> "" Then 'biar dak error
'            sql = sql & kolom(i)
'            If i < (UBound(kolom)) Then sql = sql & ", "
'        'End If
'    Next
'    sql = sql & " from " & tbl & " WHERE " & qWhere
'    Set Rs = Conn.Execute(sql)
'    While Not Rs.EOF
'        With Lv.ListItems.Add
'            .Text = Rs(0)
'            Dim j As Integer
'            For j = 1 To Rs.Fields.Count - 1
'                .SubItems(j) = Rs(j)
'            Next
'        End With
'        Rs.MoveNext
'    Wend
'    Lv.Visible = True
'Exit Sub
'err:
'    Lv.ListItems.Clear
'    Lv.Visible = False
'End Sub

'
''fungsi untuk menukar kolom
'Function getValueFrom(tbl As String, kol As String, strWhere As String) As String
'getValueFrom = ""
'On Error GoTo err
'    sql = "select " & kol & " from " & tbl & "  WHERE " & strWhere
'    Set Rs = Conn.Execute(sql)
'    If Not Rs.EOF Then
'        If IsNull(Rs(0)) = False Then
'            getValueFrom = Rs(0)
'        End If
'    End If
'Exit Function
'err:
'    getValueFrom = ""
'End Function
