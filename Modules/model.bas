Attribute VB_Name = "model"
Public Sub set_data_masuk(lv As ListView)
On Error Resume Next
    For i = 1 To lv.ListItems.Count
        lv.ListItems(i).SubItems(3) = Format(lv.ListItems(i).SubItems(3), "dd-mm-yy H:m")
        lv.ListItems(i).SubItems(4) = Format(lv.ListItems(i).SubItems(4), "dd-mm-yy H:m")
        lv.ListItems(i).SubItems(5) = IIf(lv.ListItems(i).SubItems(5) = "99", "batal", _
            IIf(lv.ListItems(i).SubItems(5) = "1", "diterima", "pesan"))
    Next
End Sub

Public Function cek_valid_stok(lv As ListView, tipe As Integer, kode As Integer, nama As Integer, jml As Integer) As Boolean
    On Error GoTo keluar
    Dim cStok As Double
    cek_valid_stok = True
    For i = 1 To lv.ListItems.Count
        If (LCase(lv.ListItems(i).SubItems(tipe)) = "parfum") Then
            cStok = get_parfum_stok(lv.ListItems(i).SubItems(kode))
        Else
            cStok = get_botol_stok(lv.ListItems(i).SubItems(kode))
        End If
        
        If (Val(lv.ListItems(i).SubItems(jml)) > cStok) Then
            MsgBox "Stok untuk " & lv.ListItems(i).SubItems(tipe) & " " & _
                lv.ListItems(i).SubItems(nama) & " hanya " & cStok, vbExclamation, "Stok tidak cukup"
            GoTo keluar
        End If
    Next
    Exit Function
keluar:
    cek_valid_stok = False
End Function


'Public Function sum_lv_items(lv As ListView, indexJml As Integer, Optional qwhere As String = "") As Double
''On Error GoTo keluar
'    sum_lv_items = 0
'    'MsgBox ("LISTCOUNT : " & lv.ListItems.Count)
'    For i = 1 To lv.ListItems.Count
'        If qwhere = "" Then
'            sum_lv_items = sum_lv_items + lv.ListItems(i).SubItems(indexJml)
'        Else
'            Dim qitems() As String
'            qitems = Split(qwhere, ",")
'            'MsgBox "QITEMS : " & UBound(qitems)
'            For ix = 0 To UBound(qitems)
'                MsgBox (qitems(ix))
'            Next
'        End If
'    Next
'    Exit Function
'keluar:
'    sum_lv_items = 0
'End Function

Public Function update_parfum_stok(iCur As Double, iJml As Double, pID As String) As Boolean
On Error GoTo keluar
    update_parfum_stok = True
    Dim stok As Double
        stok = iCur + iJml
    sql = "UPDATE parfum set parfum_stok = " & stok & " WHERE parfum_id = " & AntiSQLiWithQuotes(pID)
    Conn.Execute (sql)
    Exit Function
keluar:
    update_parfum_stok = False
End Function

Public Function update_botol_stok(iCur As Double, iJml As Double, pID As String) As Boolean
On Error GoTo keluar
    update_botol_stok = True
    Dim stok As Double
        stok = iCur + iJml
    sql = "UPDATE botol set botol_stok = " & stok & " WHERE botol_id = " & AntiSQLiWithQuotes(pID)
    Conn.Execute (sql)
    Exit Function
keluar:
    update_botol_stok = False
End Function

Public Function get_parfum_stok(pID As String) As Double
On Error GoTo keluar
    get_parfum_stok = 0
    Dim stok As Variant
    Dim qwhere As String
        qwhere = " parfum_id = " & AntiSQLiWithQuotes(pID)
        stok = modulGencil.getValues("parfum", "parfum_stok", qwhere)
        get_parfum_stok = CDbl(stok)
    Exit Function
keluar:
    get_parfum_stok = 0
End Function

Public Function get_botol_stok(pID As String) As Double
On Error GoTo keluar
    get_botol_stok = 0
    Dim stok As Variant
    Dim qwhere As String
        qwhere = " botol_id = " & AntiSQLiWithQuotes(pID)
        stok = modulGencil.getValues("botol", "botol_stok", qwhere)
        get_botol_stok = CDbl(stok)
    Exit Function
keluar:
    get_botol_stok = 0
End Function

Public Function is_query_have_row(query As String) As Boolean
On Error GoTo keluar
    is_query_have_row = False
    Set Rs = Conn.Execute(query)
    If Not Rs.EOF Then
        is_query_have_row = True
    End If
    Exit Function
keluar:
    is_query_have_row = False
End Function

Public Sub list_kategori(cbo As ComboBox)
    sql = "SELECT * FROM kategori order by kategori_nama"
    Set Rs = Conn.Execute(sql)
    While Not Rs.EOF
        cbo.AddItem Rs!kategori_nama
        Rs.MoveNext
    Wend
End Sub

Public Sub lv_inventory(lv As ListView, tglAwal As String, tglAkhir As String)
    Dim i%
    
    For i = 1 To lv.ListItems.Count
        lv.ListItems(i).SubItems(6) = get_sirkulasi_total(lv.ListItems(i).SubItems(3), lv.ListItems(i).SubItems(2), tglAwal, tglAkhir)
        lv.ListItems(i).SubItems(7) = get_output_total(lv.ListItems(i).SubItems(3), lv.ListItems(i).SubItems(2), tglAwal, tglAkhir)
        lv.ListItems(i).SubItems(8) = get_kecelakaan_total(lv.ListItems(i).SubItems(3), lv.ListItems(i).SubItems(2), tglAwal, tglAkhir)
        lv.ListItems(i).SubItems(4) = IIf(LCase(lv.ListItems(i).SubItems(3)) = "parfum", get_parfum_kategori(lv.ListItems(i).SubItems(2)), lv.ListItems(i).SubItems(4))
    Next
End Sub

Function get_parfum_kategori(id As String) As String
    get_parfum_kategori = ""
    sql = "SELECT k.kategori_nama " & _
         "FROM parfum_kategori pk " & _
         "JOIN kategori k on pk.pk_kategori_id = k.kategori_id " & _
         "where pk_parfum_id = " & AntiSQLiWithQuotes(id)
    Set Rs = Conn.Execute(sql)
    Dim i%, temp$
    i = 1
    While Not Rs.EOF
        temp = IIf(i = 1, Rs!kategori_nama, ", " & Rs!kategori_nama)
        get_parfum_kategori = get_parfum_kategori & temp
        i = i + 1
        Rs.MoveNext
    Wend
End Function

Function get_sirkulasi_total(tipe As String, id As String, tglAwal As String, tglAkhir As String) As Double
Dim strWhere$
    get_sirkulasi_total = 0
    strWhere = " WHERE detail_botol_id = " & AntiSQLiWithQuotes(id)
        If (LCase(tipe) = "parfum") Then
            strWhere = " WHERE detail_parfum_id = " & AntiSQLiWithQuotes(id)
        End If
    strWhere = strWhere & " AND s.sirkulasi_status = '1' AND s.sirkulasi_tanggal_terima BETWEEN " & AntiSQLiWithQuotes(tglAwal) & " AND " & AntiSQLiWithQuotes(tglAkhir)
    sql = "SELECT sd.detail_parfum_id as parfum_id, sd.detail_botol_id as botol_id, SUM(detail_jml_terima_bersih) as jml_bersih, SUM(detail_jml_terima_kotor) As jml_kotor " & _
          "From sirkulasi_detail sd JOIN sirkulasi s on sd.detail_sirkulasi_id = s.sirkulasi_id " & _
          strWhere & _
          " GROUP BY sd.detail_parfum_id, sd.detail_botol_id"
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        get_sirkulasi_total = Rs!jml_bersih
    End If
End Function

Function get_kecelakaan_total(tipe As String, id As String, tglAwal As String, tglAkhir As String) As Double
Dim strWhere$
    get_kecelakaan_total = 0
    strWhere = " WHERE kecelakaan_botol_id = " & AntiSQLiWithQuotes(id)
        If (LCase(tipe) = "parfum") Then
            strWhere = " WHERE kecelakaan_parfum_id = " & AntiSQLiWithQuotes(id)
        End If
    strWhere = strWhere & " AND kecelakaan_tanggal BETWEEN " & AntiSQLiWithQuotes(tglAwal) & " AND " & AntiSQLiWithQuotes(tglAkhir)
    sql = "SELECT kecelakaan_parfum_id, kecelakaan_botol_id, SUM(kecelakaan_jumlah) as jml  FROM kecelakaan " & strWhere
    sql = sql & " GROUP BY kecelakaan_parfum_id, kecelakaan_botol_id"
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        get_kecelakaan_total = Rs!jml
    End If
End Function

Function get_output_total(tipe As String, id As String, tglAwal As String, tglAkhir As String) As Double
Dim strWhere$
    get_output_total = 0
    strWhere = " WHERE odetail_botol_id = " & AntiSQLiWithQuotes(id)
        If (LCase(tipe) = "parfum") Then
            strWhere = " WHERE odetail_parfum_id = " & AntiSQLiWithQuotes(id)
        End If
    strWhere = strWhere & " AND o.output_tanggal BETWEEN " & AntiSQLiWithQuotes(tglAwal) & " AND " & AntiSQLiWithQuotes(tglAkhir)
    sql = "SELECT od.odetail_parfum_id as parfum_id, od.odetail_botol_id as botol_id, SUM(od.odetail_jml) as jml_keluar " & _
          "From output_detail od JOIN output o ON od.odetail_output_id = o.output_id " & _
          strWhere & _
          " GROUP BY od.odetail_parfum_id, od.odetail_botol_id "
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        get_output_total = Rs!jml_keluar
    End If
End Function


Function paging_inventory(cbo As ComboBox, nilai As Integer, strTable As String, Optional strWhere As String = "") As Boolean
    Dim total As Integer
    Dim perpage As Integer
        perpage = Val(cbo.Text)
    
    sql = "SELECT count(DISTINCT(id)) as jml FROM vw_summary_inventory "
    sql = sql & strWhere
    Set Rs = Conn.Execute(sql)
    total = Rs.Fields(0)
    
    If (nilai * perpage) > (total + perpage) Then
        paging_inventory = False
    Else
        paging_inventory = True
    End If
End Function

Public Sub show_inventory(lv As ListView, Optional strWhere As String = "", _
                            Optional strLimit As String = "", _
                            Optional intStart As Integer = 1)
On Error GoTo err
Dim i As Integer
Dim j As Integer
    lv.ListItems.Clear
    sql = "SELECT * FROM vw_summary_inventory"
    sql = sql & strWhere
    sql = sql & " ORDER BY nama "
    sql = sql & strLimit
    Set Rs = Conn.Execute(sql)
    j = intStart
    While Not Rs.EOF
        With lv.ListItems.Add
            .Text = ""
            .SubItems(1) = j
            
             For i = 1 To Rs.Fields.Count
                .SubItems(i + 1) = IIf(IsNull(Rs(i - 1)), "", Rs(i - 1))
            Next
            j = j + 1
        End With
    
        Rs.MoveNext
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


Function paging_kecelakaan(cbo As ComboBox, nilai As Integer, Optional strWhere As String = "") As Boolean
    Dim total As Integer
    Dim perpage As Integer
        perpage = Val(cbo.Text)
    
    sql = "SELECT count(kecelakaan_id) as jml " & _
          "FROM kecelakaan k " & _
          "  LEFT JOIN parfum p on k.kecelakaan_parfum_id = p.parfum_id " & _
          "  LEFT JOIN botol b on k.kecelakaan_botol_id = b.botol_id " & _
          "  LEFT JOIN user u on k.kecelakaan_user_id = u.user_id "
    sql = sql & strWhere
    Set Rs = Conn.Execute(sql)
    total = Rs.Fields(0)
    
    If (nilai * perpage) > (total + perpage) Then
        paging_kecelakaan = False
    Else
        paging_kecelakaan = True
    End If
End Function

Public Sub show_kecelakaan(lv As ListView, Optional strWhere As String = "", Optional strLimit As String = "", Optional intStart As Integer = 1)
On Error GoTo err
Dim i As Integer
Dim j As Integer
    lv.ListItems.Clear
    sql = "SELECT k.kecelakaan_id, k.kecelakaan_tanggal, " & _
          "IF(k.kecelakaan_parfum_id is NULL,'BOTOL','PARFUM') as tipe, " & _
          "IF(k.kecelakaan_parfum_id is NULL, CONCAT(b.botol_tipe,' ', b.botol_ukuran, ' ml') ,p.parfum_nama) as nama, " & _
          "k.kecelakaan_jumlah , k.kecelakaan_keterangan, u.user_nama, " & _
          "IF(k.kecelakaan_parfum_id is NULL, b.botol_id ,p.parfum_id) as id " & _
          "FROM kecelakaan k " & _
          "  LEFT JOIN parfum p on k.kecelakaan_parfum_id = p.parfum_id " & _
          "  LEFT JOIN botol b on k.kecelakaan_botol_id = b.botol_id " & _
          "  LEFT JOIN user u on k.kecelakaan_user_id = u.user_id "
    sql = sql & strWhere
    sql = sql & " ORDER BY kecelakaan_tanggal desc "
    sql = sql & strLimit
    
    Set Rs = Conn.Execute(sql)
    j = intStart
    While Not Rs.EOF
        With lv.ListItems.Add
            .Text = ""
            .SubItems(1) = j
            For i = 1 To Rs.Fields.Count
                If (i = 2) Then
                    .SubItems(i + 1) = IIf(IsNull(Rs(i - 1)), "", Format(Rs(i - 1), "dd-mm-yyyy (hh:mm:ss)"))
                Else
                    .SubItems(i + 1) = IIf(IsNull(Rs(i - 1)), "", Rs(i - 1))
                End If
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

Public Sub show_form(frm As Form, frm_parent As Form, Optional modal As Integer = 1)
    frm.Show modal, frm_parent
End Sub

Public Sub user_set_lv(lv As ListView)
    Dim ix As Integer
    If lv.ListItems.Count > 0 Then
        For ix = 1 To lv.ListItems.Count
          lv.ListItems(ix).SubItems(5) = IIf(lv.ListItems(ix).SubItems(5) = "1", "Admin", "User")
        Next
    End If
End Sub

Public Sub hapus_unset_kategori(lv As ListView, strID As String)
On Error GoTo err
    Dim ix, j As Integer
    Dim param, tmp, query As String
    Dim arr() As String
        ReDim arr(0 To lv.ListItems.Count) As String
        param = ""
        j = 0
        query = "DELETE FROM parfum_kategori WHERE pk_parfum_id = " & AntiSQLiWithQuotes(strID)
        
    If lv.ListItems.Count > 0 Then
        For ix = 1 To lv.ListItems.Count
            tmp = getValues("kategori", "kategori_id", "kategori_nama = " & AntiSQLiWithQuotes(lv.ListItems(ix).SubItems(3)))
            If lenString(str(tmp)) > 0 Then
                arr(j) = AntiSQLiWithQuotes(str(tmp))
                j = j + 1
            End If
        Next
        
        ReDim Preserve arr(0 To j - 1) As String
        'join array
            param = Join(arr, ",")
            If param <> "" Then
                query = query & " AND pk_kategori_id not IN (" & param & ")"
            End If
    End If
    
    'execute
    Conn.Execute query
    Exit Sub
err:
    MsgBox "salah yak e"
End Sub

Public Function isArrayEmpty(parArray As Variant) As Boolean
'Returns true if:
'  - parArray is not an array
'  - parArray is a dynamic array that has not been initialised (ReDim)
'  - parArray is a dynamic array has been erased (Erase)

  If IsArray(parArray) = False Then isArrayEmpty = True
  On Error Resume Next
  If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False

End Function

Public Function get_lv_kategori(lv As ListView)
Dim arr() As String
Dim i As Integer
Dim j As Integer
Dim tmp As String
    
    If lv.ListItems.Count > 0 Then
        ReDim arr(0 To lv.ListItems.Count - 1) As String
        j = 0
        For i = 1 To lv.ListItems.Count
            If lenString(lv.ListItems(i).SubItems(2)) = 0 Then
                tmp = getValues("kategori", "kategori_id", "kategori_nama = " & AntiSQLiWithQuotes(lv.ListItems(i).SubItems(3)))
                If lenString(tmp) > 0 Then
                    arr(j) = tmp
                    j = j + 1
                End If
            End If
        Next
    End If
    
    get_lv_kategori = arr
End Function

Public Function get_last_id(tbl As String, kolom As String)
On Error GoTo err
get_last_id = 1
    sql = "SELECT MAX(" & kolom & ")+1 as id FROM " & tbl
    Set Rs = Conn.Execute(sql)
    If Not Rs.EOF Then
        get_last_id = IIf(IsNull(Rs(0)) Or Rs(0) = "" Or Rs(0) = 0, 1, Val(Rs(0)))
    End If
Exit Function
err:
 get_last_id = 1
End Function

Public Sub get_kategori(cbo As ComboBox)
On Error GoTo err
    cbo.Clear
    sql = "SELECT * FROM kategori order by kategori_nama"
    Set Rs = Conn.Execute(sql)
    While Not Rs.EOF
        cbo.AddItem Rs(1)
        Rs.MoveNext
    Wend
    Rs.Close
    
    cbo.ListIndex = 0
    Exit Sub
err:
    MsgBox "Error load data kategori", vbCritical, "Error"
End Sub

Function paging_parfum(cbo As ComboBox, nilai As Integer, strTable As String, Optional strWhere As String = "") As Boolean
    Dim total As Integer
    Dim perpage As Integer
        perpage = Val(cbo.Text)
    
    sql = "SELECT count(DISTINCT(parfum_id)) as jml FROM parfum p " & _
          "LEFT JOIN parfum_kategori pk on p.parfum_id = pk.pk_parfum_id " & _
          "LEFT JOIN kategori k on pk.pk_kategori_id = k.kategori_id "
    sql = sql & strWhere
    Set Rs = Conn.Execute(sql)
    total = Rs.Fields(0)
    
    If (nilai * perpage) > (total + perpage) Then
        paging_parfum = False
    Else
        paging_parfum = True
    End If
End Function

Public Sub show_parfum(lv As ListView, Optional strWhere As String = "", Optional strLimit As String = "", Optional intStart As Integer = 1)
On Error GoTo err
Dim i As Integer
Dim j As Integer
Dim tmp As String
Dim cat As String
    tmp = ""
    cat = ""
    lv.ListItems.Clear
    sql = "SELECT p.parfum_id, p.parfum_nama, k.kategori_nama, p.parfum_remarks, p.parfum_status, p.parfum_stok " & _
          "FROM parfum p " & _
          "LEFT JOIN parfum_kategori pk on p.parfum_id = pk.pk_parfum_id " & _
          "LEFT JOIN kategori k on pk.pk_kategori_id = k.kategori_id "
    sql = sql & strWhere
    sql = sql & " ORDER BY parfum_nama "
    sql = sql & strLimit
    Set Rs = Conn.Execute(sql)
    j = intStart
    While Not Rs.EOF
            If tmp = Rs(0) Then
                cat = cat & ", " & Rs(2)
                lv.ListItems(j - 1).SubItems(4) = cat
            Else
                With lv.ListItems.Add
                    tmp = Rs(0)
                    cat = IIf(IsNull(Rs(2)), "-", Rs(2))
                    .Text = ""
                    .SubItems(1) = j
                    For i = 1 To Rs.Fields.Count
                        If (i + 1) = 6 Then
                            .SubItems(i + 1) = IIf(IsNull(Rs(i - 1)), "", IIf(Rs(i - 1) = 1, "Tersedia", "Kosong"))
                        Else
                            .SubItems(i + 1) = IIf(IsNull(Rs(i - 1)), "", Rs(i - 1))
                        End If
                    Next
                    j = j + 1
                End With
            End If
        Rs.MoveNext
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
