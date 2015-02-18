Attribute VB_Name = "model"

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
