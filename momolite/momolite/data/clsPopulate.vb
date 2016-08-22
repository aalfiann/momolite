
'The MIT License (MIT)

'Copyright(c) 2016 M ABD AZIZ ALFIAN (http://about.me/azizalfian)

'Permission Is hereby granted, free Of charge, to any person obtaining a copy of this software And associated documentation 
'files(the "Software"), To deal In the Software without restriction, including without limitation the rights To use, copy, modify,
'merge, publish, distribute, sublicense, And/Or sell copies Of the Software, And to permit persons to whom the Software Is furnished 
'to do so, subject to the following conditions:

'The above copyright notice And this permission notice shall be included In all copies Or substantial portions Of the Software.

'THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of 
'MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS Or COPYRIGHT HOLDERS BE LIABLE 
'For ANY CLAIM, DAMAGES Or OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT Or OTHERWISE, ARISING FROM, OUT OF Or IN CONNECTION 
'With THE SOFTWARE Or THE USE Or OTHER DEALINGS In THE SOFTWARE.

Imports System.Windows.Forms
Imports MySql.Data.MySqlClient
Namespace data
    ''' <summary>Class Populate</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>31 Juli 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>1.5.0</version>
    ''' <requirement>
    ''' - Imports System.Windows.Forms
    ''' </requirement>
    Public Class Populate

        ''' <summary>
        ''' Class populate dari objek datatable
        ''' </summary>
        Public Class FromDataTable
#Region "Property Progress Bar"
            Private _progbar As ProgressBar = Nothing

            ''' <summary>
            ''' Menentukan objek ProgressBar untuk menampilkan progressbar
            ''' </summary>
            ''' <returns>ProgressBar</returns>
            Public Property Progressbar() As ProgressBar
                Set(value As ProgressBar)
                    _progbar = value
                End Set
                Get
                    Return _progbar
                End Get
            End Property
#End Region

            ''' <summary>
            ''' Menampilkan data ke dalam ListView
            ''' </summary>
            ''' <param name="dt">Nama objek datatable</param>
            ''' <param name="lv">Nama objek ListView</param>
            ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Function ToListView(dt As DataTable, lv As ListView, Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    With lv
                        .Clear()
                        .View = View.Details
                        .FullRowSelect = True
                        .GridLines = True
                        'buat kolom otomatis
                        .Columns.Clear()
                        For Each kolom As DataColumn In dt.Columns
                            .Columns.Add(kolom.ColumnName)
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'menambah data row
                        For Each row As DataRow In dt.Rows
                            If formatDateTime = Nothing Then
                                .Items.Add(row.Item(0).ToString)
                            Else
                                If TypeOf (row.Item(0)) Is DateTime Then
                                    .Items.Add(DirectCast(row.Item(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                Else
                                    .Items.Add(row.Item(0).ToString)
                                End If
                            End If


                            For i As Integer = 1 To dt.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    .Items(.Items.Count - 1).SubItems.Add(row.Item(i).ToString)
                                Else
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        .Items(.Items.Count - 1).SubItems.Add(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                    Else
                                        .Items(.Items.Count - 1).SubItems.Add(row.Item(i).ToString)
                                    End If
                                End If
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        If autoSize = True Then .AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Menampilkan data ke dalam DataGridView
            ''' </summary>
            ''' <param name="dt">Nama objek datatable</param>
            ''' <param name="dg">Nama objek DataGridView</param>
            ''' <param name="addRows">Row DataGridView dapat ditambahkan oleh user. Default = False.</param>
            ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Function ToDataGridView(dt As DataTable, dg As DataGridView, Optional addRows As Boolean = False,
                                      Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                Try
                    'Clear DataGridView
                    dg.DataSource = Nothing
                    dg.Columns.Clear()
                    dg.Rows.Clear()

                    If formatDateTime = Nothing Then
                        If Progressbar IsNot Nothing Then 'iteration mode
                            'Buat header
                            For Each kolom As DataColumn In dt.Columns
                                dg.Columns.Add(kolom.ColumnName, kolom.ColumnName)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'Menambah data row
                            Dim rowData As String() = New String(dt.Columns.Count - 1) {}
                            For Each row As DataRow In dt.Rows
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    rowData(i) = row.Item(i).ToString
                                Next
                                dg.Rows.Add(rowData)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            dg.AllowUserToAddRows = addRows
                            If autoSize = True Then
                                dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                                dg.AutoResizeColumns()
                            End If
                        Else 'simple way
                            'Standar datetime
                            Application.DoEvents()
                            dg.DataSource = dt
                            dg.AllowUserToAddRows = addRows
                            If autoSize = True Then
                                dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                                dg.AutoResizeColumns()
                            End If
                        End If
                    Else
                        'Formatted datetime

                        'Buat header
                        For Each kolom As DataColumn In dt.Columns
                            dg.Columns.Add(kolom.ColumnName, kolom.ColumnName)
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'Menambah data row
                        Dim rowData As String() = New String(dt.Columns.Count - 1) {}
                        Application.DoEvents()
                        For Each row As DataRow In dt.Rows
                            For i As Integer = 0 To dt.Columns.Count - 1
                                If TypeOf (row.Item(i)) Is DateTime Then
                                    rowData(i) = DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                Else
                                    rowData(i) = row.Item(i).ToString
                                End If
                            Next
                            dg.Rows.Add(rowData)
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next

                        dg.AllowUserToAddRows = addRows
                        If autoSize = True Then
                            dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                            dg.AutoResizeColumns()
                        End If
                    End If

                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Menampilkan data ke dalam TreeView
            ''' </summary>
            ''' <param name="dt">Nama objek datatable</param>
            ''' <param name="tv">Nama objek TreeView</param>
            ''' <param name="keyValue">Custom Key TreeView</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Function ToTreeView(dt As DataTable, tv As TreeView, keyValue As String, Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    With tv
                        .Nodes.Clear()
                        Dim key As String
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'menambah data row
                        For Each row As DataRow In dt.Rows
                            key = keyValue + (.Nodes.Count - 1).ToString
                            If formatDateTime = Nothing Then
                                .Nodes.Add(key, row.Item(0).ToString)
                            Else
                                If TypeOf (row.Item(0)) Is DateTime Then
                                    .Nodes.Add(key, DirectCast(row.Item(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                Else
                                    .Nodes.Add(key, row.Item(0).ToString)
                                End If
                            End If

                            For i As Integer = 1 To dt.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, row.Item(i).ToString)
                                Else
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                    Else
                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, row.Item(i).ToString)
                                    End If
                                End If

                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .ExpandAll()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Menampilkan data ke ComboBox
            ''' </summary>
            ''' <param name="dt">Nama objek datatable</param>
            ''' <param name="cmb">Nama objek ComboBox</param>
            ''' <param name="displayMember">Kolom Query yang akan ditampilkan di ComboBox sebagai DisplayMember</param>
            ''' <param name="valueMember">Kolom Query yang akan di jadikan Nilai dari DisplayMember di Combobox</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Function ToComboBox(dt As DataTable, cmb As ComboBox, displayMember As String, valueMember As String,
                                       Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    If dt.Rows.Count > 0 Then
                        cmb.DataSource = dt
                        cmb.DisplayMember = displayMember
                        cmb.ValueMember = valueMember
                        If formatDateTime <> Nothing Then
                            cmb.FormatString = formatDateTime
                        End If
                    End If
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function
        End Class

        ''' <summary>
        ''' Class populate dari objek dataset
        ''' </summary>
        Public Class FromDataSet
#Region "Property Progress Bar"
            Private _progbar As ProgressBar = Nothing

            ''' <summary>
            ''' Menentukan objek ProgressBar untuk menampilkan progressbar
            ''' </summary>
            ''' <returns>ProgressBar</returns>
            Public Property Progressbar() As ProgressBar
                Set(value As ProgressBar)
                    _progbar = value
                End Set
                Get
                    Return _progbar
                End Get
            End Property
#End Region

            ''' <summary>
            ''' Menampilkan data ke dalam ListView
            ''' </summary>
            ''' <param name="ds">Nama objek dataset</param>
            ''' <param name="lv">Nama objek ListView</param>
            ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
            ''' <param name="tables">Index Table di DataSet. Default = 0.</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Function ToListView(ds As DataSet, lv As ListView, Optional tables As Integer = 0, Optional autoSize As Boolean = True,
                                  Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    With lv
                        .Clear()
                        .View = View.Details
                        .FullRowSelect = True
                        .GridLines = True
                        'buat kolom otomatis
                        .Columns.Clear()
                        For Each kolom As DataColumn In ds.Tables(tables).Columns
                            .Columns.Add(kolom.ColumnName)
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = ds.Tables(tables).Rows.Count
                        End If
                        'menambah data row
                        For Each row As DataRow In ds.Tables(tables).Rows
                            If formatDateTime = Nothing Then
                                .Items.Add(row.Item(0).ToString)
                            Else
                                If TypeOf (row.Item(0)) Is DateTime Then
                                    .Items.Add(DirectCast(row.Item(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                Else
                                    .Items.Add(row.Item(0).ToString)
                                End If
                            End If

                            For i As Integer = 1 To ds.Tables(tables).Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    .Items(.Items.Count - 1).SubItems.Add(row.Item(i).ToString)
                                Else
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        .Items(.Items.Count - 1).SubItems.Add(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                    Else
                                        .Items(.Items.Count - 1).SubItems.Add(row.Item(i).ToString)
                                    End If
                                End If

                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        If autoSize = True Then .AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Menampilkan data ke dalam DataGridView
            ''' </summary>
            ''' <param name="ds">Nama objek dataset</param>
            ''' <param name="dg">Nama objek DataGridView</param>
            ''' <param name="addRows">Row DataGridView dapat ditambahkan oleh user</param>
            ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
            ''' <param name="table">Index Table di DataSet. Default = 0.</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Function ToDataGridView(ds As DataSet, dg As DataGridView, Optional table As Integer = 0, Optional addRows As Boolean = False,
                                      Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                Try
                    If formatDateTime = Nothing Then
                        If Progressbar IsNot Nothing Then 'iteration mode
                            'Buat header
                            For Each kolom As DataColumn In ds.Tables(table).Columns
                                dg.Columns.Add(kolom.ColumnName, kolom.ColumnName)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = ds.Tables(table).Rows.Count
                            End If
                            'Menambah data row
                            Dim rowData As String() = New String(ds.Tables(table).Columns.Count - 1) {}
                            For Each row As DataRow In ds.Tables(table).Rows
                                For i As Integer = 0 To ds.Tables(table).Columns.Count - 1
                                    rowData(i) = row.Item(i).ToString
                                Next
                                dg.Rows.Add(rowData)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            dg.AllowUserToAddRows = addRows
                            If autoSize = True Then
                                dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                                dg.AutoResizeColumns()
                            End If
                        Else 'simple way
                            'Standar datetime
                            dg.DataSource = ds.Tables(table)
                            dg.AllowUserToAddRows = addRows
                            If autoSize = True Then
                                dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                                dg.AutoResizeColumns()
                            End If
                        End If
                    Else
                        'Formatted datetime

                        'Buat header
                        For Each kolom As DataColumn In ds.Tables(table).Columns
                            dg.Columns.Add(kolom.ColumnName, kolom.ColumnName)
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = ds.Tables(table).Rows.Count
                        End If
                        'Menambah data row
                        Dim rowData As String() = New String(ds.Tables(table).Columns.Count - 1) {}
                        For Each row As DataRow In ds.Tables(table).Rows
                            For i As Integer = 0 To ds.Tables(table).Columns.Count - 1
                                If TypeOf (row.Item(i)) Is DateTime Then
                                    rowData(i) = DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                Else
                                    rowData(i) = row.Item(i).ToString
                                End If
                            Next
                            dg.Rows.Add(rowData)
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next

                        dg.AllowUserToAddRows = addRows
                        If autoSize = True Then
                            dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                            dg.AutoResizeColumns()
                        End If
                    End If
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Menampilkan data ke dalam TreeView
            ''' </summary>
            ''' <param name="ds">Nama objek datatable</param>
            ''' <param name="tv">Nama objek TreeView</param>
            ''' <param name="keyValue">Custom Key TreeView</param>
            ''' <param name="table">Index Table di DataSet. Default = 0.</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Function ToTreeView(ds As DataSet, tv As TreeView, keyValue As String, Optional table As Integer = 0,
                                       Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    With tv
                        .Nodes.Clear()
                        Dim key As String
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = ds.Tables(table).Rows.Count
                        End If
                        'menambah data row
                        For Each row As DataRow In ds.Tables(table).Rows
                            key = keyValue + (.Nodes.Count - 1).ToString
                            If formatDateTime = Nothing Then
                                'Standar datetime
                                .Nodes.Add(key, row.Item(0).ToString)
                            Else
                                'Formatted datetime
                                If TypeOf (row.Item(0)) Is DateTime Then
                                    .Nodes.Add(key, DirectCast(row.Item(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                Else
                                    .Nodes.Add(key, row.Item(0).ToString)
                                End If
                            End If

                            For i As Integer = 1 To ds.Tables(table).Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, row.Item(i).ToString)
                                Else
                                    'Formatted datetime
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                    Else
                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, row.Item(i).ToString)
                                    End If
                                End If

                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .ExpandAll()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Menampilkan data ke ComboBox
            ''' </summary>
            ''' <param name="ds">Nama objek dataset</param>
            ''' <param name="cmb">Nama objek ComboBox</param>
            ''' <param name="displayMember">Kolom Query yang akan ditampilkan di ComboBox sebagai DisplayMember</param>
            ''' <param name="valueMember">Kolom Query yang akan di jadikan Nilai dari DisplayMember di Combobox</param>
            ''' <param name="table">Index Table di DataSet. Default = 0.</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Function ToComboBox(ds As DataSet, cmb As ComboBox, displayMember As String, valueMember As String,
                                  Optional table As Integer = 0, Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    If ds.Tables(table).Rows.Count > 0 Then
                        cmb.DataSource = ds.Tables(table)
                        cmb.DisplayMember = displayMember
                        cmb.ValueMember = valueMember
                        If formatDateTime <> Nothing Then
                            cmb.FormatString = formatDateTime
                        End If
                    End If
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function
        End Class

        ''' <summary>
        ''' Class paginasi data
        ''' </summary>
        Public Class Paginate

            ''' <summary>
            ''' Class paginasi data dari objek datatable
            ''' </summary>
            Public Class FromDataTable


                ''' <summary>
                ''' Menampilkan data yang telah di paginasi ke dalam ListView
                ''' </summary>
                ''' <param name="dt">Nama objek DataTable</param>
                ''' <param name="lv">Nama objek ListView</param>
                ''' <param name="page">Halaman yang akan dipilih untuk menampilkan data</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="imgIndex">Index gambar. Default = 0.</param>
                ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Function ToListView(ByVal dt As DataTable, ByVal lv As ListView, ByVal page As Integer, ByVal itemPerPage As Integer,
                                      Optional ByVal imgIndex As Integer = 0, Optional autoSize As Boolean = True,
                                           Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                    Try
                        Dim m As Integer = 0

                        With lv
                            .Clear()
                            .View = View.Details
                            .FullRowSelect = True
                            .GridLines = True
                            .Columns.Clear()

                            'Menambahkan Nama Kolom ke listview dari datatable
                            Dim i As Integer
                            For i = 0 To dt.Columns.Count - 1
                                .Columns.Add(dt.Columns(i).ColumnName)
                            Next

                            'Menambahkan records ke listview dari datatable
                            Dim l As Integer, k As Integer

                            l = (page - 1) * itemPerPage
                            k = ((page) * itemPerPage)

                            While l < k
                                If l >= dt.Rows.Count Then
                                    Exit While
                                End If

                                If formatDateTime = Nothing Then
                                    .Items.Add(dt.Rows(l)(0).ToString(), imgIndex)
                                Else
                                    If TypeOf (dt.Rows(l)(0)) Is DateTime Then
                                        .Items.Add(DirectCast(dt.Rows(l)(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), imgIndex)
                                    Else
                                        .Items.Add(dt.Rows(l)(0).ToString(), imgIndex)
                                    End If
                                End If

                                For j As Integer = 1 To dt.Columns.Count - 1
                                    If Not IsDBNull(dt.Rows(l)(j)) Then
                                        If formatDateTime = Nothing Then
                                            .Items(m).SubItems.Add(dt.Rows(l)(j).ToString())
                                        Else
                                            If TypeOf (dt.Rows(l)(j)) Is DateTime Then
                                                .Items(m).SubItems.Add(DirectCast(dt.Rows(l)(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                            Else
                                                .Items(m).SubItems.Add(dt.Rows(l)(j).ToString())
                                            End If
                                        End If

                                    Else
                                        .Items(m).SubItems.Add("")
                                    End If
                                Next
                                m += 1
                                l += 1
                            End While

                            'Mengukur size kolom listview
                            If autoSize = True Then .AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                        End With
                        Return True
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataTable")
                        Return False
                    Finally
                        GC.Collect()
                    End Try

                End Function

                ''' <summary>
                ''' Menampilkan data yang telah di paginasi ke dalam DataGridView
                ''' </summary>
                ''' <param name="dt">Nama objek DataTable</param>
                ''' <param name="dg">Nama objek DataGridView</param>
                ''' <param name="page">Halaman yang akan dipilih untuk menampilkan data</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="addRows">Row DataGridView dapat ditambahkan oleh user. Default = False.</param>
                ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Function ToDataGridView(dt As DataTable, dg As DataGridView, page As Integer, itemPerPage As Integer, Optional addRows As Boolean = False,
                                          Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                    Try
                        dg.Rows.Clear()
                        dg.Columns.Clear()
                        dg.AllowUserToAddRows = addRows

                        'Load the table data in the datagridview
                        Dim i As Integer = 0
                        Dim j As Integer = 0
                        Dim m As Integer = 0

                        'For Adding Column Names from the datatable
                        For i = 0 To dt.Columns.Count - 1
                            dg.Columns.Add(dt.Columns(i).ColumnName, dt.Columns(i).ColumnName)
                        Next

                        'for adding records to the listview from datatable
                        Dim l As Integer, k As Integer

                        l = (page - 1) * itemPerPage 'Start Row
                        k = ((page) * itemPerPage) 'End Row

                        While l < k
                            If l >= dt.Rows.Count Then
                                Exit While
                            End If

                            'Check Row Position
                            Dim checkRow As Integer
                            If dt.Rows.Count - l <= itemPerPage Then
                                checkRow = dt.Rows.Count - l
                            Else
                                checkRow = itemPerPage
                            End If

                            'Proses input data ke dalam DataGridView
                            For x As Integer = 0 To checkRow - 1
                                dg.Rows.Add()
                                For j = 0 To dt.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        If Not System.Convert.IsDBNull(dt.Rows(l)(j)) Then
                                            dg.Rows(x).Cells(j).Value = dt.Rows(l)(j).ToString()
                                        Else
                                            dg.Rows(x).Cells(j).Value = ""
                                        End If
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dt.Rows(l)(j)) Is DateTime Then
                                            If Not System.Convert.IsDBNull(dt.Rows(l)(j)) Then
                                                dg.Rows(x).Cells(j).Value = DirectCast(dt.Rows(l)(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                            Else
                                                dg.Rows(x).Cells(j).Value = ""
                                            End If
                                        Else
                                            If Not System.Convert.IsDBNull(dt.Rows(l)(j)) Then
                                                dg.Rows(x).Cells(j).Value = dt.Rows(l)(j).ToString()
                                            Else
                                                dg.Rows(x).Cells(j).Value = ""
                                            End If
                                        End If
                                    End If
                                Next
                                m += 1
                                l += 1
                            Next
                        End While
                        If autoSize = True Then
                            dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                            dg.AutoResizeColumns()
                        End If
                        Return True
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataTable")
                        Return False
                    Finally
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Menghitung total jumlah halaman yang telah di paginasi
                ''' </summary>
                ''' <param name="dt">Nama objek DataTable</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function TotalPage(dt As DataTable, itemPerPage As Integer, Optional showException As Boolean = True) As Integer
                    Try
                        Dim totalPages As Double
                        Dim totalRecords As Integer

                        totalRecords = dt.Rows.Count
                        totalPages = CDbl(totalRecords) / CDbl(itemPerPage)
                        totalPages = CInt(Math.Ceiling(totalPages))

                        If totalPages > 0 Then
                            Return CInt(totalPages)
                        Else
                            Return CInt(totalPages = 1)
                        End If
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataTable")
                        Return 1
                    End Try
                End Function

                ''' <summary>
                ''' Menghitung Row awal di halaman paginasi
                ''' </summary>
                ''' <param name="page">Halaman yang akan dipilih untuk menampilkan data</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function StartRow(page As Integer, itemPerPage As Integer, Optional showException As Boolean = True) As Integer
                    Try
                        Dim SR As Integer
                        SR = ((page - 1) * itemPerPage) + 1 'Start Row
                        Return SR
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataTable")
                        Return 0
                    End Try
                End Function

                ''' <summary>
                ''' Menghitung Row akhir di halaman paginasi
                ''' </summary>
                ''' <param name="page">Halaman yang akan dipilih untuk menampilkan data</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function EndRow(page As Integer, itemPerPage As Integer, Optional showException As Boolean = True) As Integer
                    Try
                        Dim ER As Integer
                        ER = ((page) * itemPerPage) 'End Row
                        Return ER
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataTable")
                        Return 0
                    End Try
                End Function

                ''' <summary>
                ''' Menampilkan data yang telah dipaginasi pada halaman berikutnya
                ''' </summary>
                ''' <param name="dt">Nama objek DataTable</param>
                ''' <param name="pageNow">Halaman sekarang</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function NextPage(dt As DataTable, pageNow As Integer, itemPerPage As Integer, Optional showException As Boolean = True) As Integer
                    Try
                        If pageNow < TotalPage(dt, itemPerPage) Then
                            pageNow += 1
                            Return pageNow
                        Else
                            Return pageNow
                        End If
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataTable")
                        Return 0
                    End Try
                End Function

                ''' <summary>
                ''' Menampilkan data yang telah dipaginasi pada halaman sebelumnya
                ''' </summary>
                ''' <param name="pageNow">Halaman sekarang</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function PrevPage(pageNow As Integer, Optional showException As Boolean = True) As Integer
                    Try
                        If pageNow > 1 Then
                            pageNow -= 1
                            Return pageNow
                        Else
                            Return 1
                        End If
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataTable")
                        Return 0
                    End Try
                End Function
            End Class

            ''' <summary>
            ''' Class paginasi data dari objek datagridview
            ''' </summary>
            Public Class FromDataSet


                ''' <summary>
                ''' Menampilkan data yang telah di paginasi ke dalam ListView
                ''' </summary>
                ''' <param name="ds">Nama objek DataSet</param>
                ''' <param name="lv">Nama objek ListView</param>
                ''' <param name="page">Halaman yang akan dipilih untuk menampilkan data</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="imgIndex">Index gambar. Default = 0.</param>
                ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                ''' <param name="table">Index Table di DataSet. Default = 0.</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Function ToListView(ds As DataSet, lv As ListView, page As Integer, itemPerPage As Integer, Optional imgIndex As Integer = 0,
                                      Optional table As Integer = 0, Optional autoSize As Boolean = True,
                                           Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                    Try
                        'Konversi ke DataTable
                        Dim DT As New DataTable
                        DT = ds.Tables(table)

                        'Load the table data in the listview
                        Dim m As Integer = 0

                        With lv
                            .Clear()
                            .View = View.Details
                            .FullRowSelect = True
                            .GridLines = True
                            .Columns.Clear()

                            ' for Adding Column Names from the datatable
                            Dim i As Integer
                            For i = 0 To DT.Columns.Count - 1
                                .Columns.Add(DT.Columns(i).ColumnName)
                            Next

                            'for adding records to the listview from datatable
                            Dim l As Integer, k As Integer

                            l = (page - 1) * itemPerPage
                            k = ((page) * itemPerPage)

                            While l < k
                                If l >= DT.Rows.Count Then
                                    Exit While
                                End If

                                If formatDateTime = Nothing Then
                                    .Items.Add(DT.Rows(l)(0).ToString(), imgIndex)
                                Else
                                    If TypeOf (DT.Rows(l)(0)) Is DateTime Then
                                        .Items.Add(DirectCast(DT.Rows(l)(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), imgIndex)
                                    Else
                                        .Items.Add(DT.Rows(l)(0).ToString(), imgIndex)
                                    End If
                                End If

                                For j As Integer = 1 To DT.Columns.Count - 1
                                    If Not System.Convert.IsDBNull(DT.Rows(l)(j)) Then
                                        If formatDateTime = Nothing Then
                                            .Items(m).SubItems.Add(DT.Rows(l)(j).ToString())
                                        Else
                                            If TypeOf (DT.Rows(l)(j)) Is DateTime Then
                                                .Items(m).SubItems.Add(DirectCast(DT.Rows(l)(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                            Else
                                                .Items(m).SubItems.Add(DT.Rows(l)(j).ToString())
                                            End If
                                        End If
                                    Else
                                        .Items(m).SubItems.Add("")
                                    End If
                                Next
                                m += 1
                                l += 1
                            End While
                            If autoSize = True Then .AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                        End With
                        Return True
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataSet")
                        Return False
                    Finally
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Menampilkan data yang telah di paginasi ke dalam DataGridView
                ''' </summary>
                ''' <param name="ds">Nama objek DataSet</param>
                ''' <param name="dg">Nama objek DataGridView</param>
                ''' <param name="page">Halaman yang akan dipilih untuk menampilkan data</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="addRows">Row DataGridView dapat ditambahkan oleh user. Default = False.</param>
                ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                ''' <param name="table">Index Table di DataSet. Default = 0.</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Function ToDataGrid(ds As DataSet, dg As DataGridView, page As Integer, itemPerPage As Integer, Optional table As Integer = 0,
                                      Optional addRows As Boolean = False, Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                    Try
                        dg.Rows.Clear()
                        dg.Columns.Clear()
                        dg.AllowUserToAddRows = addRows

                        'Konversi ke DataTable
                        Dim DT As New DataTable
                        DT = ds.Tables(table)

                        'Load the table data in the datagridview
                        Dim i As Integer = 0
                        Dim j As Integer = 0
                        Dim m As Integer = 0

                        'For Adding Column Names from the datatable
                        For i = 0 To DT.Columns.Count - 1
                            dg.Columns.Add(DT.Columns(i).ColumnName, DT.Columns(i).ColumnName)
                        Next

                        'for adding records to the listview from datatable
                        Dim l As Integer, k As Integer

                        l = (page - 1) * itemPerPage 'Start Row
                        k = ((page) * itemPerPage) 'End Row

                        While l < k
                            If l >= DT.Rows.Count Then
                                Exit While
                            End If

                            'Check Row Position
                            Dim checkRow As Integer
                            If DT.Rows.Count - l <= itemPerPage Then
                                checkRow = DT.Rows.Count - l
                            Else
                                checkRow = itemPerPage
                            End If

                            'Proses input data ke dalam DataGridView
                            For x As Integer = 0 To checkRow - 1
                                dg.Rows.Add()
                                For j = 0 To DT.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        If Not System.Convert.IsDBNull(DT.Rows(l)(j)) Then
                                            dg.Rows(x).Cells(j).Value = DT.Rows(l)(j).ToString()
                                        Else
                                            dg.Rows(x).Cells(j).Value = ""
                                        End If
                                    Else
                                        'Formatted datetime
                                        If TypeOf (DT.Rows(l)(j)) Is DateTime Then
                                            If Not System.Convert.IsDBNull(DT.Rows(l)(j)) Then
                                                dg.Rows(x).Cells(j).Value = DirectCast(DT.Rows(l)(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                            Else
                                                dg.Rows(x).Cells(j).Value = ""
                                            End If
                                        Else
                                            If Not System.Convert.IsDBNull(DT.Rows(l)(j)) Then
                                                dg.Rows(x).Cells(j).Value = DT.Rows(l)(j).ToString()
                                            Else
                                                dg.Rows(x).Cells(j).Value = ""
                                            End If
                                        End If
                                    End If

                                Next
                                m += 1
                                l += 1
                            Next
                        End While
                        If autoSize = True Then
                            dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                            dg.AutoResizeColumns()
                        End If
                        Return True
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataSet")
                        Return False
                    Finally
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Menghitung total jumlah halaman yang telah di paginasi
                ''' </summary>
                ''' <param name="ds">Nama objek DataSet</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="table">Index Table di DataSet. Default = 0.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function TotalPage(ds As DataSet, itemPerPage As Integer, Optional table As Integer = 0, Optional showException As Boolean = True) As Integer
                    Try
                        'Konversi ke DataTable
                        Dim DT As New DataTable
                        DT = ds.Tables(table)

                        Dim totalPages As Double
                        Dim totalRecords As Integer

                        totalRecords = DT.Rows.Count
                        totalPages = CDbl(totalRecords) / CDbl(itemPerPage)
                        totalPages = CInt(Math.Ceiling(totalPages))

                        If totalPages > 0 Then
                            Return CInt(totalPages)
                        Else
                            Return CInt(totalPages = 1)
                        End If
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataSet")
                        Return 1
                    End Try
                End Function

                ''' <summary>
                ''' Menghitung Row awal di halaman paginasi
                ''' </summary>
                ''' <param name="page">Halaman yang akan dipilih untuk menampilkan data</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function StartRow(page As Integer, itemPerPage As Integer, Optional showException As Boolean = True) As Integer
                    Try
                        Dim SR As Integer
                        SR = ((page - 1) * itemPerPage) + 1 'Start Row
                        Return SR
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataSet")
                        Return 0
                    End Try
                End Function

                ''' <summary>
                ''' Menghitung Row akhir di halaman paginasi
                ''' </summary>
                ''' <param name="page">Halaman yang akan dipilih untuk menampilkan data</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function EndRow(page As Integer, itemPerPage As Integer, Optional showException As Boolean = True) As Integer
                    Try
                        Dim ER As Integer
                        ER = ((page) * itemPerPage) 'End Row
                        Return ER
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataSet")
                        Return 0
                    End Try
                End Function

                ''' <summary>
                ''' Menampilkan data yang telah dipaginasi pada halaman berikutnya
                ''' </summary>
                ''' <param name="ds">Nama objek DataSet</param>
                ''' <param name="pageNow">Halaman sekarang</param>
                ''' <param name="itemPerPage">Jumlah item yang akan di tampilkan di setiap halaman</param>
                ''' <param name="table">Index Table di DataSet. Default = 0.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function NextPage(ds As DataSet, pageNow As Integer, itemPerPage As Integer,
                                         Optional table As Integer = 0, Optional showException As Boolean = True) As Integer
                    Try
                        'Konversi ke DataTable
                        Dim DT As New DataTable
                        DT = ds.Tables(table)

                        If pageNow < TotalPage(ds, itemPerPage, table) Then
                            pageNow += 1
                            Return pageNow
                        Else
                            Return pageNow
                        End If
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataSet")
                        Return 0
                    End Try
                End Function

                ''' <summary>
                ''' Menampilkan data yang telah dipaginasi pada halaman sebelumnya
                ''' </summary>
                ''' <param name="pageNow">Halaman sekarang</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Integer</returns>
                Public Function PrevPage(pageNow As Integer, Optional showException As Boolean = True) As Integer
                    Try
                        If pageNow > 1 Then
                            pageNow -= 1
                            Return pageNow
                        Else
                            Return 1
                        End If
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Paginate.FromDataSet")
                        Return 0
                    End Try
                End Function
            End Class
        End Class

        ''' <summary>
        ''' Class options control objek
        ''' </summary>
        Public Class Options

            ''' <summary>
            ''' Class listview control
            ''' </summary>
            Public Class ListviewControl
                ''' <summary>
                ''' Membuat auto size kolom ListView secara custom dari datatable
                ''' </summary>
                ''' <param name="dt">Nama objek DataTable</param>
                ''' <param name="lv">Nama objek ListView</param>
                ''' <param name="defWidth">Size max awal secara default. DefaultWidth = 1450.</param>
                ''' <param name="maxWidth">Size max akhir. Default = 2000.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Function CustomSizeDT(dt As DataTable, lv As ListView, Optional defWidth As Integer = 1450,
                                        Optional maxWidth As Integer = 2000, Optional showException As Boolean = True) As Boolean
                    Try
                        Dim xsize As Integer = 0
                        Dim i As Integer = 0

                        'for rearrange the column size
                        For i = 0 To dt.Columns.Count - 1
                            xsize = CInt(lv.Width / dt.Columns.Count)

                            If xsize > defWidth Then
                                lv.Columns(i).Width = xsize
                                lv.Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
                            Else
                                lv.Columns(i).Width = maxWidth
                                lv.Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
                            End If
                        Next
                        Return True
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Options.ListviewControl")
                        Return False
                    Finally
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat auto size kolom ListView secara custom daari dataset
                ''' </summary>
                ''' <param name="ds">Nama objek DataSet</param>
                ''' <param name="lv">Nama objek ListView</param>
                ''' <param name="defWidth">Size max awal secara default. DefaultWidth = 1450.</param>
                ''' <param name="maxWidth">Size max akhir. Default = 2000.</param>
                ''' <param name="table">Index Table di DataSet. Default = 0.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Function CustomSizeDS(ds As DataSet, lv As ListView, Optional table As Integer = 0, Optional defWidth As Integer = 1450,
                                        Optional maxWidth As Integer = 2000, Optional showException As Boolean = True) As Boolean
                    Try
                        Dim xsize As Integer = 0
                        Dim i As Integer = 0

                        Dim DT As New DataTable
                        DT = ds.Tables(table)

                        'for rearrange the column size
                        For i = 0 To DT.Columns.Count - 1
                            xsize = CInt(lv.Width / DT.Columns.Count)

                            If xsize > defWidth Then
                                lv.Columns(i).Width = xsize
                                lv.Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
                            Else
                                lv.Columns(i).Width = maxWidth
                                lv.Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
                            End If
                        Next
                        Return True
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Options.ListviewControl")
                        Return False
                    Finally
                        GC.Collect()
                    End Try
                End Function
            End Class

            ''' <summary>
            ''' Class datagridview control
            ''' </summary>
            Public Class DataGridViewControl
                ''' <summary>
                ''' Membuat auto size kolom DataGridView
                ''' </summary>
                ''' <param name="dg">Nama objek DataGridView</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Function AutoSize(dg As DataGridView, Optional showException As Boolean = True) As Boolean
                    Try
                        dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                        dg.AutoResizeColumns()
                        Return True
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Populate.Options.DataGridViewControl")
                        Return False
                    Finally
                        GC.Collect()
                    End Try
                End Function
            End Class
        End Class
    End Class
End Namespace