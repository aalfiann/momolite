
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

Imports System.Data.OleDb
Imports System.Windows.Forms
Namespace data
    ''' <summary>Class Import</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>08 August 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>2.5.0</version>
    ''' <remarks>Microsoft Access Database 12 harus sudah terinstall jika</remarks>
    ''' <requirement>
    ''' - Imports System.Data.OleDb
    ''' - Imports System.Windows.Forms
    ''' </requirement>
    Public Class Import

        ''' <summary>
        ''' Class import dari objek text
        ''' </summary>
        Public Class Text
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

#Region "Datatable"
            ''' <summary>
            ''' Import text ke datatable
            ''' </summary>
            ''' <param name="pathFile">Lokasi path file txt</param>
            ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
            ''' <param name="header">True or False. Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataTable</returns>
            Public Function ToDataTable(ByVal pathFile As String, Optional ByVal delimiter As String = ",",
                                        Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                Try
                    Dim DT As New DataTable()
                    Try
                        Using txtReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                            txtReader.SetDelimiters(New String() {delimiter})
                            txtReader.HasFieldsEnclosedInQuotes = True

                            'read column names
                            Dim colFields As String() = txtReader.ReadFields()
                            If header = True Then
                                For Each column As String In colFields
                                    Dim datecolumn As New DataColumn(column)
                                    datecolumn.AllowDBNull = True
                                    DT.Columns.Add(datecolumn)
                                Next
                            Else
                                For i As Integer = 0 To colFields.Length - 1
                                    DT.Columns.Add("Col " + (i).ToString)
                                Next
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            While Not txtReader.EndOfData
                                Dim fieldData As String() = txtReader.ReadFields()
                                'Making empty value as null
                                For i As Integer = 0 To fieldData.Length - 1
                                    If fieldData(i) = "" Then
                                        fieldData(i) = Nothing
                                    End If
                                Next
                                DT.Rows.Add(fieldData)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            End While
                        End Using
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text")
                    End Try
                    Return DT
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

#Region "Dataset"
            ''' <summary>
            ''' Import text ke dataset
            ''' </summary>
            ''' <param name="pathFile">Lokasi path file txt</param>
            ''' <param name="tableName">Nama table DataSet</param>
            ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
            ''' <param name="header">True or False. Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataSet</returns>
            Public Function ToDataSet(ByVal pathFile As String, tableName As String, Optional ByVal delimiter As String = ",",
                                      Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                Try
                    Dim DT As New DataTable()
                    Try
                        Using txtReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                            txtReader.SetDelimiters(New String() {delimiter})
                            txtReader.HasFieldsEnclosedInQuotes = True

                            'read column names
                            Dim colFields As String() = txtReader.ReadFields()
                            If header = True Then
                                For Each column As String In colFields
                                    Dim datecolumn As New DataColumn(column)
                                    datecolumn.AllowDBNull = True
                                    DT.Columns.Add(datecolumn)
                                Next
                            Else
                                For i As Integer = 0 To colFields.Length - 1
                                    DT.Columns.Add("Col " + (i).ToString)
                                Next
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            While Not txtReader.EndOfData
                                Dim fieldData As String() = txtReader.ReadFields()
                                'Making empty value as null
                                For i As Integer = 0 To fieldData.Length - 1
                                    If fieldData(i) = "" Then
                                        fieldData(i) = Nothing
                                    End If
                                Next
                                DT.Rows.Add(fieldData)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            End While
                        End Using
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text")
                    End Try
                    Dim DS As New DataSet
                    DT.TableName = tableName
                    DS.Tables.Add(DT)
                    Return DS
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

            ''' <summary>
            ''' Class import dengan data telah di dekripsi
            ''' </summary>
            Public Class Decrypt

                ''' <summary>
                ''' Class deskripsi DES
                ''' </summary>
                Public Class DES
                    Private des As New crypt.DES
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
                    ''' Import text ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String, Optional ByVal delimiter As String = ",",
                                                Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using txtReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    txtReader.SetDelimiters(New String() {delimiter})
                                    txtReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = txtReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = des.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not txtReader.EndOfData
                                        Dim fieldData As String() = txtReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = des.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.DES")
                            End Try
                            Return DT
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.DES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import text ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String, Optional ByVal delimiter As String = ",",
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using txtReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    txtReader.SetDelimiters(New String() {delimiter})
                                    txtReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = txtReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = des.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not txtReader.EndOfData
                                        Dim fieldData As String() = txtReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = des.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.DES")
                            End Try
                            Dim DS As New DataSet
                            DT.TableName = tableName
                            DS.Tables.Add(DT)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.DES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function
                End Class

                ''' <summary>
                ''' Class dekripsi 3DES
                ''' </summary>
                Public Class TripleDES
                    Private tripledes As New crypt.TripleDES
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
                    ''' Import text ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String, Optional ByVal delimiter As String = ",",
                                                Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using txtReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    txtReader.SetDelimiters(New String() {delimiter})
                                    txtReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = txtReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = tripledes.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not txtReader.EndOfData
                                        Dim fieldData As String() = txtReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = tripledes.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.TripleDES")
                            End Try
                            Return DT
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.TripleDES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import text ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String, Optional ByVal delimiter As String = ",",
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using txtReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    txtReader.SetDelimiters(New String() {delimiter})
                                    txtReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = txtReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = tripledes.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not txtReader.EndOfData
                                        Dim fieldData As String() = txtReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = tripledes.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.TripleDES")
                            End Try
                            Dim DS As New DataSet
                            DT.TableName = tableName
                            DS.Tables.Add(DT)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.TripleDES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function
                End Class

                ''' <summary>
                ''' Class dekripsi AES
                ''' </summary>
                Public Class AES
                    Private aes As New crypt.AES
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
                    ''' Import text ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String, Optional ByVal delimiter As String = ",",
                                                Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using txtReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    txtReader.SetDelimiters(New String() {delimiter})
                                    txtReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = txtReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = aes.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not txtReader.EndOfData
                                        Dim fieldData As String() = txtReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = aes.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.AES")
                            End Try
                            Return DT
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.AES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import text ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String, Optional ByVal delimiter As String = ",",
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using txtReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    txtReader.SetDelimiters(New String() {delimiter})
                                    txtReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = txtReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = aes.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not txtReader.EndOfData
                                        Dim fieldData As String() = txtReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = aes.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.AES")
                            End Try
                            Dim DS As New DataSet
                            DT.TableName = tableName
                            DS.Tables.Add(DT)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Text.Decrypt.AES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function
                End Class
            End Class

        End Class

        ''' <summary>
        ''' Class import dari objek csv
        ''' </summary>
        Public Class CSV
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

#Region "Datatable"
            ''' <summary>
            ''' Import csv ke datatable
            ''' </summary>
            ''' <param name="pathFile">Lokasi path file txt</param>
            ''' <param name="header">True or False. Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataTable</returns>
            Public Function ToDataTable(ByVal pathFile As String, Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                Try
                    Dim DT As New DataTable()
                    Try
                        Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                            csvReader.SetDelimiters(New String() {","})
                            csvReader.HasFieldsEnclosedInQuotes = True

                            'read column names
                            Dim colFields As String() = csvReader.ReadFields()
                            If header = True Then
                                For Each column As String In colFields
                                    Dim datecolumn As New DataColumn(column)
                                    datecolumn.AllowDBNull = True
                                    DT.Columns.Add(datecolumn)
                                Next
                            Else
                                For i As Integer = 0 To colFields.Length - 1
                                    DT.Columns.Add("Col " + (i).ToString)
                                Next
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            While Not csvReader.EndOfData
                                Dim fieldData As String() = csvReader.ReadFields()

                                For i As Integer = 0 To fieldData.Length - 1
                                    'Making empty value as null
                                    If fieldData(i) = "" Then fieldData(i) = Nothing
                                Next
                                DT.Rows.Add(fieldData)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            End While
                        End Using
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV")
                    End Try
                    Return DT
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

#Region "Dataset"
            ''' <summary>
            ''' Import csv ke dataset
            ''' </summary>
            ''' <param name="pathFile">Lokasi path file txt</param>
            ''' <param name="tableName">Nama table DataSet</param>
            ''' <param name="header">True or False. Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataSet</returns>
            Public Function ToDataSet(ByVal pathFile As String, tableName As String, Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                Try
                    Dim DT As New DataTable()
                    Try
                        Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                            csvReader.SetDelimiters(New String() {","})
                            csvReader.HasFieldsEnclosedInQuotes = True

                            'read column names
                            Dim colFields As String() = csvReader.ReadFields()
                            If header = True Then
                                For Each column As String In colFields
                                    Dim datecolumn As New DataColumn(column)
                                    datecolumn.AllowDBNull = True
                                    DT.Columns.Add(datecolumn)
                                Next
                            Else
                                For i As Integer = 0 To colFields.Length - 1
                                    DT.Columns.Add("Col " + (i).ToString)
                                Next
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            While Not csvReader.EndOfData
                                Dim fieldData As String() = csvReader.ReadFields()
                                For i As Integer = 0 To fieldData.Length - 1
                                    'Making empty value as null
                                    If fieldData(i) = "" Then fieldData(i) = Nothing
                                Next
                                DT.Rows.Add(fieldData)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            End While
                        End Using
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV")
                    End Try
                    Dim DS As New DataSet
                    DT.TableName = tableName
                    DS.Tables.Add(DT)
                    Return DS
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

            ''' <summary>
            ''' Class import dengan data telah di dekripsi
            ''' </summary>
            Public Class Decrypt

                ''' <summary>
                ''' Class dekripsi DES
                ''' </summary>
                Public Class DES
                    Private des As New crypt.DES
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
                    ''' Import csv ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String,
                                                Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    csvReader.SetDelimiters(New String() {","})
                                    csvReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = csvReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = des.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not csvReader.EndOfData
                                        Dim fieldData As String() = csvReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = des.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.DES")
                            End Try
                            Return DT
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.DES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import csv ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String,
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    csvReader.SetDelimiters(New String() {","})
                                    csvReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = csvReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = des.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not csvReader.EndOfData
                                        Dim fieldData As String() = csvReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = des.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.DES")
                            End Try
                            Dim DS As New DataSet
                            DT.TableName = tableName
                            DS.Tables.Add(DT)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.DES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class

                ''' <summary>
                ''' Class dekripsi 3DES
                ''' </summary>
                Public Class TripleDES
                    Private tripledes As New crypt.TripleDES
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
                    ''' Import csv ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String,
                                                Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    csvReader.SetDelimiters(New String() {","})
                                    csvReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = csvReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = tripledes.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not csvReader.EndOfData
                                        Dim fieldData As String() = csvReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = tripledes.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.TripleDES")
                            End Try
                            Return DT
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.TripleDES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import csv ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String,
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    csvReader.SetDelimiters(New String() {","})
                                    csvReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = csvReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = tripledes.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not csvReader.EndOfData
                                        Dim fieldData As String() = csvReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = tripledes.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.TripleDES")
                            End Try
                            Dim DS As New DataSet
                            DT.TableName = tableName
                            DS.Tables.Add(DT)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.TripleDES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class

                ''' <summary>
                ''' Class dekripsi AES
                ''' </summary>
                Public Class AES
                    Private aes As New crypt.AES
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
                    ''' Import csv ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String,
                                                Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    csvReader.SetDelimiters(New String() {","})
                                    csvReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = csvReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = aes.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not csvReader.EndOfData
                                        Dim fieldData As String() = csvReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = aes.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.AES")
                            End Try
                            Return DT
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.AES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import csv ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String,
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DT As New DataTable()
                            Try
                                Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pathFile)
                                    csvReader.SetDelimiters(New String() {","})
                                    csvReader.HasFieldsEnclosedInQuotes = True

                                    'read column names
                                    Dim colFields As String() = csvReader.ReadFields()

                                    'proses decrypt
                                    For a As Integer = 0 To colFields.Length - 1
                                        colFields(a) = aes.Decode(colFields(a), secretKey)
                                    Next

                                    If header = True Then
                                        For Each column As String In colFields
                                            Dim datecolumn As New DataColumn(column)
                                            datecolumn.AllowDBNull = True
                                            DT.Columns.Add(datecolumn)
                                        Next
                                    Else
                                        For i As Integer = 0 To colFields.Length - 1
                                            DT.Columns.Add("Col " + (i).ToString)
                                        Next
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Progressbar.Value = 0
                                        Progressbar.Maximum = DT.Rows.Count
                                    End If
                                    'menyimpan data
                                    While Not csvReader.EndOfData
                                        Dim fieldData As String() = csvReader.ReadFields()

                                        'proses decrypt value
                                        For b As Integer = 0 To fieldData.Length - 1
                                            fieldData(b) = aes.Decode(fieldData(b), secretKey)
                                        Next

                                        'Making empty value as null
                                        For i As Integer = 0 To fieldData.Length - 1
                                            If fieldData(i) = "" Then
                                                fieldData(i) = Nothing
                                            End If
                                        Next
                                        DT.Rows.Add(fieldData)
                                        'set progressbar
                                        If Progressbar IsNot Nothing Then
                                            Application.DoEvents()
                                            Progressbar.Value += 1
                                        End If
                                    End While
                                End Using
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.AES")
                            End Try
                            Dim DS As New DataSet
                            DT.TableName = tableName
                            DS.Tables.Add(DT)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.CSV.Decrypt.AES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class

            End Class

        End Class

        ''' <summary>
        ''' Class import dari objek tsv
        ''' </summary>
        Public Class TSV
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

#Region "Datatable"
            ''' <summary>
            ''' Import tsv ke datatable
            ''' </summary>
            ''' <param name="pathFile">Lokasi path file txt</param>
            ''' <param name="header">True or False</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataTable</returns>
            Public Function ToDataTable(ByVal pathFile As String, ByVal header As Boolean, Optional showException As Boolean = True) As DataTable
                Try
                    Dim source As String = String.Empty
                    Dim dt As DataTable = New DataTable

                    If IO.File.Exists(pathFile) Then
                        source = IO.File.ReadAllText(pathFile)
                    Else
                        Throw New IO.FileNotFoundException("Could not find the file at " & pathFile, pathFile)
                    End If

                    Dim rows() As String = source.Split({Environment.NewLine}, StringSplitOptions.None)

                    For i As Integer = 0 To rows(0).Split(Chr(9)).Length - 1
                        Dim column As String = rows(0).Split(Chr(9))(i)
                        dt.Columns.Add(If(header, column, "Col " & i + 1))
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = dt.Rows.Count
                    End If
                    'menyimpan data
                    For i As Integer = If(header, 1, 0) To rows.Length - 1
                        Dim dr As DataRow = dt.NewRow

                        For x As Integer = 0 To rows(i).Split(Chr(9)).Length - 1
                            If x <= dt.Columns.Count - 1 Then
                                dr(x) = rows(i).Split(Chr(9))(x)
                            Else
                                Throw New Exception("The number of columns on row " & i + If(header, 0, 1) & " is greater than the amount of columns in the " & If(header, "header.", "first row."))
                            End If
                        Next

                        dt.Rows.Add(dr)
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next

                    Return dt
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.TSV")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

#Region "Dataset"
            ''' <summary>
            ''' Import tsv ke dataset
            ''' </summary>
            ''' <param name="pathFile">Lokasi path file txt</param>
            ''' <param name="tableName">Nama table DataSet</param>
            ''' <param name="header">True or False</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataSet</returns>
            Public Function ToDataSet(ByVal pathFile As String, tableName As String, ByVal header As Boolean, Optional showException As Boolean = True) As DataSet
                Try
                    Dim source As String = String.Empty
                    Dim dt As DataTable = New DataTable

                    If IO.File.Exists(pathFile) Then
                        source = IO.File.ReadAllText(pathFile)
                    Else
                        Throw New IO.FileNotFoundException("Could not find the file at " & pathFile, pathFile)
                    End If

                    Dim rows() As String = source.Split({Environment.NewLine}, StringSplitOptions.None)

                    For i As Integer = 0 To rows(0).Split(Chr(9)).Length - 1
                        Dim column As String = rows(0).Split(Chr(9))(i)
                        dt.Columns.Add(If(header, column, "Col " & i + 1))
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = dt.Rows.Count
                    End If
                    'menyimpan data
                    For i As Integer = If(header, 1, 0) To rows.Length - 1
                        Dim dr As DataRow = dt.NewRow

                        For x As Integer = 0 To rows(i).Split(Chr(9)).Length - 1
                            If x <= dt.Columns.Count - 1 Then
                                dr(x) = rows(i).Split(Chr(9))(x)
                            Else
                                Throw New Exception("The number of columns on row " & i + If(header, 0, 1) & " is greater than the amount of columns in the " & If(header, "header.", "first row."))
                            End If
                        Next

                        dt.Rows.Add(dr)
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next

                    Dim DS As New DataSet
                    dt.TableName = tableName
                    DS.Tables.Add(dt)
                    Return DS
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.TSV")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

            ''' <summary>
            ''' Class import dengan data yang telah di dekripsi
            ''' </summary>
            Public Class Decrypt

                ''' <summary>
                ''' Class dekripsi DES
                ''' </summary>
                Public Class DES
                    Private des As New crypt.DES
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
                    ''' Import tsv ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String, Optional ByVal header As Boolean = True,
                                                Optional showException As Boolean = True) As DataTable
                        Try
                            Dim source As String = String.Empty
                            Dim dt As DataTable = New DataTable

                            If IO.File.Exists(pathFile) Then
                                source = IO.File.ReadAllText(pathFile)
                            Else
                                Throw New IO.FileNotFoundException("Could not find the file at " & pathFile, pathFile)
                            End If

                            Dim rows() As String = source.Split({Environment.NewLine}, StringSplitOptions.None)

                            For i As Integer = 0 To rows(0).Split(Chr(9)).Length - 1
                                Dim column As String = des.Decode(rows(0).Split(Chr(9))(i), secretKey)
                                dt.Columns.Add(If(header, column, "Col " & i + 1))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For i As Integer = If(header, 1, 0) To rows.Length - 1
                                Dim dr As DataRow = dt.NewRow

                                For x As Integer = 0 To rows(i).Split(Chr(9)).Length - 1
                                    If x <= dt.Columns.Count - 1 Then
                                        dr(x) = des.Decode(rows(i).Split(Chr(9))(x), secretKey)
                                    Else
                                        Throw New Exception("The number of columns on row " & i + If(header, 0, 1) & " is greater than the amount of columns in the " & If(header, "header.", "first row."))
                                    End If
                                Next

                                dt.Rows.Add(dr)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return dt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.TSV.Decrypt.DES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import tsv ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String,
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim source As String = String.Empty
                            Dim dt As DataTable = New DataTable

                            If IO.File.Exists(pathFile) Then
                                source = IO.File.ReadAllText(pathFile)
                            Else
                                Throw New IO.FileNotFoundException("Could not find the file at " & pathFile, pathFile)
                            End If

                            Dim rows() As String = source.Split({Environment.NewLine}, StringSplitOptions.None)

                            For i As Integer = 0 To rows(0).Split(Chr(9)).Length - 1
                                Dim column As String = des.Decode(rows(0).Split(Chr(9))(i), secretKey)
                                dt.Columns.Add(If(header, column, "Col " & i + 1))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For i As Integer = If(header, 1, 0) To rows.Length - 1
                                Dim dr As DataRow = dt.NewRow

                                For x As Integer = 0 To rows(i).Split(Chr(9)).Length - 1
                                    If x <= dt.Columns.Count - 1 Then
                                        dr(x) = des.Decode(rows(i).Split(Chr(9))(x), secretKey)
                                    Else
                                        Throw New Exception("The number of columns on row " & i + If(header, 0, 1) & " is greater than the amount of columns in the " & If(header, "header.", "first row."))
                                    End If
                                Next

                                dt.Rows.Add(dr)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Dim DS As New DataSet
                            dt.TableName = tableName
                            DS.Tables.Add(dt)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.TSV.Decrypt.DES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class

                ''' <summary>
                ''' Class dekripsi 3DES
                ''' </summary>
                Public Class TripleDES
                    Private tripledes As New crypt.TripleDES
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
                    ''' Import tsv ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String,
                                                Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim source As String = String.Empty
                            Dim dt As DataTable = New DataTable

                            If IO.File.Exists(pathFile) Then
                                source = IO.File.ReadAllText(pathFile)
                            Else
                                Throw New IO.FileNotFoundException("Could not find the file at " & pathFile, pathFile)
                            End If

                            Dim rows() As String = source.Split({Environment.NewLine}, StringSplitOptions.None)

                            For i As Integer = 0 To rows(0).Split(Chr(9)).Length - 1
                                Dim column As String = tripledes.Decode(rows(0).Split(Chr(9))(i), secretKey)
                                dt.Columns.Add(If(header, column, "Col " & i + 1))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For i As Integer = If(header, 1, 0) To rows.Length - 1
                                Dim dr As DataRow = dt.NewRow

                                For x As Integer = 0 To rows(i).Split(Chr(9)).Length - 1
                                    If x <= dt.Columns.Count - 1 Then
                                        dr(x) = tripledes.Decode(rows(i).Split(Chr(9))(x), secretKey)
                                    Else
                                        Throw New Exception("The number of columns on row " & i + If(header, 0, 1) & " is greater than the amount of columns in the " & If(header, "header.", "first row."))
                                    End If
                                Next

                                dt.Rows.Add(dr)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return dt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.TSV.Decrypt.TripleDES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import tsv ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String,
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim source As String = String.Empty
                            Dim dt As DataTable = New DataTable

                            If IO.File.Exists(pathFile) Then
                                source = IO.File.ReadAllText(pathFile)
                            Else
                                Throw New IO.FileNotFoundException("Could not find the file at " & pathFile, pathFile)
                            End If

                            Dim rows() As String = source.Split({Environment.NewLine}, StringSplitOptions.None)

                            For i As Integer = 0 To rows(0).Split(Chr(9)).Length - 1
                                Dim column As String = tripledes.Decode(rows(0).Split(Chr(9))(i), secretKey)
                                dt.Columns.Add(If(header, column, "Col " & i + 1))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For i As Integer = If(header, 1, 0) To rows.Length - 1
                                Dim dr As DataRow = dt.NewRow

                                For x As Integer = 0 To rows(i).Split(Chr(9)).Length - 1
                                    If x <= dt.Columns.Count - 1 Then
                                        dr(x) = tripledes.Decode(rows(i).Split(Chr(9))(x), secretKey)
                                    Else
                                        Throw New Exception("The number of columns on row " & i + If(header, 0, 1) & " is greater than the amount of columns in the " & If(header, "header.", "first row."))
                                    End If
                                Next

                                dt.Rows.Add(dr)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Dim DS As New DataSet
                            dt.TableName = tableName
                            DS.Tables.Add(dt)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.TSV.Decrypt.TripleDES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class

                ''' <summary>
                ''' Class dekripsi AES
                ''' </summary>
                Public Class AES
                    Private aes As New crypt.AES
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
                    ''' Import tsv ke datatable
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal pathFile As String, secretKey As String,
                                                Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim source As String = String.Empty
                            Dim dt As DataTable = New DataTable

                            If IO.File.Exists(pathFile) Then
                                source = IO.File.ReadAllText(pathFile)
                            Else
                                Throw New IO.FileNotFoundException("Could not find the file at " & pathFile, pathFile)
                            End If

                            Dim rows() As String = source.Split({Environment.NewLine}, StringSplitOptions.None)

                            For i As Integer = 0 To rows(0).Split(Chr(9)).Length - 1
                                Dim column As String = aes.Decode(rows(0).Split(Chr(9))(i), secretKey)
                                dt.Columns.Add(If(header, column, "Col " & i + 1))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For i As Integer = If(header, 1, 0) To rows.Length - 1
                                Dim dr As DataRow = dt.NewRow

                                For x As Integer = 0 To rows(i).Split(Chr(9)).Length - 1
                                    If x <= dt.Columns.Count - 1 Then
                                        dr(x) = aes.Decode(rows(i).Split(Chr(9))(x), secretKey)
                                    Else
                                        Throw New Exception("The number of columns on row " & i + If(header, 0, 1) & " is greater than the amount of columns in the " & If(header, "header.", "first row."))
                                    End If
                                Next

                                dt.Rows.Add(dr)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return dt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.TSV.Decrypt.AES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import tsv ke dataset
                    ''' </summary>
                    ''' <param name="pathFile">Lokasi path file txt</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="header">True or False</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(ByVal pathFile As String, tableName As String, secretKey As String,
                                              Optional ByVal header As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim source As String = String.Empty
                            Dim dt As DataTable = New DataTable

                            If IO.File.Exists(pathFile) Then
                                source = IO.File.ReadAllText(pathFile)
                            Else
                                Throw New IO.FileNotFoundException("Could not find the file at " & pathFile, pathFile)
                            End If

                            Dim rows() As String = source.Split({Environment.NewLine}, StringSplitOptions.None)

                            For i As Integer = 0 To rows(0).Split(Chr(9)).Length - 1
                                Dim column As String = aes.Decode(rows(0).Split(Chr(9))(i), secretKey)
                                dt.Columns.Add(If(header, column, "Col " & i + 1))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For i As Integer = If(header, 1, 0) To rows.Length - 1
                                Dim dr As DataRow = dt.NewRow

                                For x As Integer = 0 To rows(i).Split(Chr(9)).Length - 1
                                    If x <= dt.Columns.Count - 1 Then
                                        dr(x) = aes.Decode(rows(i).Split(Chr(9))(x), secretKey)
                                    Else
                                        Throw New Exception("The number of columns on row " & i + If(header, 0, 1) & " is greater than the amount of columns in the " & If(header, "header.", "first row."))
                                    End If
                                Next

                                dt.Rows.Add(dr)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Dim DS As New DataSet
                            dt.TableName = tableName
                            DS.Tables.Add(dt)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.TSV.Decrypt.AES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class
            End Class

        End Class

        ''' <summary>
        ''' Class import dari objek excel
        ''' </summary>
        Public Class Excell

#Region "DataSet"

            ''' <summary>
            ''' Import Excell ke Dataset
            ''' </summary>
            ''' <param name="fileName">Lokasi path file Excell</param>
            ''' <param name="hasHeaders">Gunakan header?</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataSet</returns>
            ''' <remarks>support return dataset ke datatable</remarks>
            Public Function ToDataSet(ByVal fileName As String, Optional ByVal hasHeaders As Boolean = True,
                                      Optional showException As Boolean = True) As DataSet
                Try
                    Dim stream As IO.FileStream = IO.File.Open(fileName, IO.FileMode.Open, IO.FileAccess.Read)
                    Dim excelReader As Excel.IExcelDataReader
                    If IO.Path.GetExtension(fileName) = ".xls" Then
                        excelReader = Excel.ExcelReaderFactory.CreateBinaryReader(stream)
                    Else
                        excelReader = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream)
                    End If

                    excelReader.IsFirstRowAsColumnNames = hasHeaders
                    Dim result As DataSet = excelReader.AsDataSet()

                    Return result
                    excelReader.Close()
                    stream.Close()
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Import Excell ke Dataset using OleDB engine
            ''' </summary>
            ''' <param name="fileName">Lokasi path file Excell</param>
            ''' <param name="hasHeaders">Gunakan header?</param>
            ''' <param name="avoidCrash">Hindari Crash?</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataSet</returns>
            ''' <remarks>support return dataset ke datatable</remarks>
            Public Function ToDataSetOleDB(ByVal fileName As String, Optional ByVal hasHeaders As Boolean = True,
                                      Optional avoidCrash As Boolean = True, Optional showException As Boolean = True) As DataSet
                Try
                    Dim HDR As String = If(hasHeaders, "Yes", "No")
                    Dim useIMEX As String = If(avoidCrash, "1", "0")
                    Dim strConn As String
                    If fileName.Substring(fileName.LastIndexOf("."c)).ToLower() = ".xlsx" Then
                        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileName & ";Extended Properties=""Excel 12.0;HDR=" & HDR & ";IMEX=" & useIMEX & """"
                    Else
                        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fileName & ";Extended Properties=""Excel 8.0;HDR=" & HDR & ";IMEX=" & useIMEX & """"
                    End If

                    Dim DS As New DataSet()

                    Using oleDBconn As New OleDbConnection(strConn)
                        oleDBconn.Open()

                        Dim schemaTable As DataTable = oleDBconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

                        For Each schemaRow As DataRow In schemaTable.Rows
                            Dim sheet As String = schemaRow("TABLE_NAME").ToString()

                            If Not sheet.EndsWith("_") Then
                                Try
                                    Dim cmd As New OleDbCommand("SELECT * FROM [" & sheet & "]", oleDBconn)
                                    cmd.CommandType = CommandType.Text

                                    Dim DSTable As New DataTable(sheet)
                                    DS.Tables.Add(DSTable)
                                    Dim adp As New OleDbDataAdapter(cmd)
                                    adp.Fill(DSTable)
                                Catch ex As Exception
                                    Throw New Exception(ex.Message + String.Format("Sheet:{0}.File:F{1}", sheet, fileName), ex)
                                End Try
                            End If
                        Next
                        Return DS
                        oleDBconn.Close()
                    End Using
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

#Region "DataTable"
            ''' <summary>
            ''' Import Excell ke DataTable
            ''' </summary>
            ''' <param name="fileName">Lokasi path file Excell</param>
            ''' <param name="hasHeaders">Gunakan header?</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataTable</returns>
            Public Function ToDataTable(ByVal fileName As String, Optional ByVal hasHeaders As Boolean = True,
                                        Optional showException As Boolean = True) As DataTable
                Try
                    Dim stream As IO.FileStream = IO.File.Open(fileName, IO.FileMode.Open, IO.FileAccess.Read)
                    Dim excelReader As Excel.IExcelDataReader
                    If IO.Path.GetExtension(fileName) = ".xls" Then
                        excelReader = Excel.ExcelReaderFactory.CreateBinaryReader(stream)
                    Else
                        excelReader = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream)
                    End If

                    excelReader.IsFirstRowAsColumnNames = hasHeaders
                    Dim result As DataSet = excelReader.AsDataSet()

                    Dim DT As New DataTable
                    DT = result.Tables(0)
                    Return DT
                    excelReader.Close()
                    stream.Close()
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Import Excell ke DataTable using OleDB engine
            ''' </summary>
            ''' <param name="fileName">Lokasi path file Excell</param>
            ''' <param name="hasHeaders">Gunakan header?</param>
            ''' <param name="avoidCrash">Hindari Crash?</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataTable</returns>
            Public Function ToDataTableOleDB(ByVal fileName As String, Optional ByVal hasHeaders As Boolean = True,
                                        Optional avoidCrash As Boolean = True, Optional showException As Boolean = True) As DataTable
                Try
                    Dim HDR As String = If(hasHeaders, "Yes", "No")
                    Dim useIMEX As String = If(avoidCrash, "1", "0")
                    Dim strConn As String
                    If fileName.Substring(fileName.LastIndexOf("."c)).ToLower() = ".xlsx" Then
                        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileName & ";Extended Properties=""Excel 12.0;HDR=" & HDR & ";IMEX=" & useIMEX & """"
                    Else
                        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fileName & ";Extended Properties=""Excel 8.0;HDR=" & HDR & ";IMEX=" & useIMEX & """"
                    End If

                    Dim DS As New DataSet()

                    Using oleDBconn As New OleDbConnection(strConn)
                        oleDBconn.Open()

                        Dim schemaTable As DataTable = oleDBconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

                        For Each schemaRow As DataRow In schemaTable.Rows
                            Dim sheet As String = schemaRow("TABLE_NAME").ToString()

                            If Not sheet.EndsWith("_") Then
                                Try
                                    Dim cmd As New OleDbCommand("SELECT * FROM [" & sheet & "]", oleDBconn)
                                    cmd.CommandType = CommandType.Text

                                    Dim DSTable As New DataTable(sheet)
                                    DS.Tables.Add(DSTable)
                                    Dim adp As New OleDbDataAdapter(cmd)
                                    adp.Fill(DSTable)
                                Catch ex As Exception
                                    Throw New Exception(ex.Message + String.Format("Sheet:{0}.File:F{1}", sheet, fileName), ex)
                                End Try
                            End If
                        Next

                        Dim DT As New DataTable
                        DT = DS.Tables(0)
                        Return DT
                        oleDBconn.Close()
                    End Using
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

            ''' <summary>
            ''' Class import dengan data yang telah di dekripsi
            ''' </summary>
            Public Class Decrypt

                ''' <summary>
                ''' Class dekripsi DES
                ''' </summary>
                Public Class DES
                    Private des As New crypt.DES
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
                    ''' Import Excell ke Dataset
                    ''' </summary>
                    ''' <param name="fileName">Lokasi path file Excell</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="hasHeaders">Gunakan header?</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    ''' <remarks>support return dataset ke datatable</remarks>
                    Public Function ToDataSet(ByVal fileName As String, tableName As String, secretKey As String, Optional ByVal hasHeaders As Boolean = True,
                                              Optional showException As Boolean = True) As DataSet
                        Try
                            Dim ex As New Excell
                            Dim DT As New DataTable
                            DT = ex.ToDataTable(fileName, hasHeaders, showException)

                            'Proses decrypt
                            Dim DTDecrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(des.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = des.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Dim DS As New DataSet
                            DTDecrypt.TableName = tableName
                            DS.Tables.Add(DTDecrypt)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell.Decrypt.DES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import Excell ke Datatable
                    ''' </summary>
                    ''' <param name="fileName">Lokasi path file Excell</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="hasHeaders">Gunakan header?</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal fileName As String, secretKey As String, Optional ByVal hasHeaders As Boolean = True,
                                                Optional showException As Boolean = True) As DataTable
                        Try
                            Dim ex As New Excell
                            Dim DT As New DataTable
                            DT = ex.ToDataTable(fileName, hasHeaders, showException)

                            'Proses decrypt
                            Dim DTDecrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(des.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = des.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return DTDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell.Decrypt.DES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class

                ''' <summary>
                ''' Class dekripsi 3DES
                ''' </summary>
                Public Class TripleDES
                    Private tripledes As New crypt.TripleDES
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
                    ''' Import Excell ke Dataset
                    ''' </summary>
                    ''' <param name="fileName">Lokasi path file Excell</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="hasHeaders">Gunakan header?</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    ''' <remarks>support return dataset ke datatable</remarks>
                    Public Function ToDataSet(ByVal fileName As String, tableName As String, secretKey As String, Optional ByVal hasHeaders As Boolean = True,
                                              Optional showException As Boolean = True) As DataSet
                        Try
                            Dim ex As New Excell
                            Dim DT As New DataTable
                            DT = ex.ToDataTable(fileName, hasHeaders, showException)

                            'Proses decrypt
                            Dim DTDecrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(tripledes.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = tripledes.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Dim DS As New DataSet
                            DTDecrypt.TableName = tableName
                            DS.Tables.Add(DTDecrypt)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell.Decrypt.TripleDES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import Excell ke Datatable
                    ''' </summary>
                    ''' <param name="fileName">Lokasi path file Excell</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="hasHeaders">Gunakan header?</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal fileName As String, secretKey As String, Optional ByVal hasHeaders As Boolean = True,
                                                Optional showException As Boolean = True) As DataTable
                        Try
                            Dim ex As New Excell
                            Dim DT As New DataTable
                            DT = ex.ToDataTable(fileName, hasHeaders, showException)

                            'Proses decrypt
                            Dim DTDecrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(tripledes.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = tripledes.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return DTDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell.Decrypt.TripleDES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class

                ''' <summary>
                ''' Class dekripsi AES
                ''' </summary>
                Public Class AES
                    Private aes As New crypt.AES
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
                    ''' Import Excell ke Dataset
                    ''' </summary>
                    ''' <param name="fileName">Lokasi path file Excell</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="hasHeaders">Gunakan header?</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    ''' <remarks>support return dataset ke datatable</remarks>
                    Public Function ToDataSet(ByVal fileName As String, tableName As String, secretKey As String, Optional ByVal hasHeaders As Boolean = True,
                                              Optional showException As Boolean = True) As DataSet
                        Try
                            Dim ex As New Excell
                            Dim DT As New DataTable
                            DT = ex.ToDataTable(fileName, hasHeaders, showException)

                            'Proses decrypt
                            Dim DTDecrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(aes.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = aes.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Dim DS As New DataSet
                            DTDecrypt.TableName = tableName
                            DS.Tables.Add(DTDecrypt)
                            Return DS
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell.Decrypt.AES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import Excell ke Datatable
                    ''' </summary>
                    ''' <param name="fileName">Lokasi path file Excell</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="hasHeaders">Gunakan header?</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(ByVal fileName As String, secretKey As String, Optional ByVal hasHeaders As Boolean = True,
                                                Optional showException As Boolean = True) As DataTable
                        Try
                            Dim ex As New Excell
                            Dim DT As New DataTable
                            DT = ex.ToDataTable(fileName, hasHeaders, showException)

                            'Proses decrypt
                            Dim DTDecrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(aes.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = aes.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return DTDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.Excell.Decrypt.AES")
                            Return Nothing
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                End Class
            End Class

        End Class

        ''' <summary>
        ''' Class import dari objek xml
        ''' </summary>
        Public Class XML

#Region "DataSet"
            ''' <summary>
            ''' Import XML ke Dataset
            ''' </summary>
            ''' <param name="pathFile">Path File XML</param>
            ''' <param name="readSchema">Gunakan schema? Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataSet</returns>
            Public Function ToDataSet(pathFile As String, Optional readSchema As Boolean = True, Optional showException As Boolean = True) As DataSet
                Try
                    Dim DS As New DataSet
                    If readSchema = True Then
                        DS.ReadXml(pathFile, XmlReadMode.ReadSchema)
                    Else
                        DS.ReadXml(pathFile, XmlReadMode.Auto)
                    End If
                    Return DS
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.XML")
                    Return Nothing
                End Try
            End Function
#End Region

#Region "DataTable"
            ''' <summary>
            ''' Import XML ke Datatable
            ''' </summary>
            ''' <param name="pathFile">Path File XML</param>
            ''' <param name="readSchema">Gunakan schema? Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataTable</returns>
            Public Function ToDataTable(pathFile As String, Optional readSchema As Boolean = True, Optional showException As Boolean = True) As DataTable
                Try
                    Dim DS As New DataSet
                    If readSchema = True Then
                        DS.ReadXml(pathFile, XmlReadMode.ReadSchema)
                    Else
                        DS.ReadXml(pathFile, XmlReadMode.Auto)
                    End If
                    Dim DT As New DataTable
                    DT = DS.Tables(0)
                    Return DT
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.XML")
                    Return Nothing
                End Try
            End Function
#End Region

            ''' <summary>
            ''' Class import dengan data yang telah di dekripsi
            ''' </summary>
            Public Class Decrypt

                ''' <summary>
                ''' Class dekripsi DES
                ''' </summary>
                Public Class DES
                    Private des As New crypt.DES
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
                    ''' Import XML ke DataTable
                    ''' </summary>
                    ''' <param name="pathFile">Path File XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="readSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(pathFile As String, secretKey As String,
                                                Optional readSchema As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DS As New DataSet
                            'baca file
                            If readSchema = True Then
                                DS.ReadXml(pathFile, XmlReadMode.ReadSchema)
                            Else
                                DS.ReadXml(pathFile, XmlReadMode.Auto)
                            End If

                            'proses decrypt
                            Dim DT, DTDecrypt As New DataTable
                            DT = DS.Tables(0)

                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(des.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = des.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return DTDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.XML.Decrypt.DES")
                            Return Nothing
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import XML ke DataSet
                    ''' </summary>
                    ''' <param name="pathFile">Path File XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="readSchema">Gunakan schema? Default = True</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(pathFile As String, tableName As String, secretKey As String,
                                              Optional readSchema As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DS As New DataSet

                            'baca file
                            If readSchema = True Then
                                DS.ReadXml(pathFile, XmlReadMode.ReadSchema)
                            Else
                                DS.ReadXml(pathFile, XmlReadMode.Auto)
                            End If


                            'proses decrypt dan konversi ke datatable
                            Dim DT, DTDecrypt As New DataTable
                            DT = DS.Tables(0)

                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(des.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = des.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses mengkonversi kembali ke dataset
                            Dim DSDecrypt As New DataSet
                            DTDecrypt.TableName = tableName
                            DSDecrypt.Tables.Add(DTDecrypt)
                            Return DSDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.XML.Decrypt.DES")
                            Return Nothing
                        End Try
                    End Function
                End Class

                ''' <summary>
                ''' Class dekripsi 3DES
                ''' </summary>
                Public Class TripleDES
                    Private tripledes As New crypt.TripleDES
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
                    ''' Import XML ke DataTable
                    ''' </summary>
                    ''' <param name="pathFile">Path File XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="readSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(pathFile As String, secretKey As String,
                                                Optional readSchema As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DS As New DataSet
                            'baca file
                            If readSchema = True Then
                                DS.ReadXml(pathFile, XmlReadMode.ReadSchema)
                            Else
                                DS.ReadXml(pathFile, XmlReadMode.Auto)
                            End If

                            'proses decrypt
                            Dim DT, DTDecrypt As New DataTable
                            DT = DS.Tables(0)

                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(tripledes.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = tripledes.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return DTDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.XML.Decrypt.TripleDES")
                            Return Nothing
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import XML ke DataSet
                    ''' </summary>
                    ''' <param name="pathFile">Path File XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="readSchema">Gunakan schema? Default = True</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(pathFile As String, tableName As String, secretKey As String,
                                              Optional readSchema As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DS As New DataSet

                            'baca file
                            If readSchema = True Then
                                DS.ReadXml(pathFile, XmlReadMode.ReadSchema)
                            Else
                                DS.ReadXml(pathFile, XmlReadMode.Auto)
                            End If

                            'proses decrypt dan konversi ke datatable
                            Dim DT, DTDecrypt As New DataTable
                            DT = DS.Tables(0)

                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(tripledes.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = tripledes.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses mengkonversi kembali ke dataset
                            Dim DSDecrypt As New DataSet
                            DTDecrypt.TableName = tableName
                            DSDecrypt.Tables.Add(DTDecrypt)
                            Return DSDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.XML.Decrypt.TripleDES")
                            Return Nothing
                        End Try
                    End Function
                End Class

                ''' <summary>
                ''' Class dekripsi AES
                ''' </summary>
                Public Class AES
                    Private aes As New crypt.AES
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
                    ''' Import XML ke DataTable
                    ''' </summary>
                    ''' <param name="pathFile">Path File XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="readSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDataTable(pathFile As String, secretKey As String,
                                                Optional readSchema As Boolean = True, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim DS As New DataSet
                            'baca file
                            If readSchema = True Then
                                DS.ReadXml(pathFile, XmlReadMode.ReadSchema)
                            Else
                                DS.ReadXml(pathFile, XmlReadMode.Auto)
                            End If

                            'proses decrypt
                            Dim DT, DTDecrypt As New DataTable
                            DT = DS.Tables(0)

                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(aes.Decode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = aes.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            Return DTDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.XML.Decrypt.AES")
                            Return Nothing
                        End Try
                    End Function

                    ''' <summary>
                    ''' Import XML ke DataSet
                    ''' </summary>
                    ''' <param name="pathFile">Path File XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="readSchema">Gunakan schema? Default = True</param>
                    ''' <param name="tableName">Nama table DataSet</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataSet(pathFile As String, tableName As String, secretKey As String,
                                              Optional readSchema As Boolean = True, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim DS As New DataSet

                            'baca file
                            If readSchema = True Then
                                DS.ReadXml(pathFile, XmlReadMode.ReadSchema)
                            Else
                                DS.ReadXml(pathFile, XmlReadMode.Auto)
                            End If

                            'proses decrypt dan konversi ke datatable
                            Dim DT, DTDecrypt As New DataTable
                            DT = DS.Tables(0)

                            For Each col As DataColumn In DT.Columns
                                DTDecrypt.Columns.Add(aes.Decode(col.ToString, secretKey))
                            Next

                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTDecrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = aes.Decode(dr(i).ToString, secretKey)
                                Next
                                DTDecrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses mengkonversi kembali ke dataset
                            Dim DSDecrypt As New DataSet
                            DTDecrypt.TableName = tableName
                            DSDecrypt.Tables.Add(DTDecrypt)
                            Return DSDecrypt
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.XML.Decrypt.AES")
                            Return Nothing
                        End Try
                    End Function
                End Class

            End Class

        End Class

        ''' <summary>
        ''' Class import dari objek json
        ''' </summary>
        Public Class JSON

#Region "Dataset"
            ''' <summary>
            ''' Import JSON ke Dataset
            ''' </summary>
            ''' <param name="pathFile">Full path lokasi file JSON</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Dataset</returns>
            Public Function ToDataSet(pathFile As String, Optional showException As Boolean = True) As DataSet
                Try
                    Dim dataJSON As String = System.IO.File.ReadAllText(pathFile)
                    If dataJSON <> Nothing Then
                        Return Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataSet)(dataJSON)
                    Else
                        Return Nothing
                    End If
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.JSON")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

#Region "Datatable"
            ''' <summary>
            ''' Import JSON ke Datatable
            ''' </summary>
            ''' <param name="pathFile">Full path lokasi file JSON</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataTable</returns>
            Public Function ToDataTable(pathFile As String, Optional showException As Boolean = True) As DataTable
                Try
                    Dim dataJSON As String = System.IO.File.ReadAllText(pathFile)
                    If dataJSON <> Nothing Then
                        Return Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataTable)(dataJSON)
                    Else
                        Return Nothing
                    End If
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.JSON")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

            ''' <summary>
            ''' Class import dengan data yang telah di dekripsi
            ''' </summary>
            Public Class Decrypt

                ''' <summary>
                ''' Class dekripsi DES
                ''' </summary>
                Public Class DES
                    Private des As New crypt.DES
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
                    ''' Deserialize dan mengisi data JSON ke DataTable
                    ''' </summary>
                    ''' <param name="pathFile">Full path lokasi file JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDatatable(pathFile As String, secretKey As String, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim dataJSON As String = System.IO.File.ReadAllText(pathFile)

                            If dataJSON <> Nothing Then
                                Dim DT As New DataTable

                                DT = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataTable)(dataJSON)

                                'Proses decrypt
                                Dim DTDecrypt As New DataTable
                                For Each col As DataColumn In DT.Columns
                                    DTDecrypt.Columns.Add(des.Decode(col.ToString, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'menyimpan data
                                For Each dr As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTDecrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = des.Decode(dr(i).ToString, secretKey)
                                    Next
                                    DTDecrypt.Rows.Add(drNew)
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                Return DTDecrypt
                            Else
                                Return Nothing
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.JSON.Decrypt.DES")
                            Return Nothing
                        End Try
                    End Function

                    ''' <summary>
                    ''' Deserialize dan mengisi data JSON ke DataSet
                    ''' </summary>
                    ''' <param name="pathFile">Full path lokasi file JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama tabel DataSet</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataset(pathFile As String, tableName As String, secretKey As String, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim dataJSON As String = System.IO.File.ReadAllText(pathFile)

                            If dataJSON <> Nothing Then
                                Dim DS As New DataSet

                                DS = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataSet)(dataJSON)

                                Dim DT As New DataTable
                                DT = DS.Tables(0)

                                'Proses decrypt
                                Dim DTDecrypt As New DataTable
                                For Each col As DataColumn In DT.Columns
                                    DTDecrypt.Columns.Add(des.Decode(col.ToString, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'menyimpan data
                                For Each dr As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTDecrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = des.Decode(dr(i).ToString, secretKey)
                                    Next
                                    DTDecrypt.Rows.Add(drNew)
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next

                                Dim DSDecrypt As New DataSet
                                DTDecrypt.TableName = tableName
                                DSDecrypt.Tables.Add(DTDecrypt)
                                Return DSDecrypt
                            Else
                                Return Nothing
                            End If

                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.JSON.Decrypt.DES")
                            Return Nothing
                        End Try
                    End Function
                End Class

                ''' <summary>
                ''' Class dekripsi 3DES
                ''' </summary>
                Public Class TripleDES
                    Private tripledes As New crypt.TripleDES
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
                    ''' Deserialize dan mengisi data JSON ke DataTable
                    ''' </summary>
                    ''' <param name="pathFile">Full path lokasi file JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDatatable(pathFile As String, secretKey As String, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim dataJSON As String = System.IO.File.ReadAllText(pathFile)

                            If dataJSON <> Nothing Then
                                Dim DT As New DataTable

                                DT = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataTable)(dataJSON)

                                'Proses decrypt
                                Dim DTDecrypt As New DataTable
                                For Each col As DataColumn In DT.Columns
                                    DTDecrypt.Columns.Add(tripledes.Decode(col.ToString, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'menyimpan data
                                For Each dr As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTDecrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = tripledes.Decode(dr(i).ToString, secretKey)
                                    Next
                                    DTDecrypt.Rows.Add(drNew)
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                Return DTDecrypt
                            Else
                                Return Nothing
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.JSON.Decrypt.TripleDES")
                            Return Nothing
                        End Try
                    End Function

                    ''' <summary>
                    ''' Deserialize dan mengisi data JSON ke DataSet
                    ''' </summary>
                    ''' <param name="pathFile">Full path lokasi file JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama tabel DataSet</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataset(pathFile As String, tableName As String, secretKey As String, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim dataJSON As String = System.IO.File.ReadAllText(pathFile)

                            If dataJSON <> Nothing Then
                                Dim DS As New DataSet

                                DS = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataSet)(dataJSON)

                                Dim DT As New DataTable
                                DT = DS.Tables(0)

                                'Proses decrypt
                                Dim DTDecrypt As New DataTable
                                For Each col As DataColumn In DT.Columns
                                    DTDecrypt.Columns.Add(tripledes.Decode(col.ToString, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'menyimpan data
                                For Each dr As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTDecrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = tripledes.Decode(dr(i).ToString, secretKey)
                                    Next
                                    DTDecrypt.Rows.Add(drNew)
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next

                                Dim DSDecrypt As New DataSet
                                DTDecrypt.TableName = tableName
                                DSDecrypt.Tables.Add(DTDecrypt)
                                Return DSDecrypt
                            Else
                                Return Nothing
                            End If

                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.JSON.Decrypt.TripleDES")
                            Return Nothing
                        End Try
                    End Function
                End Class

                ''' <summary>
                ''' Class dekripsi AES
                ''' </summary>
                Public Class AES
                    Private aes As New crypt.AES
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
                    ''' Deserialize dan mengisi data JSON ke DataTable
                    ''' </summary>
                    ''' <param name="pathFile">Full path lokasi file JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataTable</returns>
                    Public Function ToDatatable(pathFile As String, secretKey As String, Optional showException As Boolean = True) As DataTable
                        Try
                            Dim dataJSON As String = System.IO.File.ReadAllText(pathFile)

                            If dataJSON <> Nothing Then
                                Dim DT As New DataTable

                                DT = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataTable)(dataJSON)

                                'Proses decrypt
                                Dim DTDecrypt As New DataTable
                                For Each col As DataColumn In DT.Columns
                                    DTDecrypt.Columns.Add(aes.Decode(col.ToString, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'menyimpan data
                                For Each dr As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTDecrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = aes.Decode(dr(i).ToString, secretKey)
                                    Next
                                    DTDecrypt.Rows.Add(drNew)
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                Return DTDecrypt
                            Else
                                Return Nothing
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.JSON.Decrypt.AES")
                            Return Nothing
                        End Try
                    End Function

                    ''' <summary>
                    ''' Deserialize dan mengisi data JSON ke DataSet
                    ''' </summary>
                    ''' <param name="pathFile">Full path lokasi file JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama tabel DataSet</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>DataSet</returns>
                    Public Function ToDataset(pathFile As String, tableName As String, secretKey As String, Optional showException As Boolean = True) As DataSet
                        Try
                            Dim dataJSON As String = System.IO.File.ReadAllText(pathFile)

                            If dataJSON <> Nothing Then
                                Dim DS As New DataSet

                                DS = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataSet)(dataJSON)

                                Dim DT As New DataTable
                                DT = DS.Tables(0)

                                'Proses decrypt
                                Dim DTDecrypt As New DataTable
                                For Each col As DataColumn In DT.Columns
                                    DTDecrypt.Columns.Add(aes.Decode(col.ToString, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'menyimpan data
                                For Each dr As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTDecrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = aes.Decode(dr(i).ToString, secretKey)
                                    Next
                                    DTDecrypt.Rows.Add(drNew)
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next

                                Dim DSDecrypt As New DataSet
                                DTDecrypt.TableName = tableName
                                DSDecrypt.Tables.Add(DTDecrypt)
                                Return DSDecrypt
                            Else
                                Return Nothing
                            End If

                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Import.JSON.Decrypt.AES")
                            Return Nothing
                        End Try
                    End Function
                End Class

            End Class
        End Class

    End Class
End Namespace