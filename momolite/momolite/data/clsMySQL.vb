
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

Imports Excels = Microsoft.Office.Interop.Excel
Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Namespace data
    ''' <summary>Class MySQL</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>21 August 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>2.9.0</version>
    ''' <requirement>
    ''' - Imports MySql.Data.MySqlClient
    ''' - Imports System.Windows.Forms
    ''' - Imports Excel = Microsoft.Office.Interop.Excel
    ''' </requirement>
    Public Class MySQL

        Private Shared sqlConn As MySqlConnection
        Private Shared _strConn As String
        Private _timeOut As Integer = 600
        Private _transaction As Boolean = True
        Private NameClass As String = "momolite.data.MySQL"
        Private _connectionString As String = Nothing
        Private _autoConnection As Boolean = True
        Private _iconnection As MySqlConnection = Nothing
        Private _connection As MySqlConnection = Nothing

#Region "Property"

        ''' <summary>
        ''' Shared Connection String Database. Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
        ''' </summary>
        ''' <value>Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"</value>
        ''' <returns>String</returns>
        Public Shared Property SharedConnectionString() As String
            Get
                Return _strConn
            End Get
            Set(ByVal value As String)
                _strConn = value
            End Set
        End Property

        ''' <summary>
        ''' Connection String Database. Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
        ''' </summary>
        ''' <value>Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"</value>
        ''' <returns>String</returns>
        Public Property ConnectionString() As String
            Get
                Return _connectionString
            End Get
            Set(ByVal value As String)
                _connectionString = value
            End Set
        End Property

        ''' <summary>
        ''' Gunakan Transaction agar lebih aman dalam eksekusi database
        ''' </summary>
        ''' <value>Default value is True</value>
        ''' <returns>Boolean</returns>
        Public Property UseTransaction As Boolean
            Get
                Return _transaction
            End Get
            Set(ByVal value As Boolean)
                If value <> _transaction Then
                    _transaction = value
                End If
            End Set
        End Property

        ''' <summary>
        ''' Connection TimeOut Data Adapter
        ''' </summary>
        ''' <value>Default value = 600</value>
        ''' <returns>Integer</returns>
        Public Property TimeOut() As Integer
            Get
                Return _timeOut
            End Get
            Set(ByVal value As Integer)
                _timeOut = value
            End Set
        End Property
#End Region

#Region "Property Manual"
        ''' <summary>
        ''' Original connection database
        ''' </summary>
        ''' <value>MySqlConnection objek</value>
        ''' <returns>MySqlConnection</returns>
        Public Property Connection() As MySqlConnection
            Get
                Return _connection
            End Get
            Set(ByVal value As MySqlConnection)
                _connection = value
            End Set
        End Property

        ''' <summary>
        ''' Gunakan Auto Connection database
        ''' </summary>
        ''' <value>Default value is True</value>
        ''' <returns>Boolean</returns>
        Public Property AutoConnection As Boolean
            Get
                Return _autoConnection
            End Get
            Set(ByVal value As Boolean)
                If value <> _autoConnection Then
                    _autoConnection = value
                End If
            End Set
        End Property
#End Region

        ''' <summary>
        ''' Class untuk membuat query database
        ''' </summary>
        Public Class BuildQuery
            Private NameClass As String = "momolite.data.MySQL.BuildQuery"

#Region "Property Connection String"
            Private _connectionString As String = Nothing
            ''' <summary>
            ''' Connection String Database. Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
            ''' </summary>
            ''' <value>Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"</value>
            ''' <returns>String</returns>
            Public Property ConnectionString() As String
                Get
                    Return _connectionString
                End Get
                Set(ByVal value As String)
                    _connectionString = value
                End Set
            End Property
#End Region

#Region "Property TimeOut"
            Private _timeOut As Integer = 600

            ''' <summary>
            ''' Connection TimeOut Data Adapter
            ''' </summary>
            ''' <value>Default value = 600</value>
            ''' <returns>Integer</returns>
            Public Property TimeOut() As Integer
                Get
                    Return _timeOut
                End Get
                Set(ByVal value As Integer)
                    _timeOut = value
                End Set
            End Property
#End Region

#Region "Property Manual"
            Private _autoConnection As Boolean = True
            Private _iconnection As MySqlConnection = Nothing
            Private _connection As MySqlConnection = Nothing

            ''' <summary>
            ''' Original connection database
            ''' </summary>
            ''' <value>MySqlConnection objek</value>
            ''' <returns>MySqlConnection</returns>
            Public Property Connection() As MySqlConnection
                Get
                    Return _connection
                End Get
                Set(ByVal value As MySqlConnection)
                    _connection = value
                End Set
            End Property

            ''' <summary>
            ''' Gunakan Auto Connection database
            ''' </summary>
            ''' <value>Default value is True</value>
            ''' <returns>Boolean</returns>
            Public Property AutoConnection As Boolean
                Get
                    Return _autoConnection
                End Get
                Set(ByVal value As Boolean)
                    If value <> _autoConnection Then
                        _autoConnection = value
                    End If
                End Set
            End Property
#End Region

            ''' <summary>
            ''' Membuat Query dan menyimpan hasilnya ke dalam memory datatable
            ''' </summary>
            ''' <param name="sqlQuery">SQL Query</param>
            ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
            ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataTable</returns>
            ''' 
            ''' <example>
            ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
            ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
            ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
            ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
            '''
            ''' Dim DB As New momolite.data.MySQL.BuildQuery 
            ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",{param1,param2})
            ''' 
            ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
            ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
            ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
            ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
            '''
            ''' Dim DB As New momolite.data.MySQL.BuildQuery 
            ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan", ,{param1,param2})
            ''' </example>
            Public Function ToDataTable(ByVal sqlQuery As String, Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                        Optional showException As Boolean = True) As DataTable
                Try

                    Dim DT As New DataTable

                    'Membuka koneksi database
                    If _autoConnection Then
                        If Connection Is Nothing Then
                            If ConnectionString <> Nothing Then
                                sqlConn = New MySqlConnection(ConnectionString)
                                sqlConn.Open()
                            Else
                                sqlConn = New MySqlConnection(SharedConnectionString)
                                sqlConn.Open()
                            End If
                        Else
                            sqlConn = Connection
                        End If
                    Else
                        sqlConn = _iconnection
                    End If

                    'Proses menjalankan dan menyimpan hasil Query ke datatable
                    Dim sqlDAdapter As New MySqlDataAdapter
                    sqlDAdapter = New MySqlDataAdapter(sqlQuery, sqlConn)
                    If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                        For i As Integer = 0 To paramAddWithValue.Length - 1
                            AdapterParameter(sqlDAdapter, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                        Next
                    ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                        For i As Integer = 0 To paramAdd.Length - 1
                            AdapterParameterAdvanced(sqlDAdapter, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                        Next
                    End If
                    sqlDAdapter.SelectCommand.CommandTimeout = TimeOut
                    sqlDAdapter.Fill(DT)
                    Return DT

                Catch ex As Exception
                    'Proses menampilkan pesan jika terjadi error
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                    Return Nothing

                Finally
                    'Menutup koneksi database
                    If Connection Is Nothing AndAlso _autoConnection Then
                        If sqlConn IsNot Nothing Then
                            sqlConn.Close()
                            sqlConn.Dispose()
                        End If
                    End If
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Membuat Query dan menyimpan hasilnya ke dalam memory dataset
            ''' </summary>
            ''' <param name="sqlQuery">SQL Query</param>
            ''' <param name="nameTableDS">Nama tabel dataset</param>
            ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
            ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>DataSet</returns>
            ''' 
            ''' <example>
            ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
            ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
            ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
            ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
            '''
            ''' Dim DB As New momolite.data.MySQL.BuildQuery 
            ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah",{param1,param2})
            ''' 
            ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
            ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
            ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
            ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
            '''
            ''' Dim DB As New momolite.data.MySQL.BuildQuery 
            ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah", ,{param1,param2})
            ''' </example>
            Public Function ToDataSet(ByVal sqlQuery As String, nameTableDS As String, Optional paramAddWithValue()() As Object = Nothing,
                                      Optional paramAdd()() As Object = Nothing, Optional showException As Boolean = True) As DataSet
                Try
                    Dim DS As New DataSet

                    'Membuka koneksi database
                    If _autoConnection Then
                        If Connection Is Nothing Then
                            If ConnectionString <> Nothing Then
                                sqlConn = New MySqlConnection(ConnectionString)
                                sqlConn.Open()
                            Else
                                sqlConn = New MySqlConnection(SharedConnectionString)
                                sqlConn.Open()
                            End If
                        Else
                            sqlConn = Connection
                        End If
                    Else
                        sqlConn = _iconnection
                    End If

                    'Proses menjalankan dan menyimpan hasil Query ke dataset
                    Dim sqlDAdapter As New MySqlDataAdapter
                    sqlDAdapter = New MySqlDataAdapter(sqlQuery, sqlConn)
                    If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                        For i As Integer = 0 To paramAddWithValue.Length - 1
                            AdapterParameter(sqlDAdapter, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                        Next
                    ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                        For i As Integer = 0 To paramAdd.Length - 1
                            AdapterParameterAdvanced(sqlDAdapter, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                        Next
                    End If
                    sqlDAdapter.SelectCommand.CommandTimeout = TimeOut
                    sqlDAdapter.Fill(DS, nameTableDS)
                    Return DS

                Catch ex As Exception
                    'Proses menampilkan pesan jika terjadi error
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                    Return Nothing

                Finally
                    'Menutup koneksi database
                    If Connection Is Nothing AndAlso _autoConnection Then
                        If sqlConn IsNot Nothing Then
                            sqlConn.Close()
                            sqlConn.Dispose()
                        End If
                    End If
                    GC.Collect()
                End Try
            End Function

#Region "Status"
            ''' <summary>
            ''' Check status koneksi database. [AutoConnection tidak akan menampilkan status yang sedang terjadi]
            ''' </summary>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            ''' <remarks>Untuk verifikasi status koneksi database</remarks>
            Public Function ConnectionStatus(Optional showException As Boolean = True) As Boolean
                Try
                    If _autoConnection Then
                        If Connection Is Nothing Then
                            If ConnectionString <> Nothing Then
                                sqlConn = New MySqlConnection(ConnectionString)
                                sqlConn.Open()
                            Else
                                sqlConn = New MySqlConnection(SharedConnectionString)
                                sqlConn.Open()
                            End If
                        Else
                            sqlConn = Connection
                        End If
                    Else
                        sqlConn = _iconnection
                    End If

                    If sqlConn.State = ConnectionState.Open Then Return True Else Return False
                Catch exs As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(exs, globals.Dev.Icons.Errors, NameClass)
                    Return False
                Finally
                    If Connection Is Nothing AndAlso _autoConnection Then
                        If sqlConn IsNot Nothing Then
                            sqlConn.Close()
                            sqlConn.Dispose()
                        End If
                    End If
                    GC.Collect()
                End Try
            End Function
#End Region

#Region "Manual Connection"
            ''' <summary>
            ''' Membuka koneksi database secara manual
            ''' </summary>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Sub OpenConnection(Optional showException As Boolean = True)
                Try
                    If _autoConnection = False Then
                        If ConnectionString <> Nothing Then
                            _iconnection = New MySqlConnection(ConnectionString)
                        Else
                            _iconnection = New MySqlConnection(SharedConnectionString)
                        End If
                        _iconnection.Open()
                    End If
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                    _iconnection = Nothing
                End Try
            End Sub

            ''' <summary>
            ''' Menutup koneksi database secara manual
            ''' </summary>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            Public Sub CloseConnection(Optional showException As Boolean = True)
                Try
                    If _autoConnection = False Then
                        If _iconnection IsNot Nothing Then
                            If _iconnection.State = ConnectionState.Open Then
                                _iconnection.Close()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                End Try
            End Sub
#End Region

            ''' <summary>
            ''' Class untuk membuat query database ke output secara langsung
            ''' </summary>
            Public Class Direct
                Private NameClass As String = "momolite.data.MySQL.BuildQuery.Direct"

#Region "Property Connection String"
                Private _connectionString As String = Nothing
                ''' <summary>
                ''' Connection String Database. Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' </summary>
                ''' <value>Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"</value>
                ''' <returns>String</returns>
                Public Property ConnectionString() As String
                    Get
                        Return _connectionString
                    End Get
                    Set(ByVal value As String)
                        _connectionString = value
                    End Set
                End Property
#End Region

#Region "Property TimeOut"
                Private _timeOut As Integer = 600

                ''' <summary>
                ''' Connection TimeOut Data Adapter
                ''' </summary>
                ''' <value>Default value = 600</value>
                ''' <returns>Integer</returns>
                Public Property TimeOut() As Integer
                    Get
                        Return _timeOut
                    End Get
                    Set(ByVal value As Integer)
                        _timeOut = value
                    End Set
                End Property
#End Region

#Region "Property Manual"
                Private _autoConnection As Boolean = True
                Private _iconnection As MySqlConnection = Nothing
                Private _connection As MySqlConnection = Nothing

                ''' <summary>
                ''' Original connection database
                ''' </summary>
                ''' <value>MySqlConnection objek</value>
                ''' <returns>MySqlConnection</returns>
                Public Property Connection() As MySqlConnection
                    Get
                        Return _connection
                    End Get
                    Set(ByVal value As MySqlConnection)
                        _connection = value
                    End Set
                End Property

                ''' <summary>
                ''' Gunakan Auto Connection database
                ''' </summary>
                ''' <value>Default value is True</value>
                ''' <returns>Boolean</returns>
                Public Property AutoConnection As Boolean
                    Get
                        Return _autoConnection
                    End Get
                    Set(ByVal value As Boolean)
                        If value <> _autoConnection Then
                            _autoConnection = value
                        End If
                    End Set
                End Property
#End Region

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya ke dalam memory dataset
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="nameTableDS">Nama tabel dataset</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Dataset</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah",{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah", ,{param1,param2})
                ''' </example>
                Public Function ToDataSet(ByVal sqlQuery As String, nameTableDS As String, Optional paramAddWithValue()() As Object = Nothing,
                                          Optional paramAdd()() As Object = Nothing, Optional showException As Boolean = True) As DataSet
                    Try

                        Dim DR As MySqlDataReader
                        Dim DS As New DataSet
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        DS.Load(DR, LoadOption.OverwriteChanges, nameTableDS)
                        DR.Close()
                        Return DS

                    Catch ex As Exception
                        'Proses menampilkan pesan jika terjadi error
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return Nothing

                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya ke dalam memory datatable
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Datatable</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan", ,{param1,param2})
                ''' </example>
                Public Function ToDataTable(ByVal sqlQuery As String, Optional paramAddWithValue()() As Object = Nothing,
                                            Optional paramAdd()() As Object = Nothing, Optional showException As Boolean = True) As DataTable
                    Try

                        Dim DR As MySqlDataReader
                        Dim DT As New DataTable
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        DT.Load(DR)
                        DR.Close()
                        Return DT

                    Catch ex As Exception
                        'Proses menampilkan pesan jika terjadi error
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return Nothing

                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

#Region "Output Object"
                ''' <summary>
                ''' Membuat Query secara direct stream ke listview
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="lv">nama objek ListView</param>
                ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",listView1,True, ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",listView1,True, , ,{param1,param2})
                ''' </example>
                Public Function ToListView(ByVal sqlQuery As String, lv As ListView, Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing,
                                           Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                    Try

                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()

                        With lv
                            .Clear()
                            .View = View.Details
                            .FullRowSelect = True
                            .GridLines = True
                            'buat kolom otomatis
                            .Columns.Clear()

                            'Buat header
                            For i As Integer = 0 To DR.FieldCount - 1
                                .Columns.Add(DR.GetName(i))
                            Next

                            'menambah data row
                            If DR.HasRows Then
                                Do While DR.Read()
                                    Dim item As New ListViewItem
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        item.Text = DR(0).ToString
                                    Else
                                        'Formatted datetime
                                        If TypeOf (DR(0)) Is DateTime Then
                                            item.Text = DirectCast(DR(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                        Else
                                            item.Text = DR(0).ToString
                                        End If
                                    End If

                                    For x As Integer = 1 To DR.FieldCount - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            item.SubItems.Add(DR(x).ToString)
                                        Else
                                            'Formatted datetime
                                            If TypeOf (DR(x)) Is DateTime Then
                                                item.SubItems.Add(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                            Else
                                                item.SubItems.Add(DR(x).ToString)
                                            End If
                                        End If

                                    Next
                                    lv.Items.Add(item)
                                Loop
                            End If
                            If autoSize = True Then .AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                        End With
                        DR.Close()
                        Return True
                    Catch ex As Exception
                        'Proses menampilkan pesan jika terjadi error
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream ke datagridview
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="dg">nama objek DataGridView</param>
                ''' <param name="addRows">Row DataGridView dapat ditambahkan oleh user. Default = False.</param>
                ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toDataGridView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",dataGridView1,False,True, ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",dataGridView1,False,True, , ,{param1,param2})
                ''' </example>
                Public Function ToDataGridView(ByVal sqlQuery As String, dg As DataGridView, Optional addRows As Boolean = False, Optional autoSize As Boolean = True,
                                Optional formatDateTime As String = Nothing,
                                Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                Optional showException As Boolean = True) As Boolean
                    Try

                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        Dim columnCount As Integer = DR.FieldCount

                        'Clear DataGridView
                        dg.DataSource = Nothing
                        dg.Columns.Clear()
                        dg.Rows.Clear()

                        'Buat header
                        For i As Integer = 0 To DR.FieldCount - 1
                            dg.Columns.Add(DR.GetName(i).ToString, DR.GetName(i).ToString)
                        Next

                        'menambah data row
                        If DR.HasRows Then
                            Dim rowData As String() = New String(columnCount - 1) {}
                            Do While DR.Read()
                                For k As Integer = 0 To DR.FieldCount - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        rowData(k) = DR(k).ToString
                                    Else
                                        'Formatted datetime
                                        If TypeOf (DR(k)) Is DateTime Then
                                            rowData(k) = DirectCast(DR(k), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                        Else
                                            rowData(k) = DR(k).ToString
                                        End If
                                    End If
                                Next
                                dg.Rows.Add(rowData)
                            Loop
                        End If
                        dg.AllowUserToAddRows = addRows
                        If autoSize = True Then
                            dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                            dg.AutoResizeColumns()
                        End If
                        DR.Close()
                        Return True
                    Catch ex As Exception
                        'Proses menampilkan pesan jika terjadi error
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream ke treeview
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="tv">Nama objek treeview</param>
                ''' <param name="keyValue">Menentukan value untuk key nodes di treeview</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toTreeView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",treeView1,"Data_Daerah", ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toTreeView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",treeView1,"Data_Daerah", , ,{param1,param2})
                ''' </example>
                Public Function ToTreeView(ByVal sqlQuery As String, tv As TreeView, keyValue As String, Optional formatDateTime As String = Nothing,
                                            Optional paramAddWithValue()() As Object = Nothing,
                                            Optional paramAdd()() As Object = Nothing, Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()

                        With tv
                            .Nodes.Clear()
                            Dim key As String
                            'menambah data row
                            If DR.HasRows Then
                                Do While DR.Read()
                                    key = keyValue + (.Nodes.Count - 1).ToString
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        .Nodes.Add(key, DR(0).ToString)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (DR(0)) Is DateTime Then
                                            .Nodes.Add(key, DirectCast(DR(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                        Else
                                            .Nodes.Add(key, DR(0).ToString)
                                        End If
                                    End If
                                    For x As Integer = 1 To DR.FieldCount - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, DR(x).ToString)
                                        Else
                                            'Formatted datetime
                                            If TypeOf (DR(x)) Is DateTime Then
                                                .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                            Else
                                                .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, DR(x).ToString)
                                            End If
                                        End If
                                    Next
                                Loop
                            End If
                            .ExpandAll()
                        End With
                        DR.Close()
                        Return True
                    Catch ex As Exception
                        'Proses menampilkan pesan jika terjadi error
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream ke combobox
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="cmb">Nama objek combobox</param>
                ''' <param name="displayMember">Kolom Query yang akan ditampilkan di ComboBox sebagai DisplayMember</param>
                ''' <param name="valueMember">Kolom Query yang akan di jadikan Nilai dari DisplayMember di Combobox</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toComboBox("select DaerahID,Kelurahan from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",comboBox1,"Kelurahan","DaerahID", ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toComboBox("select DaerahID,Kelurahan from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",comboBox1,"Kelurahan","DaerahID", , ,{param1,param2})
                ''' </example>
                Public Function ToComboBox(sqlQuery As String, cmb As ComboBox, displayMember As String, valueMember As String,
                                      Optional formatDateTime As String = Nothing,
                                      Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                      Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        Dim DT As New DataTable
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        DT.Load(DR)
                        DR.Close()

                        If DT.Rows.Count > 0 Then
                            cmb.DataSource = DT
                            cmb.DisplayMember = displayMember
                            cmb.ValueMember = valueMember
                            If formatDateTime <> Nothing Then
                                cmb.FormatString = formatDateTime
                            End If
                        End If
                        Return True
                    Catch ex As Exception
                        'Proses menampilkan pesan jika terjadi error
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk text
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="pathSaveFile">Full path lokasi export text</param>
                ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                ''' <param name="useHeader">Gunakan Header? Default = True</param>
                ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Boolean</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toText("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.txt",10,",",True,True, ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toText("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.txt",10,",",True,True, , ,{param1,param2})
                ''' </example>
                Public Function ToText(ByVal sqlQuery As String, ByVal pathSaveFile As String, Optional lastRow As Integer = 0, Optional delimiter As String = ",",
                                       Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True, Optional formatDateTime As String = Nothing,
                                       Optional paramAddWithValue()() As Object = Nothing,
                                       Optional paramAdd()() As Object = Nothing, Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        If DR.HasRows Then
                            Dim varSeparator As String = delimiter
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menambah header
                                    varText = Nothing
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        If encapsulation = True Then
                                            varText += """" + DR.GetName(i).Replace(Chr(34), "") + """" + varSeparator
                                        Else
                                            varText += DR.GetName(i).Replace(varSeparator, "") + varSeparator
                                        End If
                                    Next

                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If

                                'Menambah row
                                Dim processed As Integer = 0
                                Do While (DR.Read())
                                    varText = Nothing
                                    For x As Integer = 0 To DR.FieldCount - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            If encapsulation = True Then
                                                varText += IIf(IsDBNull(DR(x)) = True, """" + varSeparator, """" + DR(x).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(DR(x)) = True, varSeparator, DR(x).ToString.Replace(varSeparator, "") + varSeparator).ToString
                                            End If
                                        Else
                                            'Formatted datetime
                                            If encapsulation = True Then
                                                If TypeOf (DR(x)) Is DateTime Then
                                                    varText += IIf(IsDBNull(DR(x)) = True, """" + varSeparator, """" + DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(34), "") + """" + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(DR(x)) = True, """" + varSeparator, """" + DR(x).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                                End If
                                            Else
                                                If TypeOf (DR(x)) Is DateTime Then
                                                    varText += IIf(IsDBNull(DR(x)) = True, varSeparator, DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(varSeparator, "") + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(DR(x)) = True, varSeparator, DR(x).ToString.Replace(varSeparator, "") + varSeparator).ToString
                                                End If
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lastRow - 1 <> processed Then
                                        .WriteLine()
                                    End If
                                    processed += 1
                                Loop
                                .Close()
                                Return True
                            End With
                        Else
                            Return False
                        End If
                        DR.Close()
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk excel
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="pathSaveFile">Full path lokasi export excel</param>
                ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Boolean</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toExcel("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xlsx",Nothing,Nothing, ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toExcel("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xlsx",Nothing,Nothing, , ,{param1,param2})
                ''' </example>
                Public Function ToExcel(ByVal sqlQuery As String, ByVal pathSaveFile As String, Optional ByVal usePassword As String = Nothing,
                                        Optional ByVal passwordWorkBook As String = Nothing, Optional formatDateTime As String = Nothing,
                                        Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                        Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        If DR.HasRows Then
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To DR.FieldCount - 1
                                varExcelWorkSheet.Cells(1, i + 1) = DR.GetName(i)
                            Next
                            'Menambah row
                            Dim processed As Integer = 0
                            Do While (DR.Read())
                                For j As Integer = 0 To DR.FieldCount - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, "", DR(j))
                                    Else
                                        'Formatted datetime
                                        If TypeOf (DR(j)) Is DateTime Then
                                            varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, "", DirectCast(DR(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                        Else
                                            varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, "", DR(j))
                                        End If
                                    End If
                                Next
                                processed += 1
                            Loop


                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Else
                            Return False
                        End If
                        DR.Close()
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk csv
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="pathSaveFile">Full path lokasi export csv</param>
                ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                ''' <param name="useHeader">Gunakan Header? Default = True</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Boolean</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toCSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.csv",10,True, ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toCSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.csv",10,True, , ,{param1,param2})
                ''' </example>
                Public Function ToCSV(ByVal sqlQuery As String, ByVal pathSaveFile As String, Optional lastRow As Integer = 0, Optional useHeader As Boolean = True,
                                      Optional formatDateTime As String = Nothing,
                                      Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                      Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        If DR.HasRows Then
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menambah header
                                    varText = Nothing
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        varText += """" + DR.GetName(i).Replace(Chr(34), "") + """" + varSeparator
                                    Next

                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If

                                'Menambah row
                                Dim processed As Integer = 0
                                Do While (DR.Read())
                                    varText = Nothing
                                    For x As Integer = 0 To DR.FieldCount - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(DR(x)) = True, """" + varSeparator, """" + DR(x).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (DR(x)) Is DateTime Then
                                                varText += IIf(IsDBNull(DR(x)) = True, """" + varSeparator, """" + DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(34), "") + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(DR(x)) = True, """" + varSeparator, """" + DR(x).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lastRow - 1 <> processed Then
                                        .WriteLine()
                                    End If
                                    processed += 1
                                Loop
                                .Close()
                                Return True
                            End With
                        Else
                            Return False
                        End If
                        DR.Close()
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk tsv
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="pathSaveFile">Full path lokasi export tsv</param>
                ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                ''' <param name="useHeader">Gunakan Header? Default = True</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Boolean</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toTSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.tsv",10,True, ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toTSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.tsv",10,True, , ,{param1,param2})
                ''' </example>
                Public Function ToTSV(ByVal sqlQuery As String, ByVal pathSaveFile As String, Optional lastRow As Integer = 0, Optional useHeader As Boolean = True,
                                      Optional formatDateTime As String = Nothing,
                                      Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                      Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        If DR.HasRows Then
                            Dim varSeparator As String = Chr(9) 'Tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menambah header
                                    varText = Nothing
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        varText += DR.GetName(i).Replace(Chr(9), "") + varSeparator
                                    Next

                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If

                                'Menambah row
                                Dim processed As Integer = 0
                                Do While (DR.Read())
                                    varText = Nothing
                                    For x As Integer = 0 To DR.FieldCount - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(DR(x)) = True, varSeparator, DR(x).ToString.Replace(Chr(9), "") + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (DR(x)) Is DateTime Then
                                                varText += IIf(IsDBNull(DR(x)) = True, varSeparator, DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(9), "") + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(DR(x)) = True, varSeparator, DR(x).ToString.Replace(Chr(9), "") + varSeparator).ToString
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lastRow - 1 <> processed Then
                                        .WriteLine()
                                    End If
                                    processed += 1
                                Loop
                                .Close()
                                Return True
                            End With
                        Else
                            Return False
                        End If
                        DR.Close()
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk html
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="pathSaveFile">Full path lokasi export html</param>
                ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Boolean</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toHTML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.html","#CCCCCC", ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toHTML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.html","#CCCCCC", , ,{param1,param2})
                ''' </example>
                Public Function ToHTML(ByVal sqlQuery As String, ByVal pathSaveFile As String, Optional bgColor As String = "#CCCCCC",
                                       Optional formatDateTime As String = Nothing,
                                       Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                       Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        If DR.HasRows Then
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile

                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For i As Integer = 0 To DR.FieldCount - 1
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + DR.GetName(i) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()

                                'Menambah row
                                Do While (DR.Read())
                                    varText = "<TR>" + vbCrLf
                                    For x As Integer = 0 To DR.FieldCount - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, "", DR(x)).ToString + "</TD>" + vbCrLf
                                        Else
                                            'Formatted datetime
                                            If TypeOf (DR(x)) Is DateTime Then
                                                varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, "", DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)).ToString + "</TD>" + vbCrLf
                                            Else
                                                varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, "", DR(x)).ToString + "</TD>" + vbCrLf
                                            End If
                                        End If
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                Loop
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Else
                            Return False
                        End If
                        DR.Close()
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk xml
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="pathSaveFile">Full path lokasi export xml</param>
                ''' <param name="tableName">Nama table data</param>
                ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Boolean</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toXML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xml","data_daerah",True, ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toXML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xml","data_daerah",True, , ,{param1,param2})
                ''' </example>
                Public Function ToXML(ByVal sqlQuery As String, pathSaveFile As String, tableName As String, Optional writeSchema As Boolean = True,
                                      Optional formatDateTime As String = Nothing,
                                      Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                      Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()
                        If DR.HasRows = True Then
                            If formatDateTime = Nothing Then
                                'Standar datetime
                                Dim DT As New DataTable
                                DT.Load(DR)
                                DT.TableName = tableName
                                If writeSchema = True Then
                                    DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                                Else
                                    DT.WriteXml(pathSaveFile)
                                End If
                            Else
                                'Formatted datetime
                                Dim DT As New DataTable
                                DT.Load(DR)
                                DT.TableName = tableName

                                Dim DTFormat As New DataTable
                                For Each col As DataColumn In DT.Columns
                                    DTFormat.Columns.Add(col.ToString)
                                Next

                                For Each row As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTFormat.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If TypeOf (row(i)) Is DateTime Then
                                            drNew(i) = DirectCast(row(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                        Else
                                            drNew(i) = row(i).ToString
                                        End If

                                    Next
                                    DTFormat.Rows.Add(drNew)
                                Next
                                DTFormat.TableName = tableName
                                If writeSchema = True Then
                                    DTFormat.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                                Else
                                    DTFormat.WriteXml(pathSaveFile)
                                End If
                            End If
                            Return True
                        Else
                            Return False
                        End If
                        DR.Close()
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

                ''' <summary>
                ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk json
                ''' </summary>
                ''' <param name="sqlQuery">SQL Query</param>
                ''' <param name="pathSaveFile">Full path lokasi export json</param>
                ''' <param name="tableName">Nama table data</param>
                ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Boolean</returns>
                ''' 
                ''' <example>
                ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toJSON("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.js","data_daerah", ,{param1,param2})
                ''' 
                ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                '''
                ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                ''' DB.toJSON("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.js","data_daerah", , ,{param1,param2})
                ''' </example>
                Public Function ToJSON(ByVal sqlQuery As String, pathSaveFile As String, Optional tableName As String = Nothing,
                                       Optional formatDateTime As String = Nothing,
                                       Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                       Optional showException As Boolean = True) As Boolean
                    Try
                        Dim DR As MySqlDataReader
                        'Membuka koneksi database
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        'Proses menjalankan Query
                        Dim sqlCommand As New MySqlCommand

                        sqlCommand.Connection = sqlConn
                        sqlCommand.CommandType = CommandType.Text
                        sqlCommand.CommandText = sqlQuery
                        sqlCommand.CommandTimeout = TimeOut
                        If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                            For i As Integer = 0 To paramAddWithValue.Length - 1
                                CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                            Next
                        ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                            For i As Integer = 0 To paramAdd.Length - 1
                                CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                            Next
                        End If

                        DR = sqlCommand.ExecuteReader()

                        If DR.HasRows = True Then
                            If formatDateTime = Nothing Then
                                'Standar datetime
                                Dim DT As New DataTable
                                DT.Load(DR)
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                If tableName = Nothing Then
                                    varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented))
                                Else
                                    DT.TableName = tableName
                                    Dim ds As New DataSet
                                    ds.Tables.Add(DT)
                                    varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented))
                                End If
                                varTargetFile.Close()
                            Else
                                'Formatted datetime
                                Dim DT As New DataTable
                                DT.Load(DR)

                                'proses formatting
                                Dim DTFormat As New DataTable
                                For Each col As DataColumn In DT.Columns
                                    DTFormat.Columns.Add(col.ToString)
                                Next

                                For Each row As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTFormat.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If TypeOf (row(i)) Is DateTime Then
                                            drNew(i) = DirectCast(row(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                        Else
                                            drNew(i) = row(i).ToString
                                        End If
                                    Next
                                    DTFormat.Rows.Add(drNew)
                                Next

                                'proses json
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                If tableName = Nothing Then
                                    varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(DTFormat, Newtonsoft.Json.Formatting.Indented))
                                Else
                                    DTFormat.TableName = tableName
                                    Dim ds As New DataSet
                                    ds.Tables.Add(DTFormat)
                                    varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented))
                                End If
                                varTargetFile.Close()
                            End If
                            Return True
                        Else
                            Return False
                        End If
                        DR.Close()
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        'Menutup koneksi database
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function

#End Region

#Region "Status"
                ''' <summary>
                ''' Check status koneksi database. [AutoConnection tidak akan menampilkan status yang sedang terjadi]
                ''' </summary>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                ''' <returns>Boolean</returns>
                ''' <remarks>Untuk verifikasi status koneksi database</remarks>
                Public Function ConnectionStatus(Optional showException As Boolean = True) As Boolean
                    Try
                        If _autoConnection Then
                            If Connection Is Nothing Then
                                If ConnectionString <> Nothing Then
                                    sqlConn = New MySqlConnection(ConnectionString)
                                    sqlConn.Open()
                                Else
                                    sqlConn = New MySqlConnection(SharedConnectionString)
                                    sqlConn.Open()
                                End If
                            Else
                                sqlConn = Connection
                            End If
                        Else
                            sqlConn = _iconnection
                        End If

                        If sqlConn.State = ConnectionState.Open Then Return True Else Return False
                    Catch exs As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(exs, globals.Dev.Icons.Errors, NameClass)
                        Return False
                    Finally
                        If Connection Is Nothing AndAlso _autoConnection Then
                            If sqlConn IsNot Nothing Then
                                sqlConn.Close()
                                sqlConn.Dispose()
                            End If
                        End If
                        GC.Collect()
                    End Try
                End Function
#End Region

#Region "Manual Connection"
                ''' <summary>
                ''' Membuka koneksi database secara manual
                ''' </summary>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Sub OpenConnection(Optional showException As Boolean = True)
                    Try
                        If _autoConnection = False Then
                            If ConnectionString <> Nothing Then
                                _iconnection = New MySqlConnection(ConnectionString)
                            Else
                                _iconnection = New MySqlConnection(SharedConnectionString)
                            End If
                            _iconnection.Open()
                        End If
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                        _iconnection = Nothing
                    End Try
                End Sub

                ''' <summary>
                ''' Menutup koneksi database secara manual
                ''' </summary>
                ''' <param name="showException">Tampilkan log exception? Default = True</param>
                Public Sub CloseConnection(Optional showException As Boolean = True)
                    Try
                        If _autoConnection = False Then
                            If _iconnection IsNot Nothing Then
                                If _iconnection.State = ConnectionState.Open Then
                                    _iconnection.Close()
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                    End Try
                End Sub
#End Region

                ''' <summary>
                ''' Class untuk membuat query database ke output secara langsung dengan data terenkripsi
                ''' </summary>
                Public Class Encrypted

                    ''' <summary>
                    ''' Class enkripsi DES
                    ''' </summary>
                    Public Class DES
                        Private des As New crypt.DES
                        Private NameClass As String = "momolite.data.MySQL.BuildQuery.Direct.Encrypted.DES"

#Region "Property Connection String"
                        Private _connectionString As String = Nothing
                        ''' <summary>
                        ''' Connection String Database. Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' </summary>
                        ''' <value>Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"</value>
                        ''' <returns>String</returns>
                        Public Property ConnectionString() As String
                            Get
                                Return _connectionString
                            End Get
                            Set(ByVal value As String)
                                _connectionString = value
                            End Set
                        End Property
#End Region

#Region "Property TimeOut"
                        Private _timeOut As Integer = 600

                        ''' <summary>
                        ''' Connection TimeOut Data Adapter
                        ''' </summary>
                        ''' <value>Default value = 600</value>
                        ''' <returns>Integer</returns>
                        Public Property TimeOut() As Integer
                            Get
                                Return _timeOut
                            End Get
                            Set(ByVal value As Integer)
                                _timeOut = value
                            End Set
                        End Property
#End Region

#Region "Property Manual"
                        Private _autoConnection As Boolean = True
                        Private _iconnection As MySqlConnection = Nothing
                        Private _connection As MySqlConnection = Nothing

                        ''' <summary>
                        ''' Original connection database
                        ''' </summary>
                        ''' <value>MySqlConnection objek</value>
                        ''' <returns>MySqlConnection</returns>
                        Public Property Connection() As MySqlConnection
                            Get
                                Return _connection
                            End Get
                            Set(ByVal value As MySqlConnection)
                                _connection = value
                            End Set
                        End Property

                        ''' <summary>
                        ''' Gunakan Auto Connection database
                        ''' </summary>
                        ''' <value>Default value is True</value>
                        ''' <returns>Boolean</returns>
                        Public Property AutoConnection As Boolean
                            Get
                                Return _autoConnection
                            End Get
                            Set(ByVal value As Boolean)
                                If value <> _autoConnection Then
                                    _autoConnection = value
                                End If
                            End Set
                        End Property
#End Region

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya ke dalam memory dataset
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="nameTableDS">Nama tabel dataset</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Dataset</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah",{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah", ,{param1,param2})
                        ''' </example>
                        Public Function ToDataSet(ByVal sqlQuery As String, nameTableDS As String, ByVal secretKey As String,
                                                  Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                  Optional showException As Boolean = True) As DataSet
                            Try

                                Dim DR As MySqlDataReader
                                Dim DS As New DataSet
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                'Proses Enkripsi
                                Dim DT As New DataTable
                                Dim DTEncrypt As New DataTable
                                DT.Load(DR)

                                For Each col As DataColumn In DT.Columns
                                    DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                                Next

                                'menyimpan data
                                For Each drow As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTEncrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = des.Encode(drow(i).ToString, secretKey)
                                    Next
                                    DTEncrypt.Rows.Add(drNew)
                                Next

                                DTEncrypt.TableName = nameTableDS
                                DS.Tables.Add(DTEncrypt)
                                DR.Close()
                                Return DS

                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return Nothing

                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya ke dalam memory datatable
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Datatable</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan", ,{param1,param2})
                        ''' </example>
                        Public Function ToDataTable(ByVal sqlQuery As String, ByVal secretKey As String, Optional paramAddWithValue()() As Object = Nothing,
                                                    Optional paramAdd()() As Object = Nothing, Optional showException As Boolean = True) As DataTable
                            Try

                                Dim DR As MySqlDataReader
                                Dim DT As New DataTable
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                DT.Load(DR)

                                'Proses Enkripsi
                                Dim DTEncrypt As New DataTable

                                For Each col As DataColumn In DT.Columns
                                    DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                                Next

                                'menyimpan data
                                For Each drow As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTEncrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = des.Encode(drow(i).ToString, secretKey)
                                    Next
                                    DTEncrypt.Rows.Add(drNew)
                                Next
                                DR.Close()
                                Return DTEncrypt

                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return Nothing

                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

#Region "Output Object"
                        ''' <summary>
                        ''' Membuat Query secara direct stream ke listview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="lv">nama objek ListView</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",listView1,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",listView1,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToListView(ByVal sqlQuery As String, lv As ListView, ByVal secretKey As String, Optional autoSize As Boolean = True,
                                                   Optional formatDateTime As String = Nothing,
                                                   Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                   Optional showException As Boolean = True) As Boolean
                            Try

                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                With lv
                                    .Clear()
                                    .View = View.Details
                                    .FullRowSelect = True
                                    .GridLines = True
                                    'buat kolom otomatis
                                    .Columns.Clear()

                                    'Buat header
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        .Columns.Add(des.Encode(DR.GetName(i), secretKey))
                                    Next

                                    'menambah data row
                                    If DR.HasRows Then
                                        Do While DR.Read()
                                            Dim item As New ListViewItem
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                item.Text = des.Encode(DR(0).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(0)) Is DateTime Then
                                                    item.Text = des.Encode(DirectCast(DR(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    item.Text = des.Encode(DR(0).ToString, secretKey)
                                                End If
                                            End If

                                            For x As Integer = 1 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    item.SubItems.Add(des.Encode(DR(x).ToString, secretKey))
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        item.SubItems.Add(des.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                    Else
                                                        item.SubItems.Add(des.Encode(DR(x).ToString, secretKey))
                                                    End If
                                                End If
                                            Next
                                            lv.Items.Add(item)
                                        Loop
                                    End If
                                    If autoSize = True Then .AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                                End With
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke datagridview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="dg">nama objek DataGridView</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="addRows">Row DataGridView dapat ditambahkan oleh user. Default = False.</param>
                        ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataGridView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",dataGridView1,False,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",dataGridView1,False,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToDataGridView(ByVal sqlQuery As String, dg As DataGridView, ByVal secretKey As String, Optional addRows As Boolean = False,
                                                       Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing,
                                                       Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                       Optional showException As Boolean = True) As Boolean
                            Try

                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                Dim columnCount As Integer = DR.FieldCount

                                'Clear DataGridView
                                dg.DataSource = Nothing
                                dg.Columns.Clear()
                                dg.Rows.Clear()

                                'Buat header
                                For i As Integer = 0 To DR.FieldCount - 1
                                    dg.Columns.Add(des.Encode(DR.GetName(i).ToString, secretKey), des.Encode(DR.GetName(i).ToString, secretKey))
                                Next

                                'menambah data row
                                If DR.HasRows Then
                                    Dim rowData As String() = New String(columnCount - 1) {}
                                    Do While DR.Read()
                                        For k As Integer = 0 To DR.FieldCount - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                rowData(k) = des.Encode(DR(k).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(k)) Is DateTime Then
                                                    rowData(k) = des.Encode(DirectCast(DR(k), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    rowData(k) = des.Encode(DR(k).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        dg.Rows.Add(rowData)
                                    Loop
                                End If
                                dg.AllowUserToAddRows = addRows
                                If autoSize = True Then
                                    dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                                    dg.AutoResizeColumns()
                                End If
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke treeview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="tv">Nama objek treeview</param>
                        ''' <param name="keyValue">Menentukan value untuk key nodes di treeview</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTreeView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",treeView1,"Data_Daerah", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTreeView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",treeView1,"Data_Daerah", , ,{param1,param2})
                        ''' </example>
                        Public Function ToTreeView(ByVal sqlQuery As String, tv As TreeView, keyValue As String, ByVal secretKey As String,
                                                   Optional formatDateTime As String = Nothing,
                                                   Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                   Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                With tv
                                    .Nodes.Clear()
                                    Dim key As String
                                    'menambah data row
                                    If DR.HasRows Then
                                        Do While DR.Read()
                                            key = keyValue + (.Nodes.Count - 1).ToString
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                .Nodes.Add(key, des.Encode(DR(0).ToString, secretKey))
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(0)) Is DateTime Then
                                                    .Nodes.Add(key, des.Encode(DirectCast(DR(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                Else
                                                    .Nodes.Add(key, des.Encode(DR(0).ToString, secretKey))
                                                End If
                                            End If
                                            For x As Integer = 1 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, des.Encode(DR(x).ToString, secretKey))
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, des.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                    Else
                                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, des.Encode(DR(x).ToString, secretKey))
                                                    End If
                                                End If
                                            Next
                                        Loop
                                    End If
                                    .ExpandAll()
                                End With
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke combobox
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="cmb">Nama objek combobox</param>
                        ''' <param name="displayMember">Kolom Query yang akan ditampilkan di ComboBox sebagai DisplayMember</param>
                        ''' <param name="valueMember">Kolom Query yang akan di jadikan Nilai dari DisplayMember di Combobox</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toComboBox("select DaerahID,Kelurahan from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",comboBox1,"Kelurahan","DaerahID", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toComboBox("select DaerahID,Kelurahan from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",comboBox1,"Kelurahan","DaerahID", , ,{param1,param2})
                        ''' </example>
                        Public Function ToComboBox(sqlQuery As String, cmb As ComboBox, displayMember As String, valueMember As String,
                                      secretKey As String, Optional formatDateTime As String = Nothing,
                                      Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                      Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                Dim DT As New DataTable
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                DT.Load(DR)
                                DR.Close()

                                If DT.Rows.Count > 0 Then
                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = des.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = des.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = des.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    'load ke combobox
                                    cmb.DataSource = DTEncrypt
                                    cmb.DisplayMember = des.Encode(displayMember, secretKey)
                                    cmb.ValueMember = des.Encode(valueMember, secretKey)
                                End If
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk text
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export text</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toText("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.txt",10,",",True,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toText("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.txt",10,",",True,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToText(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                               Optional delimiter As String = ",", Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True,
                                               Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = delimiter
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                If encapsulation = True Then
                                                    varText += """" + des.Encode(DR.GetName(i), secretKey) + """" + varSeparator
                                                Else
                                                    varText += des.Encode(DR.GetName(i), secretKey) + varSeparator
                                                End If
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    If encapsulation = True Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, des.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                    End If
                                                Else
                                                    'Formatted datetime
                                                    If encapsulation = True Then
                                                        If TypeOf (DR(x)) Is DateTime Then
                                                            varText += IIf(IsDBNull(DR(x)) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                        Else
                                                            varText += IIf(IsDBNull(DR(x)) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                        End If
                                                    Else
                                                        If TypeOf (DR(x)) Is DateTime Then
                                                            varText += IIf(IsDBNull(DR(x)) = True, varSeparator, des.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                        Else
                                                            varText += IIf(IsDBNull(DR(x)) = True, varSeparator, des.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk excel
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export excel</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                        ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toExcel("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xlsx",Nothing,Nothing, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toExcel("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xlsx",Nothing,Nothing, , ,{param1,param2})
                        ''' </example>
                        Public Function ToExcel(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String,
                                                Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                                Optional formatDateTime As String = Nothing,
                                                Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varExcelApp As Excels.Application
                                    Dim varExcelWorkBook As Excels.Workbook
                                    Dim varExcelWorkSheet As Excels.Worksheet
                                    Dim misValue As Object = System.Reflection.Missing.Value

                                    varExcelApp = New Excels.Application
                                    varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                                    varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                                    'Menambah header
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        varExcelWorkSheet.Cells(1, i + 1) = des.Encode(DR.GetName(i), secretKey)
                                    Next
                                    'Menambah row
                                    Dim processed As Integer = 0
                                    Do While (DR.Read())
                                        For j As Integer = 0 To DR.FieldCount - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, des.Encode("", secretKey), des.Encode(DR(j).ToString, secretKey))
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(j)) Is DateTime Then
                                                    varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, des.Encode("", secretKey), des.Encode(DirectCast(DR(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                Else
                                                    varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, des.Encode("", secretKey), des.Encode(DR(j).ToString, secretKey))
                                                End If
                                            End If
                                        Next
                                        processed += 1
                                    Loop


                                    If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                                    If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                                    varExcelWorkSheet.SaveAs(pathSaveFile)
                                    varExcelWorkBook.Close()
                                    varExcelApp.Quit()

                                    releaseObject(varExcelApp)
                                    releaseObject(varExcelWorkBook)
                                    releaseObject(varExcelWorkSheet)
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk csv
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export csv</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toCSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.csv",10,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toCSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.csv",10,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToCSV(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                              Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = "," 'comma
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                varText += """" + des.Encode(DR.GetName(i), secretKey) + """" + varSeparator
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += IIf(IsDBNull(DR(x)) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk tsv
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export tsv</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.tsv",10,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.tsv",10,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToTSV(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                              Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = Chr(9) 'Tab
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                varText += des.Encode(DR.GetName(i), secretKey) + varSeparator
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += IIf(IsDBNull(DR(x)) = True, varSeparator, des.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, des.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, des.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk html
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export html</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toHTML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.html","#CCCCCC", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toHTML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.html","#CCCCCC", , ,{param1,param2})
                        ''' </example>
                        Public Function ToHTML(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String,
                                               Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile

                                        'Menyimpan header
                                        varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                        For i As Integer = 0 To DR.FieldCount - 1
                                            varText += "<TH bgcolor='" + bgColor + "' align='center'>" + des.Encode(DR.GetName(i), secretKey) + "</TH>" + vbCrLf
                                        Next
                                        varText += "</TR>" + vbCrLf
                                        .Write(varText)
                                        .WriteLine()

                                        'Menambah row
                                        Do While (DR.Read())
                                            varText = "<TR>" + vbCrLf
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, des.Encode("", secretKey), des.Encode(DR(x).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, des.Encode("", secretKey), des.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                                    Else
                                                        varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, des.Encode("", secretKey), des.Encode(DR(x).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                                    End If
                                                End If
                                            Next
                                            varText += "</TR>" + vbCrLf
                                            .Write(varText)
                                            .WriteLine()
                                        Loop
                                        varText = "</TABLE>"
                                        .Write(varText)
                                        .Close()
                                    End With
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk xml
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export xml</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="tableName">Nama table data</param>
                        ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toXML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xml","data_daerah",True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toXML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xml","data_daerah",True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToXML(ByVal sqlQuery As String, pathSaveFile As String, secretKey As String, tableName As String,
                                              Optional writeSchema As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows = True Then
                                    Dim DT As New DataTable
                                    DT.Load(DR)

                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = des.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = des.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = des.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    DTEncrypt.TableName = tableName
                                    If writeSchema = True Then
                                        DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                                    Else
                                        DTEncrypt.WriteXml(pathSaveFile)
                                    End If
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk json
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export json</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="tableName">Nama table data</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toJSON("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.js","data_daerah", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toJSON("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.js","data_daerah", , ,{param1,param2})
                        ''' </example>
                        Public Function ToJSON(ByVal sqlQuery As String, pathSaveFile As String, ByVal secretKey As String, Optional tableName As String = Nothing,
                                               Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                If DR.HasRows = True Then
                                    Dim DT As New DataTable
                                    DT.Load(DR)

                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = des.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = des.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = des.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    If tableName = Nothing Then
                                        varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented))
                                    Else
                                        DTEncrypt.TableName = tableName
                                        Dim ds As New DataSet
                                        ds.Tables.Add(DTEncrypt)
                                        varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented))
                                    End If
                                    varTargetFile.Close()
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

#End Region

#Region "Status"
                        ''' <summary>
                        ''' Check status koneksi database. [AutoConnection tidak akan menampilkan status yang sedang terjadi]
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' <remarks>Untuk verifikasi status koneksi database</remarks>
                        Public Function ConnectionStatus(Optional showException As Boolean = True) As Boolean
                            Try
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                If sqlConn.State = ConnectionState.Open Then Return True Else Return False
                            Catch exs As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(exs, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function
#End Region

#Region "Manual Connection"
                        ''' <summary>
                        ''' Membuka koneksi database secara manual
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        Public Sub OpenConnection(Optional showException As Boolean = True)
                            Try
                                If _autoConnection = False Then
                                    If ConnectionString <> Nothing Then
                                        _iconnection = New MySqlConnection(ConnectionString)
                                    Else
                                        _iconnection = New MySqlConnection(SharedConnectionString)
                                    End If
                                    _iconnection.Open()
                                End If
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                _iconnection = Nothing
                            End Try
                        End Sub

                        ''' <summary>
                        ''' Menutup koneksi database secara manual
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        Public Sub CloseConnection(Optional showException As Boolean = True)
                            Try
                                If _autoConnection = False Then
                                    If _iconnection IsNot Nothing Then
                                        If _iconnection.State = ConnectionState.Open Then
                                            _iconnection.Close()
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                            End Try
                        End Sub
#End Region
                    End Class

                    ''' <summary>
                    ''' Class enkripsi 3DES
                    ''' </summary>
                    Public Class TripleDES
                        Private tripledes As New crypt.TripleDES
                        Private NameClass As String = "momolite.data.MySQL.BuildQuery.Direct.Encrypted.TripleDES"

#Region "Property Connection String"
                        Private _connectionString As String = Nothing
                        ''' <summary>
                        ''' Connection String Database. Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' </summary>
                        ''' <value>Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"</value>
                        ''' <returns>String</returns>
                        Public Property ConnectionString() As String
                            Get
                                Return _connectionString
                            End Get
                            Set(ByVal value As String)
                                _connectionString = value
                            End Set
                        End Property
#End Region

#Region "Property TimeOut"
                        Private _timeOut As Integer = 600

                        ''' <summary>
                        ''' Connection TimeOut Data Adapter
                        ''' </summary>
                        ''' <value>Default value = 600</value>
                        ''' <returns>Integer</returns>
                        Public Property TimeOut() As Integer
                            Get
                                Return _timeOut
                            End Get
                            Set(ByVal value As Integer)
                                _timeOut = value
                            End Set
                        End Property
#End Region

#Region "Property Manual"
                        Private _autoConnection As Boolean = True
                        Private _iconnection As MySqlConnection = Nothing
                        Private _connection As MySqlConnection = Nothing

                        ''' <summary>
                        ''' Original connection database
                        ''' </summary>
                        ''' <value>MySqlConnection objek</value>
                        ''' <returns>MySqlConnection</returns>
                        Public Property Connection() As MySqlConnection
                            Get
                                Return _connection
                            End Get
                            Set(ByVal value As MySqlConnection)
                                _connection = value
                            End Set
                        End Property

                        ''' <summary>
                        ''' Gunakan Auto Connection database
                        ''' </summary>
                        ''' <value>Default value is True</value>
                        ''' <returns>Boolean</returns>
                        Public Property AutoConnection As Boolean
                            Get
                                Return _autoConnection
                            End Get
                            Set(ByVal value As Boolean)
                                If value <> _autoConnection Then
                                    _autoConnection = value
                                End If
                            End Set
                        End Property
#End Region

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya ke dalam memory dataset
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="nameTableDS">Nama tabel dataset</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Dataset</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah",{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah", ,{param1,param2})
                        ''' </example>
                        Public Function ToDataSet(ByVal sqlQuery As String, nameTableDS As String, ByVal secretKey As String,
                                                  Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                  Optional showException As Boolean = True) As DataSet
                            Try

                                Dim DR As MySqlDataReader
                                Dim DS As New DataSet
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                'Proses Enkripsi
                                Dim DT As New DataTable
                                Dim DTEncrypt As New DataTable
                                DT.Load(DR)

                                For Each col As DataColumn In DT.Columns
                                    DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                                Next

                                For Each drow As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTEncrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = tripledes.Encode(drow(i).ToString, secretKey)
                                    Next
                                    DTEncrypt.Rows.Add(drNew)
                                Next

                                DTEncrypt.TableName = nameTableDS
                                DS.Tables.Add(DTEncrypt)
                                DR.Close()
                                Return DS

                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return Nothing

                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya ke dalam memory datatable
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Datatable</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan", ,{param1,param2})
                        ''' </example>
                        Public Function ToDataTable(ByVal sqlQuery As String, ByVal secretKey As String, Optional paramAddWithValue()() As Object = Nothing,
                                                    Optional paramAdd()() As Object = Nothing, Optional showException As Boolean = True) As DataTable
                            Try

                                Dim DR As MySqlDataReader
                                Dim DT As New DataTable
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                DT.Load(DR)

                                'Proses Enkripsi
                                Dim DTEncrypt As New DataTable

                                For Each col As DataColumn In DT.Columns
                                    DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                                Next

                                For Each drow As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTEncrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = tripledes.Encode(drow(i).ToString, secretKey)
                                    Next
                                    DTEncrypt.Rows.Add(drNew)
                                Next
                                DR.Close()
                                Return DTEncrypt

                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return Nothing

                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

#Region "Output Object"
                        ''' <summary>
                        ''' Membuat Query secara direct stream ke listview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="lv">nama objek ListView</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",listView1,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",listView1,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToListView(ByVal sqlQuery As String, lv As ListView, ByVal secretKey As String, Optional autoSize As Boolean = True,
                                                   Optional formatDateTime As String = Nothing,
                                                   Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                   Optional showException As Boolean = True) As Boolean
                            Try

                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                With lv
                                    .Clear()
                                    .View = View.Details
                                    .FullRowSelect = True
                                    .GridLines = True
                                    'buat kolom otomatis
                                    .Columns.Clear()

                                    'Buat header
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        .Columns.Add(tripledes.Encode(DR.GetName(i), secretKey))
                                    Next

                                    'menambah data row
                                    If DR.HasRows Then
                                        Do While DR.Read()
                                            Dim item As New ListViewItem
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                item.Text = tripledes.Encode(DR(0).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(0)) Is DateTime Then
                                                    item.Text = tripledes.Encode(DirectCast(DR(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    item.Text = tripledes.Encode(DR(0).ToString, secretKey)
                                                End If
                                            End If

                                            For x As Integer = 1 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    item.SubItems.Add(tripledes.Encode(DR(x).ToString, secretKey))
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        item.SubItems.Add(tripledes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                    Else
                                                        item.SubItems.Add(tripledes.Encode(DR(x).ToString, secretKey))
                                                    End If
                                                End If
                                            Next
                                            lv.Items.Add(item)
                                        Loop
                                    End If
                                    If autoSize = True Then .AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                                End With
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke datagridview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="dg">nama objek DataGridView</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="addRows">Row DataGridView dapat ditambahkan oleh user. Default = False.</param>
                        ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataGridView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",dataGridView1,False,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",dataGridView1,False,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToDataGridView(ByVal sqlQuery As String, dg As DataGridView, ByVal secretKey As String, Optional addRows As Boolean = False,
                                                       Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing,
                                                       Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                       Optional showException As Boolean = True) As Boolean
                            Try

                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                Dim columnCount As Integer = DR.FieldCount

                                'Clear DataGridView
                                dg.DataSource = Nothing
                                dg.Columns.Clear()
                                dg.Rows.Clear()

                                'Buat header
                                For i As Integer = 0 To DR.FieldCount - 1
                                    dg.Columns.Add(tripledes.Encode(DR.GetName(i).ToString, secretKey), tripledes.Encode(DR.GetName(i).ToString, secretKey))
                                Next

                                'menambah data row
                                If DR.HasRows Then
                                    Dim rowData As String() = New String(columnCount - 1) {}
                                    Do While DR.Read()
                                        For k As Integer = 0 To DR.FieldCount - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                rowData(k) = tripledes.Encode(DR(k).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(k)) Is DateTime Then
                                                    rowData(k) = tripledes.Encode(DirectCast(DR(k), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    rowData(k) = tripledes.Encode(DR(k).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        dg.Rows.Add(rowData)
                                    Loop
                                End If
                                dg.AllowUserToAddRows = addRows
                                If autoSize = True Then
                                    dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                                    dg.AutoResizeColumns()
                                End If
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke treeview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="tv">Nama objek treeview</param>
                        ''' <param name="keyValue">Menentukan value untuk key nodes di treeview</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTreeView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",treeView1,"Data_Daerah", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTreeView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",treeView1,"Data_Daerah", , ,{param1,param2})
                        ''' </example>
                        Public Function ToTreeView(ByVal sqlQuery As String, tv As TreeView, keyValue As String, ByVal secretKey As String,
                                                   Optional formatDateTime As String = Nothing,
                                                   Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                   Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                With tv
                                    .Nodes.Clear()
                                    Dim key As String
                                    'menambah data row
                                    If DR.HasRows Then
                                        Do While DR.Read()
                                            key = keyValue + (.Nodes.Count - 1).ToString
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                .Nodes.Add(key, tripledes.Encode(DR(0).ToString, secretKey))
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(0)) Is DateTime Then
                                                    .Nodes.Add(key, tripledes.Encode(DirectCast(DR(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                Else
                                                    .Nodes.Add(key, tripledes.Encode(DR(0).ToString, secretKey))
                                                End If
                                            End If
                                            For x As Integer = 1 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, tripledes.Encode(DR(x).ToString, secretKey))
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, tripledes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                    Else
                                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, tripledes.Encode(DR(x).ToString, secretKey))
                                                    End If
                                                End If
                                            Next
                                        Loop
                                    End If
                                    .ExpandAll()
                                End With
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke combobox
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="cmb">Nama objek combobox</param>
                        ''' <param name="displayMember">Kolom Query yang akan ditampilkan di ComboBox sebagai DisplayMember</param>
                        ''' <param name="valueMember">Kolom Query yang akan di jadikan Nilai dari DisplayMember di Combobox</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toComboBox("select DaerahID,Kelurahan from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",comboBox1,"Kelurahan","DaerahID", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toComboBox("select DaerahID,Kelurahan from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",comboBox1,"Kelurahan","DaerahID", , ,{param1,param2})
                        ''' </example>
                        Public Function ToComboBox(sqlQuery As String, cmb As ComboBox, displayMember As String, valueMember As String,
                                      secretKey As String, Optional formatDateTime As String = Nothing,
                                      Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                      Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                Dim DT As New DataTable
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                DT.Load(DR)
                                DR.Close()

                                If DT.Rows.Count > 0 Then
                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = tripledes.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = tripledes.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = tripledes.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    'load ke combobox
                                    cmb.DataSource = DTEncrypt
                                    cmb.DisplayMember = tripledes.Encode(displayMember, secretKey)
                                    cmb.ValueMember = tripledes.Encode(valueMember, secretKey)
                                End If
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk text
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export text</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toText("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.txt",10,",",True,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toText("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.txt",10,",",True,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToText(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                               Optional delimiter As String = ",", Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True,
                                               Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = delimiter
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                If encapsulation = True Then
                                                    varText += """" + tripledes.Encode(DR.GetName(i), secretKey) + """" + varSeparator
                                                Else
                                                    varText += tripledes.Encode(DR.GetName(i), secretKey) + varSeparator
                                                End If
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    If encapsulation = True Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, tripledes.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                    End If
                                                Else
                                                    'Formatted datetime
                                                    If encapsulation = True Then
                                                        If TypeOf (DR(x)) Is DateTime Then
                                                            varText += IIf(IsDBNull(DR(x)) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                        Else
                                                            varText += IIf(IsDBNull(DR(x)) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                        End If
                                                    Else
                                                        If TypeOf (DR(x)) Is DateTime Then
                                                            varText += IIf(IsDBNull(DR(x)) = True, varSeparator, tripledes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                        Else
                                                            varText += IIf(IsDBNull(DR(x)) = True, varSeparator, tripledes.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk excel
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export excel</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                        ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toExcel("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xlsx",Nothing,Nothing, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toExcel("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xlsx",Nothing,Nothing, , ,{param1,param2})
                        ''' </example>
                        Public Function ToExcel(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String,
                                                Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                                Optional formatDateTime As String = Nothing,
                                                Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varExcelApp As Excels.Application
                                    Dim varExcelWorkBook As Excels.Workbook
                                    Dim varExcelWorkSheet As Excels.Worksheet
                                    Dim misValue As Object = System.Reflection.Missing.Value

                                    varExcelApp = New Excels.Application
                                    varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                                    varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                                    'Menambah header
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        varExcelWorkSheet.Cells(1, i + 1) = tripledes.Encode(DR.GetName(i), secretKey)
                                    Next
                                    'Menambah row
                                    Dim processed As Integer = 0
                                    Do While (DR.Read())
                                        For j As Integer = 0 To DR.FieldCount - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DR(j).ToString, secretKey))
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(j)) Is DateTime Then
                                                    varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DirectCast(DR(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                Else
                                                    varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DR(j).ToString, secretKey))
                                                End If
                                            End If
                                        Next
                                        processed += 1
                                    Loop

                                    If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                                    If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                                    varExcelWorkSheet.SaveAs(pathSaveFile)
                                    varExcelWorkBook.Close()
                                    varExcelApp.Quit()

                                    releaseObject(varExcelApp)
                                    releaseObject(varExcelWorkBook)
                                    releaseObject(varExcelWorkSheet)
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk csv
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export csv</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toCSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.csv",10,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toCSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.csv",10,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToCSV(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                              Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = "," 'comma
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                varText += """" + tripledes.Encode(DR.GetName(i), secretKey) + """" + varSeparator
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += IIf(IsDBNull(DR(x)) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk tsv
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export tsv</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.tsv",10,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.tsv",10,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToTSV(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                              Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = Chr(9) 'Tab
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                varText += tripledes.Encode(DR.GetName(i), secretKey) + varSeparator
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += IIf(IsDBNull(DR(x)) = True, varSeparator, tripledes.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, tripledes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, tripledes.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk html
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export html</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toHTML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.html","#CCCCCC", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toHTML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.html","#CCCCCC", , ,{param1,param2})
                        ''' </example>
                        Public Function ToHTML(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String,
                                               Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile

                                        'Menyimpan header
                                        varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                        For i As Integer = 0 To DR.FieldCount - 1
                                            varText += "<TH bgcolor='" + bgColor + "' align='center'>" + tripledes.Encode(DR.GetName(i), secretKey) + "</TH>" + vbCrLf
                                        Next
                                        varText += "</TR>" + vbCrLf
                                        .Write(varText)
                                        .WriteLine()

                                        'Menambah row
                                        Do While (DR.Read())
                                            varText = "<TR>" + vbCrLf
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DR(x).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                                    Else
                                                        varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DR(x).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                                    End If
                                                End If
                                            Next
                                            varText += "</TR>" + vbCrLf
                                            .Write(varText)
                                            .WriteLine()
                                        Loop
                                        varText = "</TABLE>"
                                        .Write(varText)
                                        .Close()
                                    End With
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk xml
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export xml</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="tableName">Nama table data</param>
                        ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toXML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xml","data_daerah",True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toXML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xml","data_daerah",True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToXML(ByVal sqlQuery As String, pathSaveFile As String, secretKey As String, tableName As String,
                                              Optional writeSchema As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows = True Then
                                    Dim DT As New DataTable
                                    DT.Load(DR)

                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = tripledes.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = tripledes.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = tripledes.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    DTEncrypt.TableName = tableName
                                    If writeSchema = True Then
                                        DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                                    Else
                                        DTEncrypt.WriteXml(pathSaveFile)
                                    End If
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk json
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export json</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="tableName">Nama table data</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toJSON("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.js","data_daerah", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toJSON("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.js","data_daerah", , ,{param1,param2})
                        ''' </example>
                        Public Function ToJSON(ByVal sqlQuery As String, pathSaveFile As String, ByVal secretKey As String, Optional tableName As String = Nothing,
                                               Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                If DR.HasRows = True Then
                                    Dim DT As New DataTable
                                    DT.Load(DR)

                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = tripledes.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = tripledes.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = tripledes.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    If tableName = Nothing Then
                                        varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented))
                                    Else
                                        DTEncrypt.TableName = tableName
                                        Dim ds As New DataSet
                                        ds.Tables.Add(DTEncrypt)
                                        varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented))
                                    End If
                                    varTargetFile.Close()
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

#End Region

#Region "Status"
                        ''' <summary>
                        ''' Check status koneksi database. [AutoConnection tidak akan menampilkan status yang sedang terjadi]
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' <remarks>Untuk verifikasi status koneksi database</remarks>
                        Public Function ConnectionStatus(Optional showException As Boolean = True) As Boolean
                            Try
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                If sqlConn.State = ConnectionState.Open Then Return True Else Return False
                            Catch exs As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(exs, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function
#End Region

#Region "Manual Connection"
                        ''' <summary>
                        ''' Membuka koneksi database secara manual
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        Public Sub OpenConnection(Optional showException As Boolean = True)
                            Try
                                If _autoConnection = False Then
                                    If ConnectionString <> Nothing Then
                                        _iconnection = New MySqlConnection(ConnectionString)
                                    Else
                                        _iconnection = New MySqlConnection(SharedConnectionString)
                                    End If
                                    _iconnection.Open()
                                End If
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                _iconnection = Nothing
                            End Try
                        End Sub

                        ''' <summary>
                        ''' Menutup koneksi database secara manual
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        Public Sub CloseConnection(Optional showException As Boolean = True)
                            Try
                                If _autoConnection = False Then
                                    If _iconnection IsNot Nothing Then
                                        If _iconnection.State = ConnectionState.Open Then
                                            _iconnection.Close()
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                            End Try
                        End Sub
#End Region

                    End Class

                    ''' <summary>
                    ''' Class enkripsi AES
                    ''' </summary>
                    Public Class AES
                        Private aes As New crypt.AES
                        Private NameClass As String = "momolite.data.MySQL.BuildQuery.Direct.Encrypted.AES"

#Region "Property Connection String"
                        Private _connectionString As String = Nothing
                        ''' <summary>
                        ''' Connection String Database. Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' </summary>
                        ''' <value>Contoh: "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"</value>
                        ''' <returns>String</returns>
                        Public Property ConnectionString() As String
                            Get
                                Return _connectionString
                            End Get
                            Set(ByVal value As String)
                                _connectionString = value
                            End Set
                        End Property
#End Region

#Region "Property TimeOut"
                        Private _timeOut As Integer = 600

                        ''' <summary>
                        ''' Connection TimeOut Data Adapter
                        ''' </summary>
                        ''' <value>Default value = 600</value>
                        ''' <returns>Integer</returns>
                        Public Property TimeOut() As Integer
                            Get
                                Return _timeOut
                            End Get
                            Set(ByVal value As Integer)
                                _timeOut = value
                            End Set
                        End Property
#End Region

#Region "Property Manual"
                        Private _autoConnection As Boolean = True
                        Private _iconnection As MySqlConnection = Nothing
                        Private _connection As MySqlConnection = Nothing

                        ''' <summary>
                        ''' Original connection database
                        ''' </summary>
                        ''' <value>MySqlConnection objek</value>
                        ''' <returns>MySqlConnection</returns>
                        Public Property Connection() As MySqlConnection
                            Get
                                Return _connection
                            End Get
                            Set(ByVal value As MySqlConnection)
                                _connection = value
                            End Set
                        End Property

                        ''' <summary>
                        ''' Gunakan Auto Connection database
                        ''' </summary>
                        ''' <value>Default value is True</value>
                        ''' <returns>Boolean</returns>
                        Public Property AutoConnection As Boolean
                            Get
                                Return _autoConnection
                            End Get
                            Set(ByVal value As Boolean)
                                If value <> _autoConnection Then
                                    _autoConnection = value
                                End If
                            End Set
                        End Property
#End Region

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya ke dalam memory dataset
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="nameTableDS">Nama tabel dataset</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Dataset</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah",{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataSet("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","data_daerah", ,{param1,param2})
                        ''' </example>
                        Public Function ToDataSet(ByVal sqlQuery As String, nameTableDS As String, ByVal secretKey As String,
                                                  Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                  Optional showException As Boolean = True) As DataSet
                            Try

                                Dim DR As MySqlDataReader
                                Dim DS As New DataSet
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                'Proses Enkripsi
                                Dim DT As New DataTable
                                Dim DTEncrypt As New DataTable
                                DT.Load(DR)

                                For Each col As DataColumn In DT.Columns
                                    DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                                Next

                                For Each drow As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTEncrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = aes.Encode(drow(i).ToString, secretKey)
                                    Next
                                    DTEncrypt.Rows.Add(drNew)
                                Next

                                DTEncrypt.TableName = nameTableDS
                                DS.Tables.Add(DTEncrypt)
                                DR.Close()
                                Return DS

                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return Nothing

                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya ke dalam memory datatable
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Datatable</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataTable("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan", ,{param1,param2})
                        ''' </example>
                        Public Function ToDataTable(ByVal sqlQuery As String, ByVal secretKey As String, Optional paramAddWithValue()() As Object = Nothing,
                                                    Optional paramAdd()() As Object = Nothing, Optional showException As Boolean = True) As DataTable
                            Try

                                Dim DR As MySqlDataReader
                                Dim DT As New DataTable
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                DT.Load(DR)

                                'Proses Enkripsi
                                Dim DTEncrypt As New DataTable

                                For Each col As DataColumn In DT.Columns
                                    DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                                Next

                                For Each drow As DataRow In DT.Rows
                                    Dim drNew As DataRow = DTEncrypt.NewRow()
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        drNew(i) = aes.Encode(drow(i).ToString, secretKey)
                                    Next
                                    DTEncrypt.Rows.Add(drNew)
                                Next
                                DR.Close()
                                Return DTEncrypt

                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return Nothing

                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

#Region "Output Object"
                        ''' <summary>
                        ''' Membuat Query secara direct stream ke listview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="lv">nama objek ListView</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",listView1,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",listView1,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToListView(ByVal sqlQuery As String, lv As ListView, ByVal secretKey As String, Optional autoSize As Boolean = True,
                                                   Optional formatDateTime As String = Nothing,
                                                   Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                   Optional showException As Boolean = True) As Boolean
                            Try

                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                With lv
                                    .Clear()
                                    .View = View.Details
                                    .FullRowSelect = True
                                    .GridLines = True
                                    'buat kolom otomatis
                                    .Columns.Clear()

                                    'Buat header
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        .Columns.Add(aes.Encode(DR.GetName(i), secretKey))
                                    Next

                                    'menambah data row
                                    If DR.HasRows Then
                                        Do While DR.Read()
                                            Dim item As New ListViewItem
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                item.Text = aes.Encode(DR(0).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(0)) Is DateTime Then
                                                    item.Text = aes.Encode(DirectCast(DR(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    item.Text = aes.Encode(DR(0).ToString, secretKey)
                                                End If
                                            End If

                                            For x As Integer = 1 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    item.SubItems.Add(aes.Encode(DR(x).ToString, secretKey))
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        item.SubItems.Add(aes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                    Else
                                                        item.SubItems.Add(aes.Encode(DR(x).ToString, secretKey))
                                                    End If
                                                End If
                                            Next
                                            lv.Items.Add(item)
                                        Loop
                                    End If
                                    If autoSize = True Then .AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                                End With
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke datagridview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="dg">nama objek DataGridView</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="addRows">Row DataGridView dapat ditambahkan oleh user. Default = False.</param>
                        ''' <param name="autoSize">Mengukur size kolom secara otomatis. Default = True.</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toDataGridView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",dataGridView1,False,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toListView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",dataGridView1,False,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToDataGridView(ByVal sqlQuery As String, dg As DataGridView, ByVal secretKey As String, Optional addRows As Boolean = False,
                                                       Optional autoSize As Boolean = True, Optional formatDateTime As String = Nothing,
                                                       Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                       Optional showException As Boolean = True) As Boolean
                            Try

                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                Dim columnCount As Integer = DR.FieldCount

                                'Clear DataGridView
                                dg.DataSource = Nothing
                                dg.Columns.Clear()
                                dg.Rows.Clear()

                                'Buat header
                                For i As Integer = 0 To DR.FieldCount - 1
                                    dg.Columns.Add(aes.Encode(DR.GetName(i).ToString, secretKey), aes.Encode(DR.GetName(i).ToString, secretKey))
                                Next

                                'menambah data row
                                If DR.HasRows Then
                                    Dim rowData As String() = New String(columnCount - 1) {}
                                    Do While DR.Read()
                                        For k As Integer = 0 To DR.FieldCount - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                rowData(k) = aes.Encode(DR(k).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(k)) Is DateTime Then
                                                    rowData(k) = aes.Encode(DirectCast(DR(k), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    rowData(k) = aes.Encode(DR(k).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        dg.Rows.Add(rowData)
                                    Loop
                                End If
                                dg.AllowUserToAddRows = addRows
                                If autoSize = True Then
                                    dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                                    dg.AutoResizeColumns()
                                End If
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke treeview
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="tv">Nama objek treeview</param>
                        ''' <param name="keyValue">Menentukan value untuk key nodes di treeview</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTreeView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",treeView1,"Data_Daerah", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTreeView("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",treeView1,"Data_Daerah", , ,{param1,param2})
                        ''' </example>
                        Public Function ToTreeView(ByVal sqlQuery As String, tv As TreeView, keyValue As String, ByVal secretKey As String,
                                                   Optional formatDateTime As String = Nothing,
                                                   Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                   Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                With tv
                                    .Nodes.Clear()
                                    Dim key As String
                                    'menambah data row
                                    If DR.HasRows Then
                                        Do While DR.Read()
                                            key = keyValue + (.Nodes.Count - 1).ToString
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                .Nodes.Add(key, aes.Encode(DR(0).ToString, secretKey))
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(0)) Is DateTime Then
                                                    .Nodes.Add(key, aes.Encode(DirectCast(DR(0), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                Else
                                                    .Nodes.Add(key, aes.Encode(DR(0).ToString, secretKey))
                                                End If
                                            End If
                                            For x As Integer = 1 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, aes.Encode(DR(x).ToString, secretKey))
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, aes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                    Else
                                                        .Nodes(.Nodes.Count - 1).Nodes.Add(key + "." + (.Nodes(.Nodes.Count - 1).Nodes.Count - 1).ToString, aes.Encode(DR(x).ToString, secretKey))
                                                    End If
                                                End If
                                            Next
                                        Loop
                                    End If
                                    .ExpandAll()
                                End With
                                DR.Close()
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream ke combobox
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="cmb">Nama objek combobox</param>
                        ''' <param name="displayMember">Kolom Query yang akan ditampilkan di ComboBox sebagai DisplayMember</param>
                        ''' <param name="valueMember">Kolom Query yang akan di jadikan Nilai dari DisplayMember di Combobox</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toComboBox("select DaerahID,Kelurahan from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",comboBox1,"Kelurahan","DaerahID", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toComboBox("select DaerahID,Kelurahan from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan",comboBox1,"Kelurahan","DaerahID", , ,{param1,param2})
                        ''' </example>
                        Public Function ToComboBox(sqlQuery As String, cmb As ComboBox, displayMember As String, valueMember As String,
                                      secretKey As String, Optional formatDateTime As String = Nothing,
                                      Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                      Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                Dim DT As New DataTable
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                DT.Load(DR)
                                DR.Close()

                                If DT.Rows.Count > 0 Then
                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = aes.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = aes.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = aes.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    'load ke combobox
                                    cmb.DataSource = DTEncrypt
                                    cmb.DisplayMember = aes.Encode(displayMember, secretKey)
                                    cmb.ValueMember = aes.Encode(valueMember, secretKey)
                                End If
                                Return True
                            Catch ex As Exception
                                'Proses menampilkan pesan jika terjadi error
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk text
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export text</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toText("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.txt",10,",",True,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toText("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.txt",10,",",True,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToText(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                               Optional delimiter As String = ",", Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True,
                                               Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = delimiter
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                If encapsulation = True Then
                                                    varText += """" + aes.Encode(DR.GetName(i), secretKey) + """" + varSeparator
                                                Else
                                                    varText += aes.Encode(DR.GetName(i), secretKey) + varSeparator
                                                End If
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    If encapsulation = True Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, aes.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                    End If
                                                Else
                                                    'Formatted datetime
                                                    If encapsulation = True Then
                                                        If TypeOf (DR(x)) Is DateTime Then
                                                            varText += IIf(IsDBNull(DR(x)) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                        Else
                                                            varText += IIf(IsDBNull(DR(x)) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                        End If
                                                    Else
                                                        If TypeOf (DR(x)) Is DateTime Then
                                                            varText += IIf(IsDBNull(DR(x)) = True, varSeparator, aes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                        Else
                                                            varText += IIf(IsDBNull(DR(x)) = True, varSeparator, aes.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk excel
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export excel</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                        ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toExcel("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xlsx",Nothing,Nothing, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toExcel("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xlsx",Nothing,Nothing, , ,{param1,param2})
                        ''' </example>
                        Public Function ToExcel(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String,
                                                Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                                Optional formatDateTime As String = Nothing,
                                                Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                                Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varExcelApp As Excels.Application
                                    Dim varExcelWorkBook As Excels.Workbook
                                    Dim varExcelWorkSheet As Excels.Worksheet
                                    Dim misValue As Object = System.Reflection.Missing.Value

                                    varExcelApp = New Excels.Application
                                    varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                                    varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                                    'Menambah header
                                    For i As Integer = 0 To DR.FieldCount - 1
                                        varExcelWorkSheet.Cells(1, i + 1) = aes.Encode(DR.GetName(i), secretKey)
                                    Next
                                    'Menambah row
                                    Dim processed As Integer = 0
                                    Do While (DR.Read())
                                        For j As Integer = 0 To DR.FieldCount - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, aes.Encode("", secretKey), aes.Encode(DR(j).ToString, secretKey))
                                            Else
                                                'Formatted datetime
                                                If TypeOf (DR(j)) Is DateTime Then
                                                    varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, aes.Encode("", secretKey), aes.Encode(DirectCast(DR(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                                Else
                                                    varExcelWorkSheet.Cells(processed + 2, j + 1) = IIf(IsDBNull(DR(j)) = True, aes.Encode("", secretKey), aes.Encode(DR(j).ToString, secretKey))
                                                End If
                                            End If
                                        Next
                                        processed += 1
                                    Loop


                                    If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                                    If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                                    varExcelWorkSheet.SaveAs(pathSaveFile)
                                    varExcelWorkBook.Close()
                                    varExcelApp.Quit()

                                    releaseObject(varExcelApp)
                                    releaseObject(varExcelWorkBook)
                                    releaseObject(varExcelWorkSheet)
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk csv
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export csv</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toCSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.csv",10,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toCSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.csv",10,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToCSV(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                              Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = "," 'comma
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                varText += """" + aes.Encode(DR.GetName(i), secretKey) + """" + varSeparator
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += IIf(IsDBNull(DR(x)) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DR(x).ToString, secretKey) + """" + varSeparator).ToString
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk tsv
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export tsv</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="lastRow">Untuk menghilangkan baris kosong terakhir maka diperlukan informasi total row. Default = 0</param>
                        ''' <param name="useHeader">Gunakan Header? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.tsv",10,True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toTSV("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.tsv",10,True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToTSV(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String, Optional lastRow As Integer = 0,
                                              Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varSeparator As String = Chr(9) 'Tab
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile
                                        If useHeader = True Then
                                            'Menambah header
                                            varText = Nothing
                                            For i As Integer = 0 To DR.FieldCount - 1
                                                varText += aes.Encode(DR.GetName(i), secretKey) + varSeparator
                                            Next

                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            .WriteLine()
                                        End If

                                        'Menambah row
                                        Dim processed As Integer = 0
                                        Do While (DR.Read())
                                            varText = Nothing
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += IIf(IsDBNull(DR(x)) = True, varSeparator, aes.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, aes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                    Else
                                                        varText += IIf(IsDBNull(DR(x)) = True, varSeparator, aes.Encode(DR(x).ToString, secretKey) + varSeparator).ToString
                                                    End If
                                                End If
                                            Next
                                            varText = Mid(varText, 1, varText.Length - 1)
                                            .Write(varText)
                                            If lastRow - 1 <> processed Then
                                                .WriteLine()
                                            End If
                                            processed += 1
                                        Loop
                                        .Close()
                                        Return True
                                    End With
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk html
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export html</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toHTML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.html","#CCCCCC", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toHTML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.html","#CCCCCC", , ,{param1,param2})
                        ''' </example>
                        Public Function ToHTML(ByVal sqlQuery As String, ByVal pathSaveFile As String, ByVal secretKey As String,
                                               Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows Then
                                    Dim varText As String
                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    With varTargetFile

                                        'Menyimpan header
                                        varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                        For i As Integer = 0 To DR.FieldCount - 1
                                            varText += "<TH bgcolor='" + bgColor + "' align='center'>" + aes.Encode(DR.GetName(i), secretKey) + "</TH>" + vbCrLf
                                        Next
                                        varText += "</TR>" + vbCrLf
                                        .Write(varText)
                                        .WriteLine()

                                        'Menambah row
                                        Do While (DR.Read())
                                            varText = "<TR>" + vbCrLf
                                            For x As Integer = 0 To DR.FieldCount - 1
                                                If formatDateTime = Nothing Then
                                                    'Standar datetime
                                                    varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, aes.Encode("", secretKey), aes.Encode(DR(x).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                                Else
                                                    'Formatted datetime
                                                    If TypeOf (DR(x)) Is DateTime Then
                                                        varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, aes.Encode("", secretKey), aes.Encode(DirectCast(DR(x), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                                    Else
                                                        varText += "<TD>" + IIf(IsDBNull(DR(x)) = True, aes.Encode("", secretKey), aes.Encode(DR(x).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                                    End If
                                                End If
                                            Next
                                            varText += "</TR>" + vbCrLf
                                            .Write(varText)
                                            .WriteLine()
                                        Loop
                                        varText = "</TABLE>"
                                        .Write(varText)
                                        .Close()
                                    End With
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk xml
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export xml</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="tableName">Nama table data</param>
                        ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toXML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xml","data_daerah",True, ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toXML("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.xml","data_daerah",True, , ,{param1,param2})
                        ''' </example>
                        Public Function ToXML(ByVal sqlQuery As String, pathSaveFile As String, secretKey As String, tableName As String,
                                              Optional writeSchema As Boolean = True, Optional formatDateTime As String = Nothing,
                                              Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                              Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()
                                If DR.HasRows = True Then
                                    Dim DT As New DataTable
                                    DT.Load(DR)

                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = aes.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = aes.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = aes.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    DTEncrypt.TableName = tableName
                                    If writeSchema = True Then
                                        DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                                    Else
                                        DTEncrypt.WriteXml(pathSaveFile)
                                    End If
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

                        ''' <summary>
                        ''' Membuat Query secara direct stream dan menyimpan hasilnya dalam bentuk json
                        ''' </summary>
                        ''' <param name="sqlQuery">SQL Query</param>
                        ''' <param name="pathSaveFile">Full path lokasi export json</param>
                        ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                        ''' <param name="tableName">Nama table data</param>
                        ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
                        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' 
                        ''' <example>
                        ''' Contoh cara menggunakan ParameterAddWithValue, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", "KAMPUNG TENGAH"}
                        ''' Dim param2() As Object = {"@kecamatan", "KAMPUNG TENGAH"}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toJSON("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.js","data_daerah", ,{param1,param2})
                        ''' 
                        ''' Contoh cara menggunakan ParameterAdd, sebagai berikut:
                        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
                        ''' Dim param1() As Object = {"@kelurahan", MySqlDbType.VarChar ,255, "KAMPUNG TENGAH",""}
                        ''' Dim param2() As Object = {"@kecamatan", MySqlDbType.VarChar, 255, "KAMPUNG TENGAH",""}
                        '''
                        ''' Dim DB As New momolite.data.MySQL.BuildQuery.Direct 
                        ''' DB.toJSON("select * from data_daerah where kelurahan=@kelurahan or kecamatan=@kecamatan","C:\test.js","data_daerah", , ,{param1,param2})
                        ''' </example>
                        Public Function ToJSON(ByVal sqlQuery As String, pathSaveFile As String, ByVal secretKey As String, Optional tableName As String = Nothing,
                                               Optional formatDateTime As String = Nothing,
                                               Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                               Optional showException As Boolean = True) As Boolean
                            Try
                                Dim DR As MySqlDataReader
                                'Membuka koneksi database
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                'Proses menjalankan Query
                                Dim sqlCommand As New MySqlCommand

                                sqlCommand.Connection = sqlConn
                                sqlCommand.CommandType = CommandType.Text
                                sqlCommand.CommandText = sqlQuery
                                sqlCommand.CommandTimeout = TimeOut
                                If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                                    For i As Integer = 0 To paramAddWithValue.Length - 1
                                        CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                                    Next
                                ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                                    For i As Integer = 0 To paramAdd.Length - 1
                                        CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                                    Next
                                End If

                                DR = sqlCommand.ExecuteReader()

                                If DR.HasRows = True Then
                                    Dim DT As New DataTable
                                    DT.Load(DR)

                                    'Proses Enkripsi
                                    Dim DTEncrypt As New DataTable

                                    For Each col As DataColumn In DT.Columns
                                        DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                                    Next

                                    For Each drow As DataRow In DT.Rows
                                        Dim drNew As DataRow = DTEncrypt.NewRow()
                                        For i As Integer = 0 To DT.Columns.Count - 1
                                            If formatDateTime = Nothing Then
                                                'Standar datetime
                                                drNew(i) = aes.Encode(drow(i).ToString, secretKey)
                                            Else
                                                'Formatted datetime
                                                If TypeOf (drow(i)) Is DateTime Then
                                                    drNew(i) = aes.Encode(DirectCast(drow(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                                Else
                                                    drNew(i) = aes.Encode(drow(i).ToString, secretKey)
                                                End If
                                            End If
                                        Next
                                        DTEncrypt.Rows.Add(drNew)
                                    Next

                                    Dim varTargetFile As IO.StreamWriter
                                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                    If tableName = Nothing Then
                                        varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented))
                                    Else
                                        DTEncrypt.TableName = tableName
                                        Dim ds As New DataSet
                                        ds.Tables.Add(DTEncrypt)
                                        varTargetFile.Write(Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented))
                                    End If
                                    varTargetFile.Close()
                                    Return True
                                Else
                                    Return False
                                End If
                                DR.Close()
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                'Menutup koneksi database
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function

#End Region

#Region "Status"
                        ''' <summary>
                        ''' Check status koneksi database. [AutoConnection tidak akan menampilkan status yang sedang terjadi]
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        ''' <returns>Boolean</returns>
                        ''' <remarks>Untuk verifikasi status koneksi database</remarks>
                        Public Function ConnectionStatus(Optional showException As Boolean = True) As Boolean
                            Try
                                If _autoConnection Then
                                    If Connection Is Nothing Then
                                        If ConnectionString <> Nothing Then
                                            sqlConn = New MySqlConnection(ConnectionString)
                                            sqlConn.Open()
                                        Else
                                            sqlConn = New MySqlConnection(SharedConnectionString)
                                            sqlConn.Open()
                                        End If
                                    Else
                                        sqlConn = Connection
                                    End If
                                Else
                                    sqlConn = _iconnection
                                End If

                                If sqlConn.State = ConnectionState.Open Then Return True Else Return False
                            Catch exs As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(exs, globals.Dev.Icons.Errors, NameClass)
                                Return False
                            Finally
                                If Connection Is Nothing AndAlso _autoConnection Then
                                    If sqlConn IsNot Nothing Then
                                        sqlConn.Close()
                                        sqlConn.Dispose()
                                    End If
                                End If
                                GC.Collect()
                            End Try
                        End Function
#End Region

#Region "Manual Connection"
                        ''' <summary>
                        ''' Membuka koneksi database secara manual
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        Public Sub OpenConnection(Optional showException As Boolean = True)
                            Try
                                If _autoConnection = False Then
                                    If ConnectionString <> Nothing Then
                                        _iconnection = New MySqlConnection(ConnectionString)
                                    Else
                                        _iconnection = New MySqlConnection(SharedConnectionString)
                                    End If
                                    _iconnection.Open()
                                End If
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                                _iconnection = Nothing
                            End Try
                        End Sub

                        ''' <summary>
                        ''' Menutup koneksi database secara manual
                        ''' </summary>
                        ''' <param name="showException">Tampilkan log exception? Default = True</param>
                        Public Sub CloseConnection(Optional showException As Boolean = True)
                            Try
                                If _autoConnection = False Then
                                    If _iconnection IsNot Nothing Then
                                        If _iconnection.State = ConnectionState.Open Then
                                            _iconnection.Close()
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                            End Try
                        End Sub
#End Region
                    End Class
                End Class
            End Class

        End Class

#Region "Execute Query"
        ''' <summary>
        ''' Mengeksekusi query via command line
        ''' </summary>
        ''' <param name="sqlQuery">SQL Query</param>
        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' 
        ''' <example>
        ''' Cara menambahkan parameter AddWithValue, sebagai berikut:
        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
        ''' Dim param1() As Object = {"@nama", "AZIZ"}
        ''' Dim param2() As Object = {"@pass", "12345"}
        ''' 
        ''' dim DB as new momolite.data.MySQL
        ''' DB.ExecuteQuery("insert into user values(@nama,@pass)",{param1,param2})
        ''' 
        ''' Cara jika menambahkan parameter Add, sebagai berikut:
        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
        ''' dim param1() As Object = {"@nama", MySqlDbType.VarChar ,255, "AZIZ",""}
        ''' dim param2() As Object = { "@pass", MySqlDbType.VarChar, 255, "12345",""};
        ''' 
        ''' dim DB as new momolite.data.MySQL
        ''' DB.ExecuteQuery("insert into user values(@nama,@pass)", ,{param1,param2})
        ''' </example>
        Public Function ExecuteQuery(ByVal sqlQuery As String, Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                Optional showException As Boolean = True) As Boolean
            Dim trans As MySqlTransaction = Nothing
            Try
                'Membuka koneksi database
                If _autoConnection Then
                    If Connection Is Nothing Then
                        If ConnectionString <> Nothing Then
                            sqlConn = New MySqlConnection(ConnectionString)
                            sqlConn.Open()
                        Else
                            sqlConn = New MySqlConnection(SharedConnectionString)
                            sqlConn.Open()
                        End If
                    Else
                        sqlConn = Connection
                    End If
                Else
                    sqlConn = _iconnection
                End If

                'Gunakan Transaction
                If UseTransaction = True Then trans = sqlConn.BeginTransaction(IsolationLevel.ReadCommitted)

                'Proses eksekusi command database
                Dim sqlCommand As New MySqlCommand
                With sqlCommand
                    .Connection = sqlConn
                    .CommandType = CommandType.Text
                    .CommandText = sqlQuery
                    .CommandTimeout = TimeOut
                    If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                        For i As Integer = 0 To paramAddWithValue.Length - 1
                            CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                        Next
                    ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                        For i As Integer = 0 To paramAdd.Length - 1
                            CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                        Next
                    End If
                    If UseTransaction = True Then trans.Commit()
                    .ExecuteNonQuery()
                End With
                Return True
            Catch ex As Exception
                'Proses menampilkan pesan jika terjadi error
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                If UseTransaction = True Then trans.Rollback()
                Return False
            Finally
                'Menutup koneksi database
                If Connection Is Nothing AndAlso _autoConnection Then
                    If sqlConn IsNot Nothing Then
                        sqlConn.Close()
                        sqlConn.Dispose()
                    End If
                End If
                GC.Collect()
            End Try
        End Function

        ''' <summary>
        ''' Mengeksekusi query via command line
        ''' </summary>
        ''' <param name="sqlQuery">SQL Query</param>
        ''' <param name="paramAddWithValue">Menambahkan Parameter AddWithValue. Default = Nothing.</param>
        ''' <param name="paramAdd">Menambahkan Parameter Add. Default = Nothing.</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' 
        ''' <example>
        ''' Cara menambahkan parameter AddWithValue, sebagai berikut:
        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
        ''' Dim param1() As Object = {"@nama", "AZIZ"}
        ''' Dim param2() As Object = {"@pass", "12345"}
        ''' 
        ''' dim DB as new momolite.data.MySQL
        ''' DB.ExecuteAffectRow("insert into user values(@nama,@pass)",{param1,param2})
        ''' 
        ''' Cara jika menambahkan parameter Add, sebagai berikut:
        ''' momolite.data.MySQL.SharedConnectionString = "server=127.0.0.1;port=3306;uid=root;pwd=root;database=momo;ConvertZeroDateTime=True;AllowUserVariables=True;"
        ''' dim param1() As Object = {"@nama", MySqlDbType.VarChar ,255, "AZIZ",""}
        ''' dim param2() As Object = { "@pass", MySqlDbType.VarChar, 255, "12345",""};
        ''' 
        ''' dim DB as new momolite.data.MySQL
        ''' DB.ExecuteAffectRow("insert into user values(@nama,@pass)", ,{param1,param2})
        ''' </example>
        Public Function ExecuteAffectRow(ByVal sqlQuery As String, Optional paramAddWithValue()() As Object = Nothing, Optional paramAdd()() As Object = Nothing,
                                Optional showException As Boolean = True) As Integer
            Dim trans As MySqlTransaction = Nothing
            Try
                'Membuka koneksi database
                If _autoConnection Then
                    If Connection Is Nothing Then
                        If ConnectionString <> Nothing Then
                            sqlConn = New MySqlConnection(ConnectionString)
                            sqlConn.Open()
                        Else
                            sqlConn = New MySqlConnection(SharedConnectionString)
                            sqlConn.Open()
                        End If
                    Else
                        sqlConn = Connection
                    End If
                Else
                    sqlConn = _iconnection
                End If

                'Gunakan Transaction
                If UseTransaction = True Then trans = sqlConn.BeginTransaction(IsolationLevel.ReadCommitted)

                'Proses eksekusi command database
                Dim sqlCommand As New MySqlCommand
                With sqlCommand
                    .Connection = sqlConn
                    .CommandType = CommandType.Text
                    .CommandText = sqlQuery
                    .CommandTimeout = TimeOut
                    If paramAddWithValue IsNot Nothing And paramAdd Is Nothing Then
                        For i As Integer = 0 To paramAddWithValue.Length - 1
                            CmdParameter(sqlCommand, paramAddWithValue(i)(0).ToString, paramAddWithValue(i)(1))
                        Next
                    ElseIf paramAddWithValue Is Nothing And paramAdd IsNot Nothing Then
                        For i As Integer = 0 To paramAdd.Length - 1
                            CmdParameterAdvanced(sqlCommand, paramAdd(i)(0).ToString, paramAdd(i)(1), paramAdd(i)(2), paramAdd(i)(3), paramAdd(i)(4))
                        Next
                    End If
                    If UseTransaction = True Then trans.Commit()
                    Return .ExecuteNonQuery()
                End With
            Catch ex As Exception
                'Proses menampilkan pesan jika terjadi error
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                If UseTransaction = True Then trans.Rollback()
                Return 0
            Finally
                'Menutup koneksi database
                If Connection Is Nothing AndAlso _autoConnection Then
                    If sqlConn IsNot Nothing Then
                        sqlConn.Close()
                        sqlConn.Dispose()
                    End If
                End If
                GC.Collect()
            End Try
        End Function
#End Region

#Region "Status"
        ''' <summary>
        ''' Check status koneksi database. [AutoConnection tidak akan menampilkan status yang sedang terjadi]
        ''' </summary>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>Boolean</returns>
        ''' <remarks>Untuk verifikasi status koneksi database</remarks>
        Public Function ConnectionStatus(Optional showException As Boolean = True) As Boolean
            Try
                If _autoConnection Then
                    If Connection Is Nothing Then
                        If ConnectionString <> Nothing Then
                            sqlConn = New MySqlConnection(ConnectionString)
                            sqlConn.Open()
                        Else
                            sqlConn = New MySqlConnection(SharedConnectionString)
                            sqlConn.Open()
                        End If
                    Else
                        sqlConn = Connection
                    End If
                Else
                    sqlConn = _iconnection
                End If

                If sqlConn.State = ConnectionState.Open Then Return True Else Return False
            Catch exs As Exception
                If showException = True Then momolite.globals.Dev.CatchException(exs, globals.Dev.Icons.Errors, NameClass)
                Return False
            Finally
                If Connection Is Nothing AndAlso _autoConnection Then
                    If sqlConn IsNot Nothing Then
                        sqlConn.Close()
                        sqlConn.Dispose()
                    End If
                End If
                GC.Collect()
            End Try
        End Function
#End Region

#Region "Manual Connection"
        ''' <summary>
        ''' Membuka koneksi database secara manual
        ''' </summary>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        Public Sub OpenConnection(Optional showException As Boolean = True)
            Try
                If _autoConnection = False Then
                    If ConnectionString <> Nothing Then
                        _iconnection = New MySqlConnection(ConnectionString)
                    Else
                        _iconnection = New MySqlConnection(SharedConnectionString)
                    End If
                    _iconnection.Open()
                End If
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
                _iconnection = Nothing
            End Try
        End Sub

        ''' <summary>
        ''' Menutup koneksi database secara manual
        ''' </summary>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        Public Sub CloseConnection(Optional showException As Boolean = True)
            Try
                If _autoConnection = False Then
                    If _iconnection IsNot Nothing Then
                        If _iconnection.State = ConnectionState.Open Then
                            _iconnection.Close()
                        End If
                    End If
                End If
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, NameClass)
            End Try
        End Sub
#End Region

#Region "Private Function"
        Private Shared Function CmdParameter(command As MySqlCommand, parameterName As String, value As Object) As Object
            Try
                Return command.Parameters.AddWithValue(parameterName, value)
            Catch ex As Exception
                momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.MySQL")
                Return Nothing
            End Try
        End Function

        Private Shared Sub CmdParameterAdvanced(command As MySqlCommand, parameterName As String, dataType As Object, size As Object, values As Object, Optional sourceColumn As Object = Nothing)
            Try
                If sourceColumn IsNot Nothing Then
                    command.Parameters.Add(parameterName, CType(dataType, MySqlDbType), CType(size, Integer), CType(sourceColumn, String)).Value = values
                ElseIf sourceColumn Is Nothing Then
                    command.Parameters.Add(parameterName, CType(dataType, MySqlDbType), CType(size, Integer)).Value = values
                End If
            Catch ex As Exception
                momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.MYSQL")
            End Try
        End Sub

        Private Shared Function AdapterParameter(dataAdapter As MySqlDataAdapter, parameterName As String, value As Object) As Object
            Try
                Return dataAdapter.SelectCommand.Parameters.AddWithValue(parameterName, value)
            Catch ex As Exception
                momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.MySQL")
                Return Nothing
            End Try
        End Function

        Private Shared Sub AdapterParameterAdvanced(dataAdapter As MySqlDataAdapter, parameterName As String, dataType As Object, size As Object, values As Object, Optional sourceColumn As Object = Nothing)
            Try
                If sourceColumn IsNot Nothing Then
                    dataAdapter.SelectCommand.Parameters.Add(parameterName, CType(dataType, MySqlDbType), CType(size, Integer), CType(sourceColumn, String)).Value = values
                ElseIf sourceColumn Is Nothing Then
                    dataAdapter.SelectCommand.Parameters.Add(parameterName, CType(dataType, MySqlDbType), CType(size, Integer)).Value = values
                End If
            Catch ex As Exception
                momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.MySQL")
            End Try
        End Sub

#End Region

#Region "Release Object"
        ''' <summary>
        ''' Melepas object dari COM Windows
        ''' </summary>
        ''' <param name="obj">Nama object</param>
        Private Shared Sub releaseObject(ByVal obj As Object)
            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                obj = Nothing
            Finally
                GC.Collect()
            End Try
        End Sub
#End Region

    End Class
End Namespace