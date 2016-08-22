
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
Imports System.Windows.Forms 
Namespace data
    ''' <summary>Class Export</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>31 Juli 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>2.8.0</version>
    ''' <remarks>Microsoft Office minimal 2007, terutama Excel harus sudah terinstall</remarks>
    ''' <requirement>
    ''' - Imports System.Windows.Forms
    ''' - Imports Excel = Microsoft.Office.Interop.Excel
    ''' </requirement>
    Public Class Export

        ''' <summary>
        ''' Class export dari objek dataset
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

#Region "DataSet"
            ''' <summary>
            ''' Export DataSet ke file Teks
            ''' </summary>
            ''' <param name="ds">Nama object datatable</param>
            ''' <param name="pathSaveFile">Lokasi path file txt</param>
            ''' <param name="table">Index table Dataset</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
            ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToText(ByVal ds As DataSet, ByVal pathSaveFile As String, Optional table As Integer = 0, Optional delimiter As String = ",",
                                   Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True, Optional formatDateTime As String = Nothing,
                                   Optional showException As Boolean = True) As Boolean
                Try
                    Dim DT As New DataTable
                    DT = ds.Tables(table)
                    If DT.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = delimiter
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As DataColumn In DT.Columns
                                If encapsulation = True Then
                                    varText += """" + column.ToString.Replace(Chr(34), "") + """" + varSeparator
                                Else
                                    varText += column.ToString.Replace(varSeparator, "") + varSeparator
                                End If
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If

                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = DT.Rows.Count
                        End If

                        'Menyimpan data
                        For Each row As DataRow In DT.Rows
                            varText = Nothing
                            For i As Integer = 0 To DT.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    If encapsulation = True Then
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + row.Item(i).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                    Else
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, row.Item(i).ToString.Replace(varSeparator, "") + varSeparator).ToString
                                    End If
                                Else
                                    'Formatted datetime
                                    If encapsulation = True Then
                                        If TypeOf (row.Item(i)) Is DateTime Then
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(34), "") + """" + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + row.Item(i).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                        End If
                                    Else
                                        If TypeOf (row.Item(i)) Is DateTime Then
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(varSeparator, "") + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, row.Item(i).ToString.Replace(varSeparator, "") + varSeparator).ToString
                                        End If
                                    End If
                                End If
                            Next

                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                .WriteLine()
                            End If

                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataSet ke Excel
            ''' </summary>
            ''' <param name="ds">Nama object dataset</param>
            ''' <param name="pathSaveFile">Lokasi path file excel</param>
            ''' <param name="table">Index table Dataset</param>
            ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
            ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToExcel(ByVal ds As DataSet, ByVal pathSaveFile As String, Optional table As Integer = 0, Optional ByVal usePassword As String = Nothing,
                                    Optional ByVal passwordWorkBook As String = Nothing, Optional formatDateTime As String = Nothing,
                                    Optional showException As Boolean = True) As Boolean
                Try
                    Dim DT As New DataTable
                    DT = ds.Tables(table)
                    If DT.Rows.Count = 0 Then Return False
                    Dim varExcelApp As Excels.Application
                    Dim varExcelWorkBook As Excels.Workbook
                    Dim varExcelWorkSheet As Excels.Worksheet
                    Dim misValue As Object = System.Reflection.Missing.Value

                    varExcelApp = New Excels.Application
                    varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                    varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                    'Menambah header
                    For i As Integer = 0 To DT.Columns.Count - 1
                        varExcelWorkSheet.Cells(1, i + 1) = DT.Columns(i).ColumnName
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = DT.Rows.Count
                    End If
                    'Menambah data
                    For h As Integer = 0 To DT.Rows.Count - 1
                        For j As Integer = 0 To DT.Columns.Count - 1
                            Application.DoEvents()
                            If formatDateTime = Nothing Then
                                'Standar datetime
                                varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, "", DT.Rows(h).Item(j))
                            Else
                                'Formatted datetime
                                If TypeOf (DT.Rows(h).Item(j)) Is DateTime Then
                                    varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, "", DirectCast(DT.Rows(h).Item(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                Else
                                    varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, "", DT.Rows(h).Item(j))
                                End If
                            End If

                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next
                    If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                    If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                    varExcelWorkSheet.SaveAs(pathSaveFile)
                    varExcelWorkBook.Close()
                    varExcelApp.Quit()

                    releaseObject(varExcelApp)
                    releaseObject(varExcelWorkBook)
                    releaseObject(varExcelWorkSheet)
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataSet")
                    Return False
                End Try
            End Function

            ''' <summary>
            ''' Export DataSet ke CSV
            ''' </summary>
            ''' <param name="ds">Nama objek DataSet</param>
            ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
            ''' <param name="table">Index table Dataset</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToCSV(ByVal ds As DataSet, ByVal pathSaveFile As String, Optional table As Integer = 0,
                                  Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                  Optional showException As Boolean = True) As Boolean
                Try
                    Dim DT As New DataTable
                    DT = ds.Tables(table)
                    If DT.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = "," 'comma
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As DataColumn In DT.Columns
                                varText += """" + column.ToString.Replace(Chr(34), "") + """" + varSeparator
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = DT.Rows.Count
                        End If
                        'Menyimpan data
                        For Each row As DataRow In DT.Rows
                            varText = Nothing
                            For i As Integer = 0 To DT.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + row.Item(i).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                Else
                                    'Formatted datetime
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(34), "") + """" + varSeparator).ToString
                                    Else
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + row.Item(i).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                    End If
                                End If
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataSet ke TSV
            ''' </summary>
            ''' <param name="ds">Nama objek DataSet</param>
            ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
            ''' <param name="table">Index table DataSet</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToTSV(ByVal ds As DataSet, ByVal pathSaveFile As String, Optional table As Integer = 0,
                                  Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                  Optional showException As Boolean = True) As Boolean
                Try
                    Dim DT As New DataTable
                    DT = ds.Tables(table)
                    If DT.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = Chr(9) 'tab
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As DataColumn In DT.Columns
                                varText += column.ToString.Replace(Chr(9), "") + varSeparator
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = DT.Rows.Count
                        End If
                        'Menyimpan data
                        For Each row As DataRow In DT.Rows
                            varText = Nothing
                            For i As Integer = 0 To DT.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, row.Item(i).ToString.Replace(Chr(9), "") + varSeparator).ToString
                                Else
                                    'Formatted datetime
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(9), "") + varSeparator).ToString
                                    Else
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, row.Item(i).ToString.Replace(Chr(9), "") + varSeparator).ToString
                                    End If
                                End If

                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataSet ke HTML
            ''' </summary>
            ''' <param name="ds">Nama objek DataSet</param>
            ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
            ''' <param name="table">Index table DataSet</param>
            ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToHTML(ByVal ds As DataSet, ByVal pathSaveFile As String, Optional table As Integer = 0,
                                   Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                   Optional showException As Boolean = True) As Boolean
                Try
                    Dim DT As New DataTable
                    DT = ds.Tables(table)
                    If DT.Rows.Count = 0 Then Return False
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        'Menyimpan header
                        varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                        For Each column As DataColumn In DT.Columns
                            varText += "<TH bgcolor='" + bgColor + "' align='center'>" + column.ToString + "</TH>" + vbCrLf
                        Next
                        varText += "</TR>" + vbCrLf
                        .Write(varText)
                        .WriteLine()
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = DT.Rows.Count
                        End If
                        'Menyimpan data
                        For Each row As DataRow In DT.Rows
                            varText = "<TR>" + vbCrLf
                            For i As Integer = 0 To DT.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, "", row.Item(i).ToString).ToString + "</TD>" + vbCrLf
                                Else
                                    'Formatted datetime
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, "", DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)).ToString + "</TD>" + vbCrLf
                                    Else
                                        varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, "", row.Item(i).ToString).ToString + "</TD>" + vbCrLf
                                    End If
                                End If

                            Next
                            varText += "</TR>" + vbCrLf
                            .Write(varText)
                            .WriteLine()
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        varText = "</TABLE>"
                        .Write(varText)
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataSet ke XML
            ''' </summary>
            ''' <param name="ds">Nama objek dataset</param>
            ''' <param name="pathSaveFile">Full path lokasi export XML</param>
            ''' <param name="tableName">Nama table data</param>
            ''' <param name="writeSchema">Gunakan schema? Default = True</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToXML(ds As DataSet, pathSaveFile As String, tableName As String,
                                  Optional writeSchema As Boolean = True, Optional formatDateTime As String = Nothing,
                                  Optional showException As Boolean = True) As Boolean
                Try
                    If formatDateTime = Nothing Then
                        If Progressbar IsNot Nothing Then 'iteration way
                            Dim DT As New DataTable
                            DT = ds.Tables(tableName)
                            Dim DTFormat As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTFormat.Columns.Add(col.ToString)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTFormat.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = dr(i).ToString
                                Next
                                DTFormat.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DTFormat.TableName = tableName
                            If writeSchema = True Then
                                DTFormat.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DTFormat.WriteXml(pathSaveFile)
                            End If
                        Else 'simple way
                            ds.DataSetName = tableName
                            If writeSchema = True Then
                                ds.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                ds.WriteXml(pathSaveFile)
                            End If
                        End If

                    Else
                        'Formatted DateTime
                        Dim DT As New DataTable
                        DT = ds.Tables(tableName)
                        Dim DTFormat As New DataTable
                        For Each col As DataColumn In DT.Columns
                            DTFormat.Columns.Add(col.ToString)
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = DT.Rows.Count
                        End If
                        'Menyimpan data
                        For Each dr As DataRow In DT.Rows
                            Dim drNew As DataRow = DTFormat.NewRow()
                            For i As Integer = 0 To DT.Columns.Count - 1
                                If TypeOf (dr(i)) Is DateTime Then
                                    drNew(i) = DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                Else
                                    drNew(i) = dr(i).ToString
                                End If

                            Next
                            DTFormat.Rows.Add(drNew)
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        DTFormat.TableName = tableName
                        If writeSchema = True Then
                            DTFormat.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                        Else
                            DTFormat.WriteXml(pathSaveFile)
                        End If
                    End If

                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataSet")
                    Return False
                End Try
            End Function

            ''' <summary>
            ''' Export DataSet ke JSON
            ''' </summary>
            ''' <param name="ds">Nama objek dataset</param>
            ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
            ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
            ''' <param name="table">Jika dalam membuat json gunakan index table di dataset untuk memilih table yang akan dibuat json. Default = 0</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToJSON(ds As DataSet, pathSaveFile As String, Optional tableName As String = Nothing,
                                   Optional table As Integer = 0, Optional formatDateTime As String = Nothing,
                                   Optional showException As Boolean = True) As Boolean
                Try
                    If formatDateTime = Nothing Then
                        If Progressbar IsNot Nothing Then 'iteration way
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            Dim DTFormat As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTFormat.Columns.Add(col.ToString)
                            Next

                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTFormat.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    drNew(i) = dr(i).ToString
                                Next
                                DTFormat.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTFormat, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DTFormat.TableName = tableName
                                Dim ds1 As New DataSet
                                ds1.Tables.Add(DTFormat)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds1, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Else 'simple way
                            'Standar DateTime
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                Dim DT As New DataTable
                                DT = ds.Tables(table)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                            Else
                                ds.DataSetName = tableName
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        End If
                    Else
                        'Formatted DateTime

                        'proses format datetime
                        Dim DT As New DataTable
                        DT = ds.Tables(table)
                        Dim DTFormat As New DataTable
                        For Each col As DataColumn In DT.Columns
                            DTFormat.Columns.Add(col.ToString)
                        Next

                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = DT.Rows.Count
                        End If
                        'Menyimpan data
                        For Each dr As DataRow In DT.Rows
                            Dim drNew As DataRow = DTFormat.NewRow()
                            For i As Integer = 0 To DT.Columns.Count - 1
                                If TypeOf (dr(i)) Is DateTime Then
                                    drNew(i) = DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                Else
                                    drNew(i) = dr(i).ToString
                                End If
                            Next
                            DTFormat.Rows.Add(drNew)
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next

                        'proses export
                        Dim dataJSON As String = Nothing
                        If tableName = Nothing Then
                            dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTFormat, Newtonsoft.Json.Formatting.Indented)
                        Else
                            DTFormat.TableName = tableName
                            Dim ds1 As New DataSet
                            ds1.Tables.Add(DTFormat)
                            dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds1, Newtonsoft.Json.Formatting.Indented)
                        End If
                        If dataJSON <> Nothing Then
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            varTargetFile.Write(dataJSON)
                            varTargetFile.Close()
                            Return True
                        Else
                            Return False
                        End If
                    End If
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataSet")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

        End Class

        ''' <summary>
        ''' Class export dari objek datatable
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

#Region "Datatable"
            ''' <summary>
            ''' Export Datatable ke file Teks
            ''' </summary>
            ''' <param name="dt">Nama object datatable</param>
            ''' <param name="pathSaveFile">Lokasi path file txt</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
            ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToText(ByVal dt As DataTable, ByVal pathSaveFile As String, Optional delimiter As String = ",",
                                   Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True, Optional formatDateTime As String = Nothing,
                                   Optional showException As Boolean = True) As Boolean
                Try
                    If dt.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = delimiter
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As DataColumn In dt.Columns
                                If encapsulation = True Then
                                    varText += """" + column.ToString.Replace(Chr(34), "") + """" + varSeparator
                                Else
                                    varText += column.ToString.Replace(varSeparator, "") + varSeparator
                                End If
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'Menyimpan data
                        For Each row As DataRow In dt.Rows
                            varText = Nothing
                            For i As Integer = 0 To dt.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    If encapsulation = True Then
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + row.Item(i).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                    Else
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, row.Item(i).ToString.Replace(varSeparator, "") + varSeparator).ToString
                                    End If
                                Else
                                    'Formatted datetime
                                    If encapsulation = True Then
                                        If TypeOf (row.Item(i)) Is DateTime Then
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(34), "") + """" + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + row.Item(i).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                        End If

                                    Else
                                        If TypeOf (row.Item(i)) Is DateTime Then
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(varSeparator, "") + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, row.Item(i).ToString.Replace(varSeparator, "") + varSeparator).ToString
                                        End If

                                    End If
                                End If
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataTable ke Excel
            ''' </summary>
            ''' <param name="dt">Nama object datatable</param>
            ''' <param name="pathSaveFile">Lokasi path file excel</param>
            ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
            ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToExcel(ByVal dt As DataTable, ByVal pathSaveFile As String, Optional ByVal usePassword As String = Nothing,
                                    Optional ByVal passwordWorkBook As String = Nothing, Optional formatDateTime As String = Nothing,
                                    Optional showException As Boolean = True) As Boolean
                Try
                    If dt.Rows.Count = 0 Then Return False
                    Dim varExcelApp As Excels.Application
                    Dim varExcelWorkBook As Excels.Workbook
                    Dim varExcelWorkSheet As Excels.Worksheet
                    Dim misValue As Object = System.Reflection.Missing.Value

                    varExcelApp = New Excels.Application
                    varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                    varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                    'Menambah header
                    For i As Integer = 0 To dt.Columns.Count - 1
                        varExcelWorkSheet.Cells(1, i + 1) = dt.Columns(i).ColumnName
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = dt.Rows.Count
                    End If
                    'Menambah data
                    For h As Integer = 0 To dt.Rows.Count - 1
                        For j As Integer = 0 To dt.Columns.Count - 1
                            Application.DoEvents()
                            If formatDateTime = Nothing Then
                                'Standar datetime
                                varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, "", dt.Rows(h).Item(j))
                            Else
                                'Formatted datetime
                                If TypeOf (dt.Rows(h).Item(j)) Is DateTime Then
                                    varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, "", DirectCast(dt.Rows(h).Item(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture))
                                Else
                                    varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, "", dt.Rows(h).Item(j))
                                End If
                            End If

                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next
                    If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                    If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                    varExcelWorkSheet.SaveAs(pathSaveFile)
                    varExcelWorkBook.Close()
                    varExcelApp.Quit()

                    releaseObject(varExcelApp)
                    releaseObject(varExcelWorkBook)
                    releaseObject(varExcelWorkSheet)
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataTable")
                    Return False
                End Try
            End Function

            ''' <summary>
            ''' Export DataTable ke CSV
            ''' </summary>
            ''' <param name="dt">Nama objek DataTable</param>
            ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToCSV(ByVal dt As DataTable, ByVal pathSaveFile As String, Optional useHeader As Boolean = True,
                                  Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    If dt.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = "," 'comma
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As DataColumn In dt.Columns
                                varText += """" + column.ToString.Replace(Chr(34), "") + """" + varSeparator
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'Menyimpan data
                        For Each row As DataRow In dt.Rows
                            varText = Nothing
                            For i As Integer = 0 To dt.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + row.Item(i).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                Else
                                    'Formatted datetime
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(34), "") + """" + varSeparator).ToString
                                    Else
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + varSeparator, """" + row.Item(i).ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                    End If
                                End If

                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataTable ke TSV
            ''' </summary>
            ''' <param name="dt">Nama objek DataTable</param>
            ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToTSV(ByVal dt As DataTable, ByVal pathSaveFile As String,
                                  Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                  Optional showException As Boolean = True) As Boolean
                Try
                    If dt.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = Chr(9) 'tab
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As DataColumn In dt.Columns
                                varText += column.ToString.Replace(Chr(9), "") + varSeparator
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'Menyimpan data
                        For Each row As DataRow In dt.Rows
                            varText = Nothing
                            For i As Integer = 0 To dt.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, row.Item(i).ToString.Replace(Chr(9), "") + varSeparator).ToString
                                Else
                                    'Formatted datetime
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture).Replace(Chr(9), "") + varSeparator).ToString
                                    Else
                                        varText += IIf(IsDBNull(row.Item(i).ToString) = True, varSeparator, row.Item(i).ToString.Replace(Chr(9), "") + varSeparator).ToString
                                    End If
                                End If
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataTable ke HTML
            ''' </summary>
            ''' <param name="dt">Nama objek DataTable</param>
            ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
            ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToHTML(ByVal dt As DataTable, ByVal pathSaveFile As String,
                                   Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                   Optional showException As Boolean = True) As Boolean
                Try
                    If dt.Rows.Count = 0 Then Return False
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        'Menyimpan header
                        varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                        For Each column As DataColumn In dt.Columns
                            varText += "<TH bgcolor='" + bgColor + "' align='center'>" + column.ToString + "</TH>" + vbCrLf
                        Next
                        varText += "</TR>" + vbCrLf
                        .Write(varText)
                        .WriteLine()
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'Menyimpan data
                        For Each row As DataRow In dt.Rows
                            varText = "<TR>" + vbCrLf
                            For i As Integer = 0 To dt.Columns.Count - 1
                                If formatDateTime = Nothing Then
                                    'Standar datetime
                                    varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, "", row.Item(i).ToString).ToString + "</TD>" + vbCrLf
                                Else
                                    'Formatted datetime
                                    If TypeOf (row.Item(i)) Is DateTime Then
                                        varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, "", DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)).ToString + "</TD>" + vbCrLf
                                    Else
                                        varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, "", row.Item(i).ToString).ToString + "</TD>" + vbCrLf
                                    End If
                                End If
                            Next
                            varText += "</TR>" + vbCrLf
                            .Write(varText)
                            .WriteLine()
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        varText = "</TABLE>"
                        .Write(varText)
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export Datatable ke XML
            ''' </summary>
            ''' <param name="dt">Nama objek datatable</param>
            ''' <param name="pathSaveFile">Full path lokasi export XML</param>
            ''' <param name="tableName">Nama table data</param>
            ''' <param name="writeSchema">Gunakan schema? Default = True</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToXML(dt As DataTable, pathSaveFile As String, tableName As String,
                                  Optional writeSchema As Boolean = True, Optional formatDateTime As String = Nothing,
                                  Optional showException As Boolean = True) As Boolean
                Try
                    If formatDateTime = Nothing Then
                        If Progressbar IsNot Nothing Then 'iteration way
                            Dim DTFormat As New DataTable
                            For Each col As DataColumn In dt.Columns
                                DTFormat.Columns.Add(col.ToString)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'proses menyimpan data
                            For Each dr As DataRow In dt.Rows
                                Dim drNew As DataRow = DTFormat.NewRow()
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    drNew(i) = dr(i).ToString
                                Next
                                DTFormat.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DTFormat.TableName = tableName
                            If writeSchema = True Then
                                DTFormat.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DTFormat.WriteXml(pathSaveFile)
                            End If
                        Else 'simple way
                            dt.TableName = tableName
                            If writeSchema = True Then
                                dt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                dt.WriteXml(pathSaveFile)
                            End If
                        End If
                    Else
                        'Formatted datetime
                        Dim DTFormat As New DataTable
                        For Each col As DataColumn In dt.Columns
                            DTFormat.Columns.Add(col.ToString)
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'proses menyimpan data
                        For Each dr As DataRow In dt.Rows
                            Dim drNew As DataRow = DTFormat.NewRow()
                            For i As Integer = 0 To dt.Columns.Count - 1
                                If TypeOf (dr(i)) Is DateTime Then
                                    drNew(i) = DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                Else
                                    drNew(i) = dr(i).ToString
                                End If

                            Next
                            DTFormat.Rows.Add(drNew)
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        DTFormat.TableName = tableName
                        If writeSchema = True Then
                            DTFormat.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                        Else
                            DTFormat.WriteXml(pathSaveFile)
                        End If
                    End If

                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataTable")
                    Return False
                End Try
            End Function

            ''' <summary>
            ''' Export Datatable ke JSON
            ''' </summary>
            ''' <param name="dt">Nama objek datatable yang telah terisi data</param>
            ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
            ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
            ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToJSON(dt As DataTable, pathSaveFile As String,
                                   Optional tableName As String = Nothing, Optional formatDateTime As String = Nothing,
                                   Optional showException As Boolean = True) As Boolean
                Try
                    If formatDateTime = Nothing Then
                        If Progressbar IsNot Nothing Then 'iteration way
                            'proses formatting
                            Dim DTFormat As New DataTable
                            For Each col As DataColumn In dt.Columns
                                DTFormat.Columns.Add(col.ToString)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'proses menyimpan data
                            For Each dr As DataRow In dt.Rows
                                Dim drNew As DataRow = DTFormat.NewRow()
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    drNew(i) = dr(i).ToString
                                Next
                                DTFormat.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTFormat, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DTFormat.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DTFormat)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Else 'simple way
                            'Standar datetime
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(dt, Newtonsoft.Json.Formatting.Indented)
                            Else
                                dt.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(dt)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        End If

                    Else
                        'Formatted datetime
                        'proses formatting
                        Dim DTFormat As New DataTable
                        For Each col As DataColumn In dt.Columns
                            DTFormat.Columns.Add(col.ToString)
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dt.Rows.Count
                        End If
                        'proses menyimpan data
                        For Each dr As DataRow In dt.Rows
                            Dim drNew As DataRow = DTFormat.NewRow()
                            For i As Integer = 0 To dt.Columns.Count - 1
                                If TypeOf (dr(i)) Is DateTime Then
                                    drNew(i) = DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture)
                                Else
                                    drNew(i) = dr(i).ToString
                                End If
                            Next
                            DTFormat.Rows.Add(drNew)
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next

                        'proses export
                        Dim dataJSON As String = Nothing
                        If tableName = Nothing Then
                            dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTFormat, Newtonsoft.Json.Formatting.Indented)
                        Else
                            DTFormat.TableName = tableName
                            Dim ds As New DataSet
                            ds.Tables.Add(DTFormat)
                            dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                        End If
                        If dataJSON <> Nothing Then
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            varTargetFile.Write(dataJSON)
                            varTargetFile.Close()
                            Return True
                        Else
                            Return False
                        End If
                    End If
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataTable")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function
#End Region

        End Class

        ''' <summary>
        ''' Class export dari objek datagridview
        ''' </summary>
        Public Class FromDataGridView
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

#Region "DataGridView"

            ''' <summary>
            ''' Export DataGridView ke file Teks
            ''' </summary>
            ''' <param name="dgv">Nama object DataGridView</param>
            ''' <param name="pathSaveFile">Lokasi path file txt</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
            ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToText(ByVal dgv As DataGridView, ByVal pathSaveFile As String, Optional delimiter As String = ",",
                                   Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True, Optional showException As Boolean = True) As Boolean
                Try
                    If dgv.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = delimiter
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        'Menyimpan header
                        If useHeader = True Then
                            varText = Nothing
                            For Each column As DataGridViewColumn In dgv.Columns
                                If encapsulation = True Then
                                    varText += """" + column.HeaderText.Replace(Chr(34), "") + """" + varSeparator
                                Else
                                    varText += column.HeaderText.Replace(varSeparator, "") + varSeparator
                                End If
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dgv.RowCount
                        End If
                        'Menyimpan data
                        For Each row As DataGridViewRow In dgv.Rows
                            varText = Nothing
                            For i As Integer = 0 To dgv.Columns.Count - 1
                                If encapsulation = True Then
                                    varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, """" + varSeparator, """" + row.Cells(i).Value.ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                                Else
                                    varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, varSeparator, row.Cells(i).Value.ToString.Replace(varSeparator, "") + varSeparator).ToString
                                End If
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataGridView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataGridView ke Excel
            ''' </summary>
            ''' <param name="dgv">Nama objek DataGridView</param>
            ''' <param name="pathSaveFile">Path lokasi hasil export excel</param>
            ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
            ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToExcel(ByVal dgv As DataGridView, ByVal pathSaveFile As String, Optional ByVal usePassword As String = Nothing,
                                    Optional ByVal passwordWorkBook As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    If dgv.RowCount = 0 Then Return False
                    Dim varExcelApp As Excels.Application
                    Dim varExcelWorkBook As Excels.Workbook
                    Dim varExcelWorkSheet As Excels.Worksheet
                    Dim misValue As Object = System.Reflection.Missing.Value

                    varExcelApp = New Excels.Application
                    varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                    varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                    'Menambah header
                    For i As Integer = 0 To dgv.ColumnCount - 1
                        varExcelWorkSheet.Cells(1, i + 1) = dgv.Columns(i).HeaderText
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = dgv.RowCount
                    End If
                    'Menambah data
                    For h As Integer = 0 To dgv.RowCount - 1
                        For j As Integer = 0 To dgv.ColumnCount - 1
                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dgv.Rows(h).Cells(j).Value.ToString()) = True, "", dgv.Rows(h).Cells(j).Value.ToString())
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next

                    If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                    If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                    varExcelWorkSheet.SaveAs(pathSaveFile)
                    varExcelWorkBook.Close()
                    varExcelApp.Quit()

                    releaseObject(varExcelApp)
                    releaseObject(varExcelWorkBook)
                    releaseObject(varExcelWorkSheet)
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataGridView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataGridView ke CSV
            ''' </summary>
            ''' <param name="dgv">Nama objek DataGridView</param>
            ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToCSV(ByVal dgv As DataGridView, ByVal pathSaveFile As String,
                                  Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                Try
                    If dgv.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = "," 'comma
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        'Menyimpan header
                        If useHeader = True Then
                            varText = Nothing
                            For Each column As DataGridViewColumn In dgv.Columns
                                varText += """" + column.HeaderText.Replace(Chr(34), "") + """" + varSeparator
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dgv.RowCount
                        End If
                        'Menyimpan data
                        For Each row As DataGridViewRow In dgv.Rows
                            varText = Nothing
                            For i As Integer = 0 To dgv.Columns.Count - 1
                                varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, """" + varSeparator, """" + row.Cells(i).Value.ToString.Replace(Chr(34), "") + """" + varSeparator).ToString
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataGridView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataGridView ke TSV
            ''' </summary>
            ''' <param name="dgv">Nama objek DataGridView</param>
            ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToTSV(ByVal dgv As DataGridView, ByVal pathSaveFile As String,
                                  Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                Try
                    If dgv.Rows.Count = 0 Then Return False
                    Dim varSeparator As String = Chr(9) 'Tab
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        'Menyimpan header
                        If useHeader = True Then
                            varText = Nothing
                            For Each column As DataGridViewColumn In dgv.Columns
                                varText += column.HeaderText.Replace(Chr(9), "") + varSeparator
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dgv.RowCount
                        End If
                        'Menyimpan data
                        For Each row As DataGridViewRow In dgv.Rows
                            varText = Nothing
                            For i As Integer = 0 To dgv.Columns.Count - 1
                                varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, varSeparator, row.Cells(i).Value.ToString.Replace(Chr(9), "") + varSeparator).ToString
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataGridView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataGridView ke HTML
            ''' </summary>
            ''' <param name="dgv">Nama objek DataGridView</param>
            ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
            ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToHTML(ByVal dgv As DataGridView, ByVal pathSaveFile As String,
                                   Optional bgColor As String = "#CCCCCC", Optional showException As Boolean = True) As Boolean
                Try
                    If dgv.RowCount = 0 Then Return False
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        'Menyimpan header
                        varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                        For Each column As DataGridViewColumn In dgv.Columns
                            varText += "<TH bgcolor='" + bgColor + "' align='center'>" + column.HeaderText + "</TH>" + vbCrLf
                        Next
                        varText += "</TR>" + vbCrLf
                        .Write(varText)
                        .WriteLine()
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = dgv.RowCount
                        End If
                        'Menyimpan data
                        For Each row As DataGridViewRow In dgv.Rows
                            varText = "<TR>" + vbCrLf
                            For i As Integer = 0 To dgv.ColumnCount - 1
                                varText += "<TD>" + IIf(IsDBNull(row.Cells(i).Value) = True, "", row.Cells(i).Value.ToString).ToString + "</TD>" + vbCrLf
                            Next
                            varText += "</TR>" + vbCrLf
                            .Write(varText)
                            .WriteLine()
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        varText = "</TABLE>"
                        .Write(varText)
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataGridView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export DataGridView ke XML
            ''' </summary>
            ''' <param name="dgv">Nama objek DataGridView</param>
            ''' <param name="pathSaveFile">Full path lokasi export XML</param>
            ''' <param name="writeSchema">Gunakan schema? Default = True</param>
            ''' <param name="tableName">Nama table data</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToXML(dgv As DataGridView, pathSaveFile As String, tableName As String,
                                  Optional writeSchema As Boolean = True, Optional showException As Boolean = True) As Boolean
                Try
                    Dim DT As New DataTable()
                    For Each col As DataGridViewColumn In dgv.Columns
                        DT.Columns.Add(col.HeaderText)
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = dgv.RowCount
                    End If
                    'Menyimpan data
                    For Each row As DataGridViewRow In dgv.Rows
                        Dim dRow As DataRow = DT.NewRow()
                        For Each cell As DataGridViewCell In row.Cells
                            dRow(cell.ColumnIndex) = cell.Value
                        Next
                        DT.Rows.Add(dRow)
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next
                    DT.TableName = tableName
                    If writeSchema = True Then
                        DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                    Else
                        DT.WriteXml(pathSaveFile)
                    End If
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataGridView")
                    Return False
                End Try
            End Function

            ''' <summary>
            ''' Export DataGridView ke JSON
            ''' </summary>
            ''' <param name="dgv">Nama objek DataGridView</param>
            ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
            ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToJSON(dgv As DataGridView, pathSaveFile As String,
                                   Optional tableName As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try

                    Dim DT As New DataTable()
                    For Each col As DataGridViewColumn In dgv.Columns
                        DT.Columns.Add(col.HeaderText)
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = dgv.RowCount
                    End If
                    'Menyimpan data
                    For Each row As DataGridViewRow In dgv.Rows
                        Dim dRow As DataRow = DT.NewRow()
                        For Each cell As DataGridViewCell In row.Cells
                            dRow(cell.ColumnIndex) = cell.Value
                        Next
                        DT.Rows.Add(dRow)
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next

                    Dim dataJSON As String = Nothing
                    If tableName = Nothing Then
                        dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                    Else
                        DT.TableName = tableName
                        Dim ds As New DataSet
                        ds.Tables.Add(DT)
                        dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                    End If
                    If dataJSON <> Nothing Then
                        Dim varTargetFile As IO.StreamWriter
                        varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                        varTargetFile.Write(dataJSON)
                        varTargetFile.Close()
                        Return True
                    Else
                        Return False
                    End If
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromDataGridView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

#End Region

        End Class

        ''' <summary>
        ''' Class export dari objek listview
        ''' </summary>
        Public Class FromListView
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

#Region "ListView"

            ''' <summary>
            ''' Export ListView ke file Teks
            ''' </summary>
            ''' <param name="lv">Nama object ListView</param>
            ''' <param name="pathSaveFile">Lokasi path file txt</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="delimiter">Karakter pemisah. Default = ","</param>
            ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToText(ByVal lv As ListView, ByVal pathSaveFile As String, Optional delimiter As String = ",",
                                   Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True, Optional showException As Boolean = True) As Boolean
                Try
                    If lv.Items.Count = 0 Then Return False
                    Dim varSeparator As String = delimiter
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As ColumnHeader In lv.Columns
                                If encapsulation = True Then
                                    varText += """" + column.Text.Replace(Chr(34), "") + """" + varSeparator
                                Else
                                    varText += column.Text.Replace(varSeparator, "") + varSeparator
                                End If

                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = lv.Items.Count
                        End If
                        'Menyimpan data
                        For Each row As ListViewItem In lv.Items
                            varText = Nothing
                            For i As Integer = 0 To lv.Columns.Count - 1
                                If encapsulation = True Then
                                    varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, """" + varSeparator, """" + row.SubItems.Item(i).Text.Replace(Chr(34), "") + """" + varSeparator).ToString
                                Else
                                    varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, varSeparator, row.SubItems.Item(i).Text.Replace(varSeparator, "") + varSeparator).ToString
                                End If
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromListView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export ListView ke Excel
            ''' </summary>
            ''' <param name="lv">Nama objek ListView</param>
            ''' <param name="pathSaveFile">Path lokasi hasil export excel</param>
            ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
            ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToExcel(ByVal lv As ListView, ByVal pathSaveFile As String, Optional ByVal usePassword As String = Nothing,
                                    Optional ByVal passwordWorkBook As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try
                    If lv.Items.Count = 0 Then Return False
                    Dim varExcelApp As Excels.Application
                    Dim varExcelWorkBook As Excels.Workbook
                    Dim varExcelWorkSheet As Excels.Worksheet
                    Dim misValue As Object = System.Reflection.Missing.Value

                    varExcelApp = New Excels.Application
                    varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                    varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                    'Menambah header
                    For i As Integer = 0 To lv.Columns.Count - 1
                        varExcelWorkSheet.Cells(1, i + 1) = lv.Columns(i).Text
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = lv.Items.Count
                    End If
                    'Menambah data
                    For h As Integer = 0 To lv.Items.Count - 1
                        For j As Integer = 0 To lv.Columns.Count - 1
                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(lv.Items.Item(h).SubItems.Item(j).Text) = True, "", lv.Items.Item(h).SubItems.Item(j).Text)
                        Next
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next

                    If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                    If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                    varExcelWorkSheet.SaveAs(pathSaveFile)
                    varExcelWorkBook.Close()
                    varExcelApp.Quit()

                    releaseObject(varExcelApp)
                    releaseObject(varExcelWorkBook)
                    releaseObject(varExcelWorkSheet)
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromListView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export ListView ke CSV
            ''' </summary>
            ''' <param name="lv">Nama objek ListView</param>
            ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToCSV(ByVal lv As ListView, ByVal pathSaveFile As String,
                                  Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                Try
                    If lv.Items.Count = 0 Then Return False
                    Dim varSeparator As String = "," 'comma
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As ColumnHeader In lv.Columns
                                varText += """" + column.Text.Replace(Chr(34), "") + """" + varSeparator
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = lv.Items.Count
                        End If
                        'Menyimpan data
                        For Each row As ListViewItem In lv.Items
                            varText = Nothing
                            For i As Integer = 0 To lv.Columns.Count - 1
                                varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, """" + varSeparator, """" + row.SubItems.Item(i).Text.Replace(Chr(34), "") + """" + varSeparator).ToString
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromListView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export ListView ke TSV
            ''' </summary>
            ''' <param name="lv">Nama objek ListView</param>
            ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
            ''' <param name="useHeader">Gunakan Header? Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToTSV(ByVal lv As ListView, ByVal pathSaveFile As String,
                                  Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                Try
                    If lv.Items.Count = 0 Then Return False
                    Dim varSeparator As String = Chr(9) 'tab
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        If useHeader = True Then
                            'Menyimpan header
                            varText = Nothing
                            For Each column As ColumnHeader In lv.Columns
                                varText = varText + column.Text.Replace(Chr(9), "") + varSeparator
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            .WriteLine()
                        End If
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = lv.Items.Count
                        End If
                        'Menyimpan data
                        For Each row As ListViewItem In lv.Items
                            varText = Nothing
                            For i As Integer = 0 To lv.Columns.Count - 1
                                varText = varText + IIf(IsDBNull(row.SubItems.Item(i).Text) = True, varSeparator, row.SubItems.Item(i).Text.Replace(Chr(9), "") + varSeparator).ToString
                            Next
                            varText = Mid(varText, 1, varText.Length - 1)
                            .Write(varText)
                            If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                .WriteLine()
                            End If
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromListView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export ListView ke HTML
            ''' </summary>
            ''' <param name="lv">Nama objek ListView</param>
            ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
            ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToHTML(ByVal lv As ListView, ByVal pathSaveFile As String,
                                   Optional bgColor As String = "#CCCCCC", Optional showException As Boolean = True) As Boolean
                Try
                    If lv.Items.Count = 0 Then Return False
                    Dim varText As String
                    Dim varTargetFile As IO.StreamWriter
                    varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                    With varTargetFile
                        'Menyimpan header
                        varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                        For Each column As ColumnHeader In lv.Columns
                            varText += "<TH bgcolor='" + bgColor + "' align='center'>" + column.Text + "</TH>" + vbCrLf
                        Next
                        varText += "</TR>" + vbCrLf
                        .Write(varText)
                        .WriteLine()
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Progressbar.Value = 0
                            Progressbar.Maximum = lv.Items.Count
                        End If
                        'Menyimpan data
                        For Each row As ListViewItem In lv.Items
                            varText = "<TR>" + vbCrLf
                            For i As Integer = 0 To lv.Columns.Count - 1
                                varText += "<TD>" + IIf(IsDBNull(row.SubItems.Item(i).Text) = True, "", row.SubItems.Item(i).Text).ToString + "</TD>" + vbCrLf
                            Next
                            varText += "</TR>" + vbCrLf
                            .Write(varText)
                            .WriteLine()
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Application.DoEvents()
                                Progressbar.Value += 1
                            End If
                        Next
                        varText = "</TABLE>"
                        .Write(varText)
                        .Close()
                    End With
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromListView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Export ListView ke XML
            ''' </summary>
            ''' <param name="lv">nama objek ListView</param>
            ''' <param name="pathSaveFile">Full path lokasi export XML</param>
            ''' <param name="tableName">Nama table data</param>
            ''' <param name="writeSchema">Gunakan schema? Default = True</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToXML(lv As ListView, pathSaveFile As String, tableName As String,
                                  Optional writeSchema As Boolean = True, Optional showException As Boolean = True) As Boolean
                Try
                    Dim DT As New DataTable
                    For Each col As ColumnHeader In lv.Columns
                        DT.Columns.Add(col.Text)
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = lv.Items.Count
                    End If
                    'menyimpan data
                    For Each item As ListViewItem In lv.Items
                        Dim row As DataRow = DT.NewRow()
                        For i As Integer = 0 To item.SubItems.Count - 1
                            row(i) = item.SubItems(i).Text
                        Next
                        DT.Rows.Add(row)
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next
                    DT.TableName = tableName
                    If writeSchema = True Then
                        DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                    Else
                        DT.WriteXml(pathSaveFile)
                    End If
                    Return True
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromListView")
                    Return False
                End Try
            End Function

            ''' <summary>
            ''' Export ListView ke JSON
            ''' </summary>
            ''' <param name="lv">nama objek ListView</param>
            ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
            ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function ToJSON(lv As ListView, pathSaveFile As String,
                                   Optional tableName As String = Nothing, Optional showException As Boolean = True) As Boolean
                Try

                    Dim DT As New DataTable
                    For Each col As ColumnHeader In lv.Columns
                        DT.Columns.Add(col.Text)
                    Next
                    'set progressbar
                    If Progressbar IsNot Nothing Then
                        Progressbar.Value = 0
                        Progressbar.Maximum = lv.Items.Count
                    End If
                    'menyimpan data
                    For Each item As ListViewItem In lv.Items
                        Dim row As DataRow = DT.NewRow()
                        For i As Integer = 0 To item.SubItems.Count - 1
                            row(i) = item.SubItems(i).Text
                        Next
                        DT.Rows.Add(row)
                        'set progressbar
                        If Progressbar IsNot Nothing Then
                            Application.DoEvents()
                            Progressbar.Value += 1
                        End If
                    Next

                    Dim dataJSON As String = Nothing
                    If tableName = Nothing Then
                        dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                    Else
                        DT.TableName = tableName
                        Dim ds As New DataSet
                        ds.Tables.Add(DT)
                        dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                    End If
                    If dataJSON <> Nothing Then
                        Dim varTargetFile As IO.StreamWriter
                        varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                        varTargetFile.Write(dataJSON)
                        varTargetFile.Close()
                        Return True
                    Else
                        Return False
                    End If
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.FromListView")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

#End Region

        End Class

        ''' <summary>
        ''' Class export dalam bentuk data terenkripsi
        ''' </summary>
        Public Class Encrypted

            ''' <summary>
            ''' Class enkripsi DES
            ''' </summary>
            Public Class DES

                ''' <summary>
                ''' Class export dari objek dataset dengan data terenkripsi
                ''' </summary>
                Public Class FromDataSet
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

#Region "DataSet"
                    ''' <summary>
                    ''' Export Dataset ke file Teks
                    ''' </summary>
                    ''' <param name="ds">Nama object dataset</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="table">Index table Dataset</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Default Delimiter menggunakan karakter koma (,)</remarks>
                    Public Function ToText(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                           Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True,
                                           Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim Delimiter As String = ","
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Delimiter
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        If encapsulation = True Then
                                            varText += """" + des.Encode(column.ToString, secretKey) + """" + varSeparator
                                        Else
                                            varText += des.Encode(column.ToString, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            If encapsulation = True Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        Else
                                            'Formatted datetime
                                            If encapsulation = True Then
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                                End If
                                            Else
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                                End If
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke Excel
                    ''' </summary>
                    ''' <param name="ds">Nama object DataSet</param>
                    ''' <param name="pathSaveFile">Lokasi path file excel</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional table As Integer = 0, Optional ByVal usePassword As String = Nothing,
                                            Optional ByVal passwordWorkBook As String = Nothing, Optional formatDateTime As String = Nothing,
                                            Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To DT.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, i + 1) = des.Encode(DT.Columns(i).ColumnName, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menambah data
                            For h As Integer = 0 To DT.Rows.Count - 1
                                For j As Integer = 0 To DT.Columns.Count - 1
                                    Application.DoEvents()
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, des.Encode("", secretKey), des.Encode(DT.Rows(h).Item(j).ToString, secretKey))
                                    Else
                                        'Formatted datetime
                                        If TypeOf (DT.Rows(h).Item(j)) Is DateTime Then
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, des.Encode("", secretKey), des.Encode(DirectCast(DT.Rows(h).Item(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                        Else
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, des.Encode("", secretKey), des.Encode(DT.Rows(h).Item(j).ToString, secretKey))
                                        End If
                                    End If
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke CSV
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        varText += """" + des.Encode(column.ToString, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            End If
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke TSV
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        varText += des.Encode(column.ToString, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke HTML
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                           Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataColumn In DT.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + des.Encode(column.ToString, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey), des.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey), des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                            Else
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey), des.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                            End If
                                        End If
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke XML
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="Table">Index table DataSet</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(ds As DataSet, pathSaveFile As String, secretKey As String, tableName As String, Optional writeSchema As Boolean = True,
                                          Optional table As Integer = 0, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = des.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = des.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = des.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If

                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DTEncrypt.TableName = tableName
                            If writeSchema = True Then
                                DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DTEncrypt.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke JSON
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="Table">Index table DataSet. Default = 0</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(ds As DataSet, pathSaveFile As String, secretKey As String, tableName As String,
                                           Optional table As Integer = 0, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = des.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = des.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = des.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DTEncrypt.TableName = tableName
                                Dim ds1 As New DataSet
                                ds1.Tables.Add(DTEncrypt)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds1, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek datatable dengan data terenkripsi
                ''' </summary>
                Public Class FromDataTable
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

#Region "Datatable"
                    ''' <summary>
                    ''' Export Datatable ke file Teks
                    ''' </summary>
                    ''' <param name="dt">Nama object datatable</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Default Delimiter menggunakan karakter koma (,)</remarks>
                    Public Function ToText(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String, Optional useHeader As Boolean = True,
                                           Optional encapsulation As Boolean = True, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            Dim Delimiter As String = ","
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Delimiter
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        If encapsulation = True Then
                                            varText += """" + des.Encode(column.ToString, secretKey) + """" + varSeparator
                                        Else
                                            varText += des.Encode(column.ToString, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            If encapsulation = True Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        Else
                                            'Formatted datetime
                                            If encapsulation = True Then
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                                End If
                                            Else
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                                End If
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke Excel
                    ''' </summary>
                    ''' <param name="dt">Nama object datatable</param>
                    ''' <param name="pathSaveFile">Lokasi path file excel</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                            Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To dt.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, i + 1) = des.Encode(dt.Columns(i).ColumnName, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'Menambah data
                            For h As Integer = 0 To dt.Rows.Count - 1
                                For j As Integer = 0 To dt.Columns.Count - 1
                                    Application.DoEvents()
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, des.Encode("", secretKey), des.Encode(dt.Rows(h).Item(j).ToString, secretKey))
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dt.Rows(h).Item(j)) Is DateTime Then
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, des.Encode("", secretKey), des.Encode(DirectCast(dt.Rows(h).Item(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                        Else
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, des.Encode("", secretKey), des.Encode(dt.Rows(h).Item(j).ToString, secretKey))
                                        End If
                                    End If

                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke CSV
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        varText += """" + des.Encode(column.ToString, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke TSV
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        varText += des.Encode(column.ToString, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke HTML
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataColumn In dt.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + des.Encode(column.ToString, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey), des.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey), des.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                            Else
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, des.Encode("", secretKey), des.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                            End If
                                        End If
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke XML
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(dt As DataTable, pathSaveFile As String, secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In dt.Columns
                                DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'Menyimpan data
                            For Each dr As DataRow In dt.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = des.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = des.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = des.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DTEncrypt.TableName = tableName
                            If writeSchema = True Then
                                DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DTEncrypt.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke JSON
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(dt As DataTable, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In dt.Columns
                                DTEncrypt.Columns.Add(des.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'Menyimpan data
                            For Each dr As DataRow In dt.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = des.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = des.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = des.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DTEncrypt.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DTEncrypt)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If

                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek datagridview dengan data terenkripsi
                ''' </summary>
                Public Class FromDataGridView
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

#Region "DataGridView"

                    ''' <summary>
                    ''' Export DataGridView ke file Teks
                    ''' </summary>
                    ''' <param name="dgv">Nama object DataGridView</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Delimiter default = ","</remarks>
                    Public Function ToText(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String, Optional useHeader As Boolean = True,
                                           Optional encapsulation As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = ","
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        If encapsulation = True Then
                                            varText += """" + des.Encode(column.HeaderText, secretKey) + """" + varSeparator
                                        Else
                                            varText += des.Encode(column.HeaderText, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        If encapsulation = True Then
                                            varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Cells(i).Value.ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Cells(i).Value.ToString, secretKey) + varSeparator).ToString
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke Excel
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Path lokasi hasil export excel</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                            Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.RowCount = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To dgv.ColumnCount - 1
                                varExcelWorkSheet.Cells(1, i + 1) = des.Encode(dgv.Columns(i).HeaderText, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'Menambah data
                            For i As Integer = 0 To dgv.RowCount - 1
                                For j As Integer = 0 To dgv.ColumnCount - 1
                                    varExcelWorkSheet.Cells(i + 2, j + 1) = IIf(IsDBNull(dgv.Rows(i).Cells(j).Value.ToString()) = True, des.Encode("", secretKey), des.Encode(dgv.Rows(i).Cells(j).Value.ToString(), secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke CSV
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        varText += """" + des.Encode(column.HeaderText, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.Cells(i).Value.ToString, secretKey) + """" + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke TSV
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'Tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        varText += des.Encode(column.HeaderText, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.Cells(i).Value.ToString, secretKey) + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke HTML
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.RowCount = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataGridViewColumn In dgv.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + des.Encode(column.HeaderText, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To dgv.ColumnCount - 1
                                        varText += "<TD>" + IIf(IsDBNull(row.Cells(i).Value) = True, des.Encode("", secretKey), des.Encode(row.Cells(i).Value.ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke XML
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(dgv As DataGridView, pathSaveFile As String, ByVal secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable()
                            For Each col As DataGridViewColumn In dgv.Columns
                                DT.Columns.Add(des.Encode(col.HeaderText, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'Menyimpan data
                            For Each row As DataGridViewRow In dgv.Rows
                                Dim dRow As DataRow = DT.NewRow()
                                For Each cell As DataGridViewCell In row.Cells
                                    dRow(cell.ColumnIndex) = des.Encode(cell.Value.ToString, secretKey)
                                Next
                                DT.Rows.Add(dRow)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DT.TableName = tableName
                            If writeSchema = True Then
                                DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DT.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataGridView")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke JSON
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(dgv As DataGridView, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable()
                            For Each col As DataGridViewColumn In dgv.Columns
                                DT.Columns.Add(des.Encode(col.HeaderText, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'Menyimpan data
                            For Each row As DataGridViewRow In dgv.Rows
                                Dim dRow As DataRow = DT.NewRow()
                                For Each cell As DataGridViewCell In row.Cells
                                    dRow(cell.ColumnIndex) = des.Encode(cell.Value.ToString, secretKey)
                                Next
                                DT.Rows.Add(dRow)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DT.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DT)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek listview dengan data terenkripsi
                ''' </summary>
                Public Class FromListView
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

#Region "ListView"

                    ''' <summary>
                    ''' Export ListView ke file Teks
                    ''' </summary>
                    ''' <param name="lv">Nama object ListView</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToText(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String, Optional useHeader As Boolean = True,
                                           Optional encapsulation As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = ","
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        If encapsulation = True Then
                                            varText += """" + des.Encode(column.Text, secretKey) + """" + varSeparator
                                        Else
                                            varText += des.Encode(column.Text, secretKey) + varSeparator
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        If encapsulation = True Then
                                            varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.SubItems.Item(i).Text, secretKey) + """" + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.SubItems.Item(i).Text, secretKey) + varSeparator).ToString
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke Excel
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Path lokasi hasil export excel</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="PasswordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String, Optional ByVal usePassword As String = Nothing,
                                            Optional ByVal passwordWorkBook As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For h As Integer = 0 To lv.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, h + 1) = des.Encode(lv.Columns(h).Text, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'Menambah data
                            For i As Integer = 0 To lv.Items.Count - 1
                                For j As Integer = 0 To lv.Columns.Count - 1
                                    varExcelWorkSheet.Cells(i + 2, j + 1) = IIf(IsDBNull(lv.Items.Item(i).SubItems.Item(j).Text) = True, des.Encode("", secretKey), des.Encode(lv.Items.Item(i).SubItems.Item(j).Text, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke CSV
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        varText += """" + des.Encode(column.Text, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, """" + des.Encode("", secretKey) + """" + varSeparator, """" + des.Encode(row.SubItems.Item(i).Text, secretKey) + """" + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke TSV
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        varText = varText + des.Encode(column.Text, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText = varText + IIf(IsDBNull(row.SubItems.Item(i).Text) = True, des.Encode("", secretKey) + varSeparator, des.Encode(row.SubItems.Item(i).Text, secretKey) + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke HTML
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As ColumnHeader In lv.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + des.Encode(column.Text, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText += "<TD>" + IIf(IsDBNull(row.SubItems.Item(i).Text) = True, des.Encode("", secretKey), des.Encode(row.SubItems.Item(i).Text, secretKey)).ToString + "</TD>" + vbCrLf
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke XML
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(lv As ListView, pathSaveFile As String, secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            For Each col As ColumnHeader In lv.Columns
                                DT.Columns.Add(des.Encode(col.Text, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'Menyimpen data
                            For Each item As ListViewItem In lv.Items
                                Dim row As DataRow = DT.NewRow()
                                For i As Integer = 0 To item.SubItems.Count - 1
                                    row(i) = des.Encode(item.SubItems(i).Text, secretKey)
                                Next
                                DT.Rows.Add(row)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value += 1
                                End If
                                Application.DoEvents()
                            Next
                            DT.TableName = tableName
                            If writeSchema = True Then
                                DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DT.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataListView")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke JSON
                    ''' </summary>
                    ''' <param name="lv">nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(lv As ListView, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable
                            For Each col As ColumnHeader In lv.Columns
                                DT.Columns.Add(des.Encode(col.Text, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'Menyimpen data
                            For Each item As ListViewItem In lv.Items
                                Dim row As DataRow = DT.NewRow()
                                For i As Integer = 0 To item.SubItems.Count - 1
                                    row(i) = des.Encode(item.SubItems(i).Text, secretKey)
                                Next
                                DT.Rows.Add(row)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value += 1
                                End If
                                Application.DoEvents()
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DT.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DT)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class
            End Class

            ''' <summary>
            ''' Class enkripsi 3DES
            ''' </summary>
            Public Class TripleDES

                ''' <summary>
                ''' Class export dari objek dataset dengan data terenkripsi
                ''' </summary>
                Public Class FromDataSet
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

#Region "DataSet"
                    ''' <summary>
                    ''' Export Dataset ke file Teks
                    ''' </summary>
                    ''' <param name="ds">Nama object dataset</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="table">Index table Dataset</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Default Delimiter menggunakan karakter koma (,)</remarks>
                    Public Function ToText(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                           Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True,
                                           Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim Delimiter As String = ","
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Delimiter
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        If encapsulation = True Then
                                            varText += """" + tripledes.Encode(column.ToString, secretKey) + """" + varSeparator
                                        Else
                                            varText += tripledes.Encode(column.ToString, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            If encapsulation = True Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        Else
                                            'Formatted datetime
                                            If encapsulation = True Then
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                                End If
                                            Else
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                                End If
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke Excel
                    ''' </summary>
                    ''' <param name="ds">Nama object DataSet</param>
                    ''' <param name="pathSaveFile">Lokasi path file excel</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional table As Integer = 0, Optional ByVal usePassword As String = Nothing,
                                            Optional ByVal passwordWorkBook As String = Nothing, Optional formatDateTime As String = Nothing,
                                            Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To DT.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, i + 1) = tripledes.Encode(DT.Columns(i).ColumnName, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menambah data
                            For h As Integer = 0 To DT.Rows.Count - 1
                                For j As Integer = 0 To DT.Columns.Count - 1
                                    Application.DoEvents()
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DT.Rows(h).Item(j).ToString, secretKey))
                                    Else
                                        'Formatted datetime
                                        If TypeOf (DT.Rows(h).Item(j)) Is DateTime Then
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DirectCast(DT.Rows(h).Item(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                        Else
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DT.Rows(h).Item(j).ToString, secretKey))
                                        End If
                                    End If
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke CSV
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        varText += """" + tripledes.Encode(column.ToString, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            End If
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke TSV
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        varText += tripledes.Encode(column.ToString, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke HTML
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                           Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataColumn In DT.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + tripledes.Encode(column.ToString, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey), tripledes.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey), tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                            Else
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey), tripledes.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                            End If
                                        End If
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke XML
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="Table">Index table DataSet</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(ds As DataSet, pathSaveFile As String, secretKey As String, tableName As String, Optional writeSchema As Boolean = True,
                                          Optional table As Integer = 0, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = tripledes.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = tripledes.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = tripledes.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If

                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DTEncrypt.TableName = tableName
                            If writeSchema = True Then
                                DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DTEncrypt.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke JSON
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="Table">Index table DataSet. Default = 0</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(ds As DataSet, pathSaveFile As String, secretKey As String, tableName As String,
                                           Optional table As Integer = 0, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = tripledes.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = tripledes.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = tripledes.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DTEncrypt.TableName = tableName
                                Dim ds1 As New DataSet
                                ds1.Tables.Add(DTEncrypt)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds1, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek datatable dengan data terenkripsi
                ''' </summary>
                Public Class FromDataTable
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

#Region "Datatable"
                    ''' <summary>
                    ''' Export Datatable ke file Teks
                    ''' </summary>
                    ''' <param name="dt">Nama object datatable</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Default Delimiter menggunakan karakter koma (,)</remarks>
                    Public Function ToText(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String, Optional useHeader As Boolean = True,
                                           Optional encapsulation As Boolean = True, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            Dim Delimiter As String = ","
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Delimiter
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        If encapsulation = True Then
                                            varText += """" + tripledes.Encode(column.ToString, secretKey) + """" + varSeparator
                                        Else
                                            varText += tripledes.Encode(column.ToString, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            If encapsulation = True Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        Else
                                            'Formatted datetime
                                            If encapsulation = True Then
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                                End If
                                            Else
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                                End If
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke Excel
                    ''' </summary>
                    ''' <param name="dt">Nama object datatable</param>
                    ''' <param name="pathSaveFile">Lokasi path file excel</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                            Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To dt.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, i + 1) = tripledes.Encode(dt.Columns(i).ColumnName, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'Menambah data
                            For h As Integer = 0 To dt.Rows.Count - 1
                                For j As Integer = 0 To dt.Columns.Count - 1
                                    Application.DoEvents()
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(dt.Rows(h).Item(j).ToString, secretKey))
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dt.Rows(h).Item(j)) Is DateTime Then
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(DirectCast(dt.Rows(h).Item(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                        Else
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, tripledes.Encode("", secretKey), tripledes.Encode(dt.Rows(h).Item(j).ToString, secretKey))
                                        End If
                                    End If

                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke CSV
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        varText += """" + tripledes.Encode(column.ToString, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke TSV
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        varText += tripledes.Encode(column.ToString, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke HTML
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataColumn In dt.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + tripledes.Encode(column.ToString, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey), tripledes.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey), tripledes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                            Else
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, tripledes.Encode("", secretKey), tripledes.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                            End If
                                        End If
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke XML
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(dt As DataTable, pathSaveFile As String, secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In dt.Columns
                                DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In dt.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = tripledes.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = tripledes.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = tripledes.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DTEncrypt.TableName = tableName
                            If writeSchema = True Then
                                DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DTEncrypt.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke JSON
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(dt As DataTable, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In dt.Columns
                                DTEncrypt.Columns.Add(tripledes.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In dt.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = tripledes.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = tripledes.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = tripledes.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DTEncrypt.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DTEncrypt)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If

                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek datagridview dengan data terenkripsi
                ''' </summary>
                Public Class FromDataGridView
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

#Region "DataGridView"

                    ''' <summary>
                    ''' Export DataGridView ke file Teks
                    ''' </summary>
                    ''' <param name="dgv">Nama object DataGridView</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="Encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Delimiter default = ","</remarks>
                    Public Function ToText(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = ","
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        If encapsulation = True Then
                                            varText += """" + tripledes.Encode(column.HeaderText, secretKey) + """" + varSeparator
                                        Else
                                            varText += tripledes.Encode(column.HeaderText, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        If encapsulation = True Then
                                            varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Cells(i).Value.ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Cells(i).Value.ToString, secretKey) + varSeparator).ToString
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke Excel
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Path lokasi hasil export excel</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                            Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.RowCount = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For h As Integer = 0 To dgv.ColumnCount - 1
                                varExcelWorkSheet.Cells(1, h + 1) = tripledes.Encode(dgv.Columns(h).HeaderText, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'Menambah data
                            For i As Integer = 0 To dgv.RowCount - 1
                                For j As Integer = 0 To dgv.ColumnCount - 1
                                    varExcelWorkSheet.Cells(i + 2, j + 1) = IIf(IsDBNull(dgv.Rows(i).Cells(j).Value.ToString()) = True, tripledes.Encode("", secretKey), tripledes.Encode(dgv.Rows(i).Cells(j).Value.ToString(), secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke CSV
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        varText += """" + tripledes.Encode(column.HeaderText, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.Cells(i).Value.ToString, secretKey) + """" + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke TSV
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'Tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        varText += tripledes.Encode(column.HeaderText, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.Cells(i).Value.ToString, secretKey) + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke HTML
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.RowCount = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataGridViewColumn In dgv.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + tripledes.Encode(column.HeaderText, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To dgv.ColumnCount - 1
                                        varText += "<TD>" + IIf(IsDBNull(row.Cells(i).Value) = True, tripledes.Encode("", secretKey), tripledes.Encode(row.Cells(i).Value.ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke XML
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(dgv As DataGridView, pathSaveFile As String, ByVal secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable()
                            For Each col As DataGridViewColumn In dgv.Columns
                                DT.Columns.Add(tripledes.Encode(col.HeaderText, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'menyimpan data
                            For Each row As DataGridViewRow In dgv.Rows
                                Dim dRow As DataRow = DT.NewRow()
                                For Each cell As DataGridViewCell In row.Cells
                                    dRow(cell.ColumnIndex) = tripledes.Encode(cell.Value.ToString, secretKey)
                                Next
                                DT.Rows.Add(dRow)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DT.TableName = tableName
                            If writeSchema = True Then
                                DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DT.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromDataGridView")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke JSON
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(dgv As DataGridView, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable()
                            For Each col As DataGridViewColumn In dgv.Columns
                                DT.Columns.Add(tripledes.Encode(col.HeaderText, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'menyimpan data
                            For Each row As DataGridViewRow In dgv.Rows
                                Dim dRow As DataRow = DT.NewRow()
                                For Each cell As DataGridViewCell In row.Cells
                                    dRow(cell.ColumnIndex) = tripledes.Encode(cell.Value.ToString, secretKey)
                                Next
                                DT.Rows.Add(dRow)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DT.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DT)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek listview dengan data terenkripsi
                ''' </summary>
                Public Class FromListView
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

#Region "ListView"

                    ''' <summary>
                    ''' Export ListView ke file Teks
                    ''' </summary>
                    ''' <param name="lv">Nama object ListView</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToText(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String, Optional useHeader As Boolean = True,
                                           Optional encapsulation As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = ","
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        If encapsulation = True Then
                                            varText += """" + tripledes.Encode(column.Text, secretKey) + """" + varSeparator
                                        Else
                                            varText += tripledes.Encode(column.Text, secretKey) + varSeparator
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        If encapsulation = True Then
                                            varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.SubItems.Item(i).Text, secretKey) + """" + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.SubItems.Item(i).Text, secretKey) + varSeparator).ToString
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke Excel
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Path lokasi hasil export excel</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String, Optional ByVal usePassword As String = Nothing,
                                            Optional ByVal passwordWorkBook As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For h As Integer = 0 To lv.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, h + 1) = tripledes.Encode(lv.Columns(h).Text, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'Menambah data
                            For i As Integer = 0 To lv.Items.Count - 1
                                For j As Integer = 0 To lv.Columns.Count - 1
                                    varExcelWorkSheet.Cells(i + 2, j + 1) = IIf(IsDBNull(lv.Items.Item(i).SubItems.Item(j).Text) = True, tripledes.Encode("", secretKey), tripledes.Encode(lv.Items.Item(i).SubItems.Item(j).Text, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke CSV
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        varText += """" + tripledes.Encode(column.Text, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, """" + tripledes.Encode("", secretKey) + """" + varSeparator, """" + tripledes.Encode(row.SubItems.Item(i).Text, secretKey) + """" + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke TSV
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        varText = varText + tripledes.Encode(column.Text, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText = varText + IIf(IsDBNull(row.SubItems.Item(i).Text) = True, tripledes.Encode("", secretKey) + varSeparator, tripledes.Encode(row.SubItems.Item(i).Text, secretKey) + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke HTML
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As ColumnHeader In lv.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + tripledes.Encode(column.Text, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText += "<TD>" + IIf(IsDBNull(row.SubItems.Item(i).Text) = True, tripledes.Encode("", secretKey), tripledes.Encode(row.SubItems.Item(i).Text, secretKey)).ToString + "</TD>" + vbCrLf
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke XML
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(lv As ListView, pathSaveFile As String, secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            For Each col As ColumnHeader In lv.Columns
                                DT.Columns.Add(tripledes.Encode(col.Text, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'menyimpan data
                            For Each item As ListViewItem In lv.Items
                                Dim row As DataRow = DT.NewRow()
                                For i As Integer = 0 To item.SubItems.Count - 1
                                    row(i) = tripledes.Encode(item.SubItems(i).Text, secretKey)
                                Next
                                DT.Rows.Add(row)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DT.TableName = tableName
                            If writeSchema = True Then
                                DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DT.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromListView")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke JSON
                    ''' </summary>
                    ''' <param name="lv">nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(lv As ListView, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable
                            For Each col As ColumnHeader In lv.Columns
                                DT.Columns.Add(tripledes.Encode(col.Text, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'menyimpan data
                            For Each item As ListViewItem In lv.Items
                                Dim row As DataRow = DT.NewRow()
                                For i As Integer = 0 To item.SubItems.Count - 1
                                    row(i) = tripledes.Encode(item.SubItems(i).Text, secretKey)
                                Next
                                DT.Rows.Add(row)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DT.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DT)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.TripleDES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class
            End Class

            ''' <summary>
            ''' Class enkripsi AES
            ''' </summary>
            Public Class AES

                ''' <summary>
                ''' Class export dari objek dataset dengan data terenkripsi
                ''' </summary>
                Public Class FromDataSet
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

#Region "DataSet"
                    ''' <summary>
                    ''' Export Dataset ke file Teks
                    ''' </summary>
                    ''' <param name="ds">Nama object dataset</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="table">Index table Dataset</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Default Delimiter menggunakan karakter koma (,)</remarks>
                    Public Function ToText(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                           Optional useHeader As Boolean = True, Optional encapsulation As Boolean = True,
                                           Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim Delimiter As String = ","
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Delimiter
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        If encapsulation = True Then
                                            varText += """" + aes.Encode(column.ToString, secretKey) + """" + varSeparator
                                        Else
                                            varText += aes.Encode(column.ToString, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            If encapsulation = True Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        Else
                                            'Formatted datetime
                                            If encapsulation = True Then
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                                End If
                                            Else
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                                End If
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke Excel
                    ''' </summary>
                    ''' <param name="ds">Nama object DataSet</param>
                    ''' <param name="pathSaveFile">Lokasi path file excel</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional table As Integer = 0, Optional ByVal usePassword As String = Nothing,
                                            Optional ByVal passwordWorkBook As String = Nothing, Optional formatDateTime As String = Nothing,
                                            Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To DT.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, i + 1) = aes.Encode(DT.Columns(i).ColumnName, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'Menambah data
                            For h As Integer = 0 To DT.Rows.Count - 1
                                For j As Integer = 0 To DT.Columns.Count - 1
                                    Application.DoEvents()
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, aes.Encode("", secretKey), aes.Encode(DT.Rows(h).Item(j).ToString, secretKey))
                                    Else
                                        'Formatted datetime
                                        If TypeOf (DT.Rows(h).Item(j)) Is DateTime Then
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, aes.Encode("", secretKey), aes.Encode(DirectCast(DT.Rows(h).Item(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                        Else
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(DT.Rows(h).Item(j)) = True, aes.Encode("", secretKey), aes.Encode(DT.Rows(h).Item(j).ToString, secretKey))
                                        End If
                                    End If
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke CSV
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        varText += """" + aes.Encode(column.ToString, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            End If
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke TSV
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In DT.Columns
                                        varText += aes.Encode(column.ToString, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If DT.Rows.IndexOf(row) <> DT.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke HTML
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="table">Index table DataSet</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal ds As DataSet, ByVal pathSaveFile As String, ByVal secretKey As String, Optional table As Integer = 0,
                                           Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            If DT.Rows.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataColumn In DT.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + aes.Encode(column.ToString, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = DT.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In DT.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To DT.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey), aes.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey), aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                            Else
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey), aes.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                            End If
                                        End If
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke XML
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="Table">Index table DataSet</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(ds As DataSet, pathSaveFile As String, secretKey As String, tableName As String, Optional writeSchema As Boolean = True,
                                          Optional table As Integer = 0, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = aes.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = aes.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = aes.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If

                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DTEncrypt.TableName = tableName
                            If writeSchema = True Then
                                DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DTEncrypt.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataSet ke JSON
                    ''' </summary>
                    ''' <param name="ds">Nama objek DataSet</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="Table">Index table DataSet. Default = 0</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(ds As DataSet, pathSaveFile As String, secretKey As String, tableName As String,
                                           Optional table As Integer = 0, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable
                            DT = ds.Tables(table)
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In DT.Columns
                                DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = DT.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In DT.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To DT.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = aes.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = aes.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = aes.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DTEncrypt.TableName = tableName
                                Dim ds1 As New DataSet
                                ds1.Tables.Add(DTEncrypt)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds1, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataSet")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek datatable dengan data terenkripsi
                ''' </summary>
                Public Class FromDataTable
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

#Region "Datatable"
                    ''' <summary>
                    ''' Export Datatable ke file Teks
                    ''' </summary>
                    ''' <param name="dt">Nama object datatable</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Default Delimiter menggunakan karakter koma (,)</remarks>
                    Public Function ToText(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String, Optional useHeader As Boolean = True,
                                           Optional encapsulation As Boolean = True, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            Dim Delimiter As String = ","
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Delimiter
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        If encapsulation = True Then
                                            varText += """" + aes.Encode(column.ToString, secretKey) + """" + varSeparator
                                        Else
                                            varText += aes.Encode(column.ToString, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            If encapsulation = True Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        Else
                                            'Formatted datetime
                                            If encapsulation = True Then
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                                End If
                                            Else
                                                If TypeOf (row.Item(i)) Is DateTime Then
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                                Else
                                                    varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                                End If
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke Excel
                    ''' </summary>
                    ''' <param name="dt">Nama object datatable</param>
                    ''' <param name="pathSaveFile">Lokasi path file excel</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                            Optional formatDateTime As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To dt.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, i + 1) = aes.Encode(dt.Columns(i).ColumnName, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'Menambah data
                            For h As Integer = 0 To dt.Rows.Count - 1
                                For j As Integer = 0 To dt.Columns.Count - 1
                                    Application.DoEvents()
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, aes.Encode("", secretKey), aes.Encode(dt.Rows(h).Item(j).ToString, secretKey))
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dt.Rows(h).Item(j)) Is DateTime Then
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, aes.Encode("", secretKey), aes.Encode(DirectCast(dt.Rows(h).Item(j), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey))
                                        Else
                                            varExcelWorkSheet.Cells(h + 2, j + 1) = IIf(IsDBNull(dt.Rows(h).Item(j)) = True, aes.Encode("", secretKey), aes.Encode(dt.Rows(h).Item(j).ToString, secretKey))
                                        End If
                                    End If

                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke CSV
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        varText += """" + aes.Encode(column.ToString, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + """" + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Item(i).ToString, secretKey) + """" + varSeparator).ToString
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke TSV
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As DataColumn In dt.Columns
                                        varText += aes.Encode(column.ToString, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey) + varSeparator).ToString
                                            Else
                                                varText += IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Item(i).ToString, secretKey) + varSeparator).ToString
                                            End If
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dt.Rows.IndexOf(row) <> dt.Rows.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke HTML
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal dt As DataTable, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            If dt.Rows.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataColumn In dt.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + aes.Encode(column.ToString, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dt.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataRow In dt.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To dt.Columns.Count - 1
                                        If formatDateTime = Nothing Then
                                            'Standar datetime
                                            varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey), aes.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                        Else
                                            'Formatted datetime
                                            If TypeOf (row.Item(i)) Is DateTime Then
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey), aes.Encode(DirectCast(row.Item(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)).ToString + "</TD>" + vbCrLf
                                            Else
                                                varText += "<TD>" + IIf(IsDBNull(row.Item(i).ToString) = True, aes.Encode("", secretKey), aes.Encode(row.Item(i).ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                            End If
                                        End If
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke XML
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(dt As DataTable, pathSaveFile As String, secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional formatDateTime As String = Nothing,
                                          Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In dt.Columns
                                DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In dt.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = aes.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = aes.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = aes.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DTEncrypt.TableName = tableName
                            If writeSchema = True Then
                                DTEncrypt.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DTEncrypt.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataTable ke JSON
                    ''' </summary>
                    ''' <param name="dt">Nama objek DataTable</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="formatDateTime">Otomatis menentukan format khusus DateTime. Default = Nothing. [Contoh: "yyyy-MM-dd HH:mm:ss"]</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(dt As DataTable, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional formatDateTime As String = Nothing,
                                           Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DTEncrypt As New DataTable
                            For Each col As DataColumn In dt.Columns
                                DTEncrypt.Columns.Add(aes.Encode(col.ToString, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dt.Rows.Count
                            End If
                            'menyimpan data
                            For Each dr As DataRow In dt.Rows
                                Dim drNew As DataRow = DTEncrypt.NewRow()
                                For i As Integer = 0 To dt.Columns.Count - 1
                                    If formatDateTime = Nothing Then
                                        'Standar datetime
                                        drNew(i) = aes.Encode(dr(i).ToString, secretKey)
                                    Else
                                        'Formatted datetime
                                        If TypeOf (dr(i)) Is DateTime Then
                                            drNew(i) = aes.Encode(DirectCast(dr(i), DateTime).ToString(formatDateTime, System.Globalization.CultureInfo.InvariantCulture), secretKey)
                                        Else
                                            drNew(i) = aes.Encode(dr(i).ToString, secretKey)
                                        End If
                                    End If
                                Next
                                DTEncrypt.Rows.Add(drNew)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DTEncrypt, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DTEncrypt.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DTEncrypt)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If

                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.DES.FromDataTable")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function
#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek datagridview dengan data terenkripsi
                ''' </summary>
                Public Class FromDataGridView
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

#Region "DataGridView"

                    ''' <summary>
                    ''' Export DataGridView ke file Teks
                    ''' </summary>
                    ''' <param name="dgv">Nama object DataGridView</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="Encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    ''' <remarks>Delimiter default = ","</remarks>
                    Public Function ToText(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String, Optional useHeader As Boolean = True,
                                           Optional Encapsulation As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = ","
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        If Encapsulation = True Then
                                            varText += """" + aes.Encode(column.HeaderText, secretKey) + """" + varSeparator
                                        Else
                                            varText += aes.Encode(column.HeaderText, secretKey) + varSeparator
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        If Encapsulation = True Then
                                            varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Cells(i).Value.ToString, secretKey) + """" + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Cells(i).Value.ToString, secretKey) + varSeparator).ToString
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke Excel
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Path lokasi hasil export excel</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                            Optional ByVal usePassword As String = Nothing, Optional ByVal passwordWorkBook As String = Nothing,
                                            Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.RowCount = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For h As Integer = 0 To dgv.ColumnCount - 1
                                varExcelWorkSheet.Cells(1, h + 1) = aes.Encode(dgv.Columns(h).HeaderText, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'Menambah data
                            For i As Integer = 0 To dgv.RowCount - 1
                                For j As Integer = 0 To dgv.ColumnCount - 1
                                    varExcelWorkSheet.Cells(i + 2, j + 1) = IIf(IsDBNull(dgv.Rows(i).Cells(j).Value.ToString()) = True, aes.Encode("", secretKey), aes.Encode(dgv.Rows(i).Cells(j).Value.ToString(), secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke CSV
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        varText += """" + aes.Encode(column.HeaderText, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.Cells(i).Value.ToString, secretKey) + """" + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke TSV
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.Rows.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'Tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                If useHeader = True Then
                                    varText = Nothing
                                    For Each column As DataGridViewColumn In dgv.Columns
                                        varText += aes.Encode(column.HeaderText, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = Nothing
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.Cells(i).Value.ToString) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.Cells(i).Value.ToString, secretKey) + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If dgv.Rows.IndexOf(row) <> dgv.RowCount - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke HTML
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal dgv As DataGridView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional showException As Boolean = True) As Boolean
                        Try
                            If dgv.RowCount = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As DataGridViewColumn In dgv.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + aes.Encode(column.HeaderText, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = dgv.Rows.Count
                                End If
                                'Menyimpan data
                                For Each row As DataGridViewRow In dgv.Rows
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To dgv.ColumnCount - 1
                                        varText += "<TD>" + IIf(IsDBNull(row.Cells(i).Value) = True, aes.Encode("", secretKey), aes.Encode(row.Cells(i).Value.ToString, secretKey)).ToString + "</TD>" + vbCrLf
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke XML
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(dgv As DataGridView, pathSaveFile As String, ByVal secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable()
                            For Each col As DataGridViewColumn In dgv.Columns
                                DT.Columns.Add(aes.Encode(col.HeaderText, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'menyimpan data
                            For Each row As DataGridViewRow In dgv.Rows
                                Dim dRow As DataRow = DT.NewRow()
                                For Each cell As DataGridViewCell In row.Cells
                                    dRow(cell.ColumnIndex) = aes.Encode(cell.Value.ToString, secretKey)
                                Next
                                DT.Rows.Add(dRow)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DT.TableName = tableName
                            If writeSchema = True Then
                                DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DT.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromDataGridView")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export DataGridView ke JSON
                    ''' </summary>
                    ''' <param name="dgv">Nama objek DataGridView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(dgv As DataGridView, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable()
                            For Each col As DataGridViewColumn In dgv.Columns
                                DT.Columns.Add(aes.Encode(col.HeaderText, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = dgv.Rows.Count
                            End If
                            'menyimpan data
                            For Each row As DataGridViewRow In dgv.Rows
                                Dim dRow As DataRow = DT.NewRow()
                                For Each cell As DataGridViewCell In row.Cells
                                    dRow(cell.ColumnIndex) = aes.Encode(cell.Value.ToString, secretKey)
                                Next
                                DT.Rows.Add(dRow)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DT.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DT)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromDataGridView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

#End Region

                End Class

                ''' <summary>
                ''' Class export dari objek listview dengan data terenkripsi
                ''' </summary>
                Public Class FromListView
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

#Region "ListView"

                    ''' <summary>
                    ''' Export ListView ke file Teks
                    ''' </summary>
                    ''' <param name="lv">Nama object ListView</param>
                    ''' <param name="pathSaveFile">Lokasi path file txt</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="encapsulation">Seluruh field akan dibungkus dengan Double Quotes. Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToText(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String, Optional useHeader As Boolean = True,
                                           Optional encapsulation As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = ","
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        If encapsulation = True Then
                                            varText += """" + aes.Encode(column.Text, secretKey) + """" + varSeparator
                                        Else
                                            varText += aes.Encode(column.Text, secretKey) + varSeparator
                                        End If

                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        If encapsulation = True Then
                                            varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.SubItems.Item(i).Text, secretKey) + """" + varSeparator).ToString
                                        Else
                                            varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.SubItems.Item(i).Text, secretKey) + varSeparator).ToString
                                        End If
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke Excel
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Path lokasi hasil export excel</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="usePassword">Gunakan Password? Default = tidak ada password</param>
                    ''' <param name="passwordWorkBook">Gunakan Password di WorkBook? Default = tidak ada password</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToExcel(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String, Optional ByVal usePassword As String = Nothing,
                                            Optional ByVal passwordWorkBook As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varExcelApp As Excels.Application
                            Dim varExcelWorkBook As Excels.Workbook
                            Dim varExcelWorkSheet As Excels.Worksheet
                            Dim misValue As Object = System.Reflection.Missing.Value

                            varExcelApp = New Excels.Application
                            varExcelWorkBook = varExcelApp.Workbooks.Add(misValue)
                            varExcelWorkSheet = CType(varExcelWorkBook.Sheets("sheet1"), Excels.Worksheet)
                            'Menambah header
                            For i As Integer = 0 To lv.Columns.Count - 1
                                varExcelWorkSheet.Cells(1, i + 1) = aes.Encode(lv.Columns(i).Text, secretKey)
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'Menambah data
                            For i As Integer = 0 To lv.Items.Count - 1
                                For j As Integer = 0 To lv.Columns.Count - 1
                                    varExcelWorkSheet.Cells(i + 2, j + 1) = IIf(IsDBNull(lv.Items.Item(i).SubItems.Item(j).Text) = True, aes.Encode("", secretKey), aes.Encode(lv.Items.Item(i).SubItems.Item(j).Text, secretKey))
                                Next
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            If usePassword <> Nothing Then varExcelWorkSheet.Protect(Password:=usePassword)
                            If passwordWorkBook <> Nothing Then varExcelWorkBook.Password = passwordWorkBook
                            varExcelWorkSheet.SaveAs(pathSaveFile)
                            varExcelWorkBook.Close()
                            varExcelApp.Quit()

                            releaseObject(varExcelApp)
                            releaseObject(varExcelWorkBook)
                            releaseObject(varExcelWorkSheet)
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke CSV
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export CSV</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToCSV(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = "," 'comma
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        varText += """" + aes.Encode(column.Text, secretKey) + """" + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText += IIf(IsDBNull(row.SubItems.Item(i).Text) = True, """" + aes.Encode("", secretKey) + """" + varSeparator, """" + aes.Encode(row.SubItems.Item(i).Text, secretKey) + """" + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke TSV
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export TSV</param>
                    ''' <param name="useHeader">Gunakan Header? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToTSV(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                          Optional useHeader As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varSeparator As String = Chr(9) 'tab
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                If useHeader = True Then
                                    'Menyimpan header
                                    varText = Nothing
                                    For Each column As ColumnHeader In lv.Columns
                                        varText = varText + aes.Encode(column.Text, secretKey) + varSeparator
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    .WriteLine()
                                End If
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = Nothing
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText = varText + IIf(IsDBNull(row.SubItems.Item(i).Text) = True, aes.Encode("", secretKey) + varSeparator, aes.Encode(row.SubItems.Item(i).Text, secretKey) + varSeparator).ToString
                                    Next
                                    varText = Mid(varText, 1, varText.Length - 1)
                                    .Write(varText)
                                    If lv.Items.IndexOf(row) <> lv.Items.Count - 1 Then
                                        .WriteLine()
                                    End If
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke HTML
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export HTML</param>
                    ''' <param name="bgColor">Warna background kolom header. Default = #CCCCCC</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToHTML(ByVal lv As ListView, ByVal pathSaveFile As String, ByVal secretKey As String,
                                           Optional bgColor As String = "#CCCCCC", Optional showException As Boolean = True) As Boolean
                        Try
                            If lv.Items.Count = 0 Then Return False
                            Dim varText As String
                            Dim varTargetFile As IO.StreamWriter
                            varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                            With varTargetFile
                                'Menyimpan header
                                varText = "<TABLE BORDER='1' cellspacing='0'>" + vbCrLf + "<TR>" + vbCrLf
                                For Each column As ColumnHeader In lv.Columns
                                    varText += "<TH bgcolor='" + bgColor + "' align='center'>" + aes.Encode(column.Text, secretKey) + "</TH>" + vbCrLf
                                Next
                                varText += "</TR>" + vbCrLf
                                .Write(varText)
                                .WriteLine()
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Progressbar.Value = 0
                                    Progressbar.Maximum = lv.Items.Count
                                End If
                                'Menyimpan data
                                For Each row As ListViewItem In lv.Items
                                    varText = "<TR>" + vbCrLf
                                    For i As Integer = 0 To lv.Columns.Count - 1
                                        varText += "<TD>" + IIf(IsDBNull(row.SubItems.Item(i).Text) = True, aes.Encode("", secretKey), aes.Encode(row.SubItems.Item(i).Text, secretKey)).ToString + "</TD>" + vbCrLf
                                    Next
                                    varText += "</TR>" + vbCrLf
                                    .Write(varText)
                                    .WriteLine()
                                    'set progressbar
                                    If Progressbar IsNot Nothing Then
                                        Application.DoEvents()
                                        Progressbar.Value += 1
                                    End If
                                Next
                                varText = "</TABLE>"
                                .Write(varText)
                                .Close()
                            End With
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke XML
                    ''' </summary>
                    ''' <param name="lv">Nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export XML</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Nama table data</param>
                    ''' <param name="writeSchema">Gunakan schema? Default = True</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToXML(lv As ListView, pathSaveFile As String, secretKey As String, tableName As String,
                                          Optional writeSchema As Boolean = True, Optional showException As Boolean = True) As Boolean
                        Try
                            Dim DT As New DataTable
                            For Each col As ColumnHeader In lv.Columns
                                DT.Columns.Add(aes.Encode(col.Text, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'menyimpan data
                            For Each item As ListViewItem In lv.Items
                                Dim row As DataRow = DT.NewRow()
                                For i As Integer = 0 To item.SubItems.Count - 1
                                    row(i) = aes.Encode(item.SubItems(i).Text, secretKey)
                                Next
                                DT.Rows.Add(row)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next
                            DT.TableName = tableName
                            If writeSchema = True Then
                                DT.WriteXml(pathSaveFile, XmlWriteMode.WriteSchema)
                            Else
                                DT.WriteXml(pathSaveFile)
                            End If
                            Return True
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromListView")
                            Return False
                        End Try
                    End Function

                    ''' <summary>
                    ''' Export ListView ke JSON
                    ''' </summary>
                    ''' <param name="lv">nama objek ListView</param>
                    ''' <param name="pathSaveFile">Full path lokasi export JSON</param>
                    ''' <param name="secretKey">Secret kode untuk enkripsi</param>
                    ''' <param name="tableName">Menentukan nama table data. Default = Nothing</param>
                    ''' <param name="showException">Tampilkan log exception? Default = True</param>
                    ''' <returns>Boolean</returns>
                    Public Function ToJSON(lv As ListView, pathSaveFile As String, secretKey As String,
                                           Optional tableName As String = Nothing, Optional showException As Boolean = True) As Boolean
                        Try
                            'proses enkripsi
                            Dim DT As New DataTable
                            For Each col As ColumnHeader In lv.Columns
                                DT.Columns.Add(aes.Encode(col.Text, secretKey))
                            Next
                            'set progressbar
                            If Progressbar IsNot Nothing Then
                                Progressbar.Value = 0
                                Progressbar.Maximum = lv.Items.Count
                            End If
                            'menyimpan data
                            For Each item As ListViewItem In lv.Items
                                Dim row As DataRow = DT.NewRow()
                                For i As Integer = 0 To item.SubItems.Count - 1
                                    row(i) = aes.Encode(item.SubItems(i).Text, secretKey)
                                Next
                                DT.Rows.Add(row)
                                'set progressbar
                                If Progressbar IsNot Nothing Then
                                    Application.DoEvents()
                                    Progressbar.Value += 1
                                End If
                            Next

                            'proses export
                            Dim dataJSON As String = Nothing
                            If tableName = Nothing Then
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(DT, Newtonsoft.Json.Formatting.Indented)
                            Else
                                DT.TableName = tableName
                                Dim ds As New DataSet
                                ds.Tables.Add(DT)
                                dataJSON = Newtonsoft.Json.JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented)
                            End If
                            If dataJSON <> Nothing Then
                                Dim varTargetFile As IO.StreamWriter
                                varTargetFile = New IO.StreamWriter(pathSaveFile, False)
                                varTargetFile.Write(dataJSON)
                                varTargetFile.Close()
                                Return True
                            Else
                                Return False
                            End If
                        Catch ex As Exception
                            If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.data.Export.Encrypted.AES.FromListView")
                            Return False
                        Finally
                            GC.Collect()
                        End Try
                    End Function

#End Region

                End Class
            End Class
        End Class
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