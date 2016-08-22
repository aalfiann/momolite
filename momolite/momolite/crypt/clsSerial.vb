
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
Namespace crypt
    ''' <summary>Class Generate Serial Number</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>19 April 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>1.2.0</version>
    ''' <requirement>
    ''' - Imports System.Windows.Forms
    ''' </requirement>
    Public Class Serial

        ''' <summary>
        ''' Membuat GUID baru
        ''' </summary>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        ''' <remarks>Nilai yang dihasilkan merupakan unique dan random dari Net Framework</remarks>
        Public Function GenerateGUID(Optional showException As Boolean = True) As String
            Try
                Return Guid.NewGuid().ToString
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.Serial")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Membuat Serial baru secara acak
        ''' </summary>
        ''' <param name="singleFormat">Single Format untuk indikasi generate format GUID. Default = "N". Pilihan Format: "N", "D", "B", "P", atau "X". Jika NOTHING/NULL maka Default = "D"</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        ''' <remarks>Nilai yang dihasilkan merupakan unique dan random</remarks>
        Public Function GenerateSerial(Optional singleFormat As String = "N", Optional showException As Boolean = True) As String
            Try
                Dim serialGuid As System.Guid = System.Guid.NewGuid()
                Dim uniqueSerial As String = serialGuid.ToString(singleFormat)
                Dim uniqueSerialLength As String = uniqueSerial.Substring(0, 28).ToUpper()

                Dim serialArray As Char() = uniqueSerialLength.ToCharArray()
                Dim finalSerialNumber As String = Nothing

                Dim j As Integer = 0
                For i As Integer = 0 To 27
                    For j = i To 4 + (i - 1)
                        finalSerialNumber += serialArray(j)
                    Next
                    If j = 28 Then
                        Exit For
                    Else
                        i = (j) - 1
                        finalSerialNumber += "-"
                    End If
                Next

                Return finalSerialNumber
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.Serial")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function
    End Class
End Namespace