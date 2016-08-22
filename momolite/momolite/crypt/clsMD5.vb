
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

Imports System.Text
Imports System.Windows.Forms
Imports Crypto = System.Security.Cryptography

Namespace crypt
    ''' <summary>Class MD5</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>19 April 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>2.1.0</version>
    ''' <requirement>
    ''' - Imports System.Text
    ''' - Imports System.Windows.Forms
    ''' - Imports Crypto = System.Security.Cryptography
    ''' </requirement>
    Public Class MD5

        ''' <summary>
        ''' Standar MD5
        ''' </summary>
        Public Class MD5

            Private ReadOnly _md5 As Crypto.MD5 = Crypto.MD5.Create()

            ''' <summary>
            ''' Membuat hash MD5
            ''' </summary>
            ''' <param name="input">Karakter string yang akan di hash MD5</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>String()</returns>
            Public Function Generate(input As String, secretKey As String, Optional showException As Boolean = True) As String
                Try
                    Dim data = _md5.ComputeHash(Encoding.UTF8.GetBytes(secretKey + input + secretKey))
                    Dim sb As New StringBuilder()
                    Array.ForEach(data, Function(x) sb.Append(x.ToString("X2")))
                    Return sb.ToString()
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.MD5.MD5")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Memverifikasi enkripsi hash MD5 dengan string sebelum di enkripsi hash MD5
            ''' </summary>
            ''' <param name="input">Karakter sebelum enkripsi MD5</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="hash">Karakter setelah enkripsi MD5</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function Validate(input As String, secretKey As String, hash As String, Optional showException As Boolean = True) As Boolean
                Try
                    Dim sourceHash = Generate(input, secretKey)
                    Dim comparer As StringComparer = StringComparer.OrdinalIgnoreCase
                    Return If(comparer.Compare(sourceHash, hash) = 0, True, False)
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.MD5.MD5")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

        End Class

        ''' <summary>
        ''' MD5 mix dengan base 64
        ''' </summary>
        Public Class MD5Base64

            Private ReadOnly _md5 As Crypto.MD5 = Crypto.MD5.Create()

            ''' <summary>
            ''' Generate hash MD5 base64
            ''' </summary>
            ''' <param name="input">Karakter string yang akan di hash MD5 base64</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>String</returns>
            Public Function Generate(input As String, secretKey As String, Optional showException As Boolean = True) As String
                Try
                    Dim data = _md5.ComputeHash(Encoding.UTF8.GetBytes(secretKey + input + secretKey))
                    Return Convert.ToBase64String(data)
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.MD5.MD5Base64")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Memvalidasi enkripsi hash MD5 base64 dengan string sebelum di enkripsi hash MD5 base64
            ''' </summary>
            ''' <param name="input">Karakter sebelum enkripsi MD5 base64</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="hash">Karakter setelah enkripsi MD5 base64</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function Validate(input As String, secretKey As String, hash As String, Optional showException As Boolean = True) As Boolean
                Try
                    Dim sourceHash = Generate(input, secretKey)
                    Dim comparer As StringComparer = StringComparer.OrdinalIgnoreCase
                    Return If(comparer.Compare(sourceHash, hash) = 0, True, False)
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.MD5.MD5Base64")
                    Return False
                Finally
                    GC.Collect()
                End Try
            End Function

        End Class

    End Class
End Namespace