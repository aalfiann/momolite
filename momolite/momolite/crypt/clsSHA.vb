
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

Imports Crypto = System.Security.Cryptography
Imports System.Windows.Forms
Imports System.Text
Namespace crypt
    ''' <summary>Class SHA</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>18 April 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>1.1.0</version>
    ''' <requirement>
    ''' - Imports System.Text
    ''' - Imports System.Windows.Forms
    ''' - Imports Crypto = System.Security.Cryptography
    ''' </requirement>
    Public Class SHA

        ''' <summary>
        ''' Enkripsi Hexadecimal menggunakan SHA1
        ''' </summary>
        Public Class SHA1

            ''' <summary>
            ''' Proses generate Hexadecimal SHA1
            ''' </summary>
            ''' <param name="input">Input string yang akan di enkripsi</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>String</returns>>
            Public Function Generate(input As String, secretKey As String, Optional showException As Boolean = True) As String
                Try
                    Dim SHA1 As Crypto.SHA1
                    SHA1 = Crypto.SHA1.Create
                    Dim hashData() As Byte = SHA1.ComputeHash(Encoding.Default.GetBytes(secretKey + input + secretKey))
                    Dim returnValue As StringBuilder = New StringBuilder
                    For i As Integer = 0 To hashData.Length - 1
                        returnValue.Append(hashData(i).ToString())
                    Next
                    Return returnValue.ToString()
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.SHA.SHA1")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Validasi Hexadecimal SHA1
            ''' </summary>
            ''' <param name="input">Input string untuk di komparasikan dengan Hexadecimal</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="hexadecimal">Input Hexadecimal yang akan di validasi</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function Validate(input As String, secretKey As String, hexadecimal As String, Optional showException As Boolean = True) As Boolean
                Try
                    Dim getHashInputData As String = Generate(input, secretKey)
                    Return If(String.Compare(getHashInputData, hexadecimal) = 0, True, False)
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.SHA.SHA1")
                    Return False
                End Try
            End Function
        End Class

        ''' <summary>
        ''' Enkripsi Hexadecimal menggunakan SHA256
        ''' </summary>
        Public Class SHA256
            ''' <summary>
            ''' Proses generate Hexadecimal SHA256
            ''' </summary>
            ''' <param name="input">Input string yang akan di enkripsi</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>String</returns>>
            Public Function Generate(input As String, secretKey As String, Optional showException As Boolean = True) As String
                Try
                    Dim SHA1 As Crypto.SHA256
                    SHA1 = Crypto.SHA256.Create
                    Dim hashData() As Byte = SHA1.ComputeHash(Encoding.Default.GetBytes(secretKey + input + secretKey))
                    Dim returnValue As StringBuilder = New StringBuilder
                    For i As Integer = 0 To hashData.Length - 1
                        returnValue.Append(hashData(i).ToString())
                    Next
                    Return returnValue.ToString()
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.SHA.SHA256")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Validasi Hexadecimal SHA256
            ''' </summary>
            ''' <param name="input">Input string untuk di komparasikan dengan Hexadecimal</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="hexadecimal">Input Hexadecimal yang akan di validasi</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function Validate(input As String, secretKey As String, hexadecimal As String, Optional showException As Boolean = True) As Boolean
                Try
                    Dim getHashInputData As String = Generate(input, secretKey)
                    Return If(String.Compare(getHashInputData, hexadecimal) = 0, True, False)
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.SHA.SHA256")
                    Return False
                End Try
            End Function
        End Class

        ''' <summary>
        ''' Enkripsi Hexadecimal menggunakan SHA384
        ''' </summary>
        Public Class SHA384

            ''' <summary>
            ''' Proses generate Hexadecimal SHA384
            ''' </summary>
            ''' <param name="input">Input string yang akan di enkripsi</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>String</returns>>
            Public Function Generate(input As String, secretKey As String, Optional showException As Boolean = True) As String
                Try
                    Dim SHA1 As Crypto.SHA384
                    SHA1 = Crypto.SHA384.Create
                    Dim hashData() As Byte = SHA1.ComputeHash(Encoding.Default.GetBytes(secretKey + input + secretKey))
                    Dim returnValue As StringBuilder = New StringBuilder
                    For i As Integer = 0 To hashData.Length - 1
                        returnValue.Append(hashData(i).ToString())
                    Next
                    Return returnValue.ToString()
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.SHA.SHA384")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Validasi Hexadecimal SHA384
            ''' </summary>
            ''' <param name="input">Input string untuk di komparasikan dengan Hexadecimal</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="hexadecimal">Input Hexadecimal yang akan di validasi</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function Validate(input As String, secretKey As String, hexadecimal As String, Optional showException As Boolean = True) As Boolean
                Try
                    Dim getHashInputData As String = Generate(input, secretKey)
                    Return If(String.Compare(getHashInputData, hexadecimal) = 0, True, False)
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.SHA.SHA384")
                    Return False
                End Try
            End Function
        End Class

        ''' <summary>
        ''' Enkripsi Hexadecimal menggunakan SHA512
        ''' </summary>
        Public Class SHA512

            ''' <summary>
            ''' Proses generate Hexadecimal SHA512
            ''' </summary>
            ''' <param name="input">Input string yang akan di enkripsi</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>String</returns>>
            Public Function Generate(input As String, secretKey As String, Optional showException As Boolean = True) As String
                Try
                    Dim SHA1 As Crypto.SHA512
                    SHA1 = Crypto.SHA512.Create
                    Dim hashData() As Byte = SHA1.ComputeHash(Encoding.Default.GetBytes(secretKey + input + secretKey))
                    Dim returnValue As StringBuilder = New StringBuilder
                    For i As Integer = 0 To hashData.Length - 1
                        returnValue.Append(hashData(i).ToString())
                    Next
                    Return returnValue.ToString()
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.SHA.SHA512")
                    Return Nothing
                Finally
                    GC.Collect()
                End Try
            End Function

            ''' <summary>
            ''' Validasi Hexadecimal SHA512
            ''' </summary>
            ''' <param name="input">Input string untuk di komparasikan dengan Hexadecimal</param>
            ''' <param name="secretKey">Input secret key value</param>
            ''' <param name="hexadecimal">Input Hexadecimal yang akan di validasi</param>
            ''' <param name="showException">Tampilkan log exception? Default = True</param>
            ''' <returns>Boolean</returns>
            Public Function Validate(input As String, secretKey As String, hexadecimal As String, Optional showException As Boolean = True) As Boolean
                Try
                    Dim getHashInputData As String = Generate(input, secretKey)
                    Return If(String.Compare(getHashInputData, hexadecimal) = 0, True, False)
                Catch ex As Exception
                    If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.SHA.SHA512")
                    Return False
                End Try
            End Function
        End Class

    End Class
End Namespace