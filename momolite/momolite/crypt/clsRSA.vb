
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
Imports System.Security.Cryptography

Namespace crypt
    ''' <summary>Class RSA</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>19 April 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>1.1.0</version>
    ''' <requirement>
    ''' - Imports System.Text
    ''' - Imports System.Windows.Forms
    ''' - Imports System.Security.Cryptography
    ''' </requirement>
    Public Class RSA

        ''' <summary>
        ''' Generate Public dan Private Keys dalam bentuk Hashtable
        ''' </summary>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>Hashtable</returns>
        ''' 
        ''' <example>Contoh untuk memanggil function generate dalam bentuk Hashtable adalah sebagai berikut:
        ''' Dim keys As Hashtable = rsa.GenerateKeysHashtable
        ''' TextBox1.Text = keys.Item(0)
        ''' TextBox2.Text = keys.Item(1)
        ''' </example>
        Public Function GenerateKeys(Optional showException As Boolean = True) As Hashtable
            Try
                Dim cspParam As CspParameters = New CspParameters
                cspParam.Flags = CspProviderFlags.UseMachineKeyStore
                Dim RSA As RSACryptoServiceProvider = New RSACryptoServiceProvider

                Dim keys As Hashtable = New Hashtable

                keys.Add(0, RSA.ToXmlString(False))
                keys.Add(1, RSA.ToXmlString(True))

                Return keys
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.RSA")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function

        ''' <summary>
        ''' Proses RSA Encode input string
        ''' </summary>
        ''' <param name="input">String yang akan di enkripsi RSA</param>
        ''' <param name="publicKey">Public Key yang akan digunakan untuk syarat proses enkripsi RSA</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        ''' <remarks>Proses enkripsi memerlukan Public Key</remarks>
        Public Function Encode(input As String, publicKey As String, Optional showException As Boolean = True) As String
            Try
                Dim RSA As RSACryptoServiceProvider = New RSACryptoServiceProvider
                RSA.FromXmlString(publicKey)
                Dim decrypt As Byte() = Encoding.Unicode.GetBytes(input)
                Dim encrypt As Byte() = RSA.Encrypt(decrypt, False)
                Return System.Convert.ToBase64String(encrypt)
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.RSA")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function

        ''' <summary>
        ''' Proses Decode input string dari enkripsi RSA
        ''' </summary>
        ''' <param name="input">Enkripsi RSA yang akan di dekripsi</param>
        ''' <param name="privateKey">Private Key yang akan digunakan untuk syarat proses dekripsi RSA</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        ''' <remarks>Proses dekripsi memerlukan Private Key</remarks>
        Public Function Decode(input As String, privateKey As String, Optional showException As Boolean = True) As String
            Try
                Dim cspParam As CspParameters = New CspParameters
                cspParam.Flags = CspProviderFlags.UseMachineKeyStore
                Dim RSA As RSACryptoServiceProvider = New RSACryptoServiceProvider(cspParam)
                RSA.FromXmlString(privateKey)
                Dim encrypt As Byte() = System.Convert.FromBase64String(input)
                Dim decrypt As Byte() = RSA.Decrypt(encrypt, False)
                Return Encoding.Unicode.GetString(decrypt)
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.RSA")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function

    End Class
End Namespace