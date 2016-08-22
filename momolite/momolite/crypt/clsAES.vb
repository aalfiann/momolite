
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
    ''' <summary>Class AES</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>19 April 2015</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>1.1.0</version>
    ''' <requirement>
    ''' - Imports System.Windows.Forms
    ''' </requirement>
    Public Class AES

        ''' <summary>
        ''' Membuat karakter string menjadi enkripsi AES
        ''' </summary>
        ''' <param name="input">Karakter string yang akan di enkripsi</param>
        ''' <param name="secretKey">Karakter string untuk dijadikan kata sandi dalam enkripsi AES</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        Public Function Encode(ByVal input As String, ByVal secretKey As String, Optional showException As Boolean = True) As String
            Dim AES As New System.Security.Cryptography.RijndaelManaged
            Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
            Dim encrypted As String = Nothing
            Try
                Dim hash(31) As Byte
                Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(secretKey))
                Array.Copy(temp, 0, hash, 0, 16)
                Array.Copy(temp, 0, hash, 15, 16)
                AES.Key = hash
                AES.Mode = Security.Cryptography.CipherMode.ECB
                Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
                Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(input)
                encrypted = Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
                Return encrypted
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.AES")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function

        ''' <summary>
        ''' Membuat karakter dari enkripsi AES menjadi Dekripsi karakter string
        ''' </summary>
        ''' <param name="input">Karakter enkripsi yang akan di dekripsi</param>
        ''' <param name="secretKey">Karakter kata sandi string dari enkripsi AES</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        Public Function Decode(ByVal input As String, ByVal secretKey As String, Optional showException As Boolean = True) As String
            Dim AES As New System.Security.Cryptography.RijndaelManaged
            Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
            Dim decrypted As String = Nothing
            Try
                Dim hash(31) As Byte
                Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(secretKey))
                Array.Copy(temp, 0, hash, 0, 16)
                Array.Copy(temp, 0, hash, 15, 16)
                AES.Key = hash
                AES.Mode = Security.Cryptography.CipherMode.ECB
                Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
                Dim Buffer As Byte() = Convert.FromBase64String(input)
                decrypted = System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
                Return decrypted
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.AES")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function
    End Class
End Namespace