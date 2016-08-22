
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

Imports System.Security.Cryptography
Imports System.Windows.Forms
Namespace crypt
    ''' <summary>Class DES dan 3DES Cryptography</summary>
    ''' <author>M ABD AZIZ ALFIAN</author>
    ''' <lastupdate>18 April 2016</lastupdate>
    ''' <url>http://about.me/azizalfian</url>
    ''' <version>2.2.0</version>
    ''' <requirement>
    ''' - Imports System.Windows.Forms
    ''' - Imports System.Security.Cryptography
    ''' </requirement>
    Public Class DES
#Region "Declaration"
        Private mDes As New DESCryptoServiceProvider
#End Region

#Region "Proces Data"
        Private Function TruncateHash(
        ByVal key As String,
        ByVal length As Integer) As Byte()

            Dim sha1 As New SHA1CryptoServiceProvider

            ' Hash the key. 
            Dim keyBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(key)
            Dim hash() As Byte = sha1.ComputeHash(keyBytes)

            ' Truncate or pad the hash. 
            ReDim Preserve hash(length - 1)
            Return hash
        End Function

        Private Function EncryptData(ByVal plaintext As String) As String

            ' Convert the plaintext string to a byte array. 
            Dim plaintextBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(plaintext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream. 
            Dim encStream As New CryptoStream(ms,
                mDes.CreateEncryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()

            ' Convert the encrypted stream to a printable string. 
            Return Convert.ToBase64String(ms.ToArray)
        End Function

        Private Function DecryptData(
        ByVal encryptedtext As String) As String

            ' Convert the encrypted text string to a byte array. 
            Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the decoder to write to the stream. 
            Dim decStream As New CryptoStream(ms,
                mDes.CreateDecryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()

            ' Convert the plaintext stream to a string. 
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        End Function
#End Region

        ''' <summary>
        ''' Membuat enkripsi DES
        ''' </summary>
        ''' <param name="input">Karater string yang akan di enkripsi</param>
        ''' <param name="secretKey">Input secret key value</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        Public Function Encode(input As String, secretKey As String, Optional showException As Boolean = True) As String
            Try
                mDes.Key = TruncateHash(secretKey, mDes.KeySize \ 8)
                mDes.IV = TruncateHash("", mDes.BlockSize \ 8)
                Dim cipherText As String = EncryptData(input)
                Return cipherText
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.DES")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function

        ''' <summary>
        ''' Mendekripsi karakter dari enkripsi DES
        ''' </summary>
        ''' <param name="input">Karater string yang akan di dekripsi</param>
        ''' <param name="secretKey">Input secret key value</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        Public Function Decode(input As String, secretKey As String, Optional showException As Boolean = True) As String
            Try
                mDes.Key = TruncateHash(secretKey, mDes.KeySize \ 8)
                mDes.IV = TruncateHash("", mDes.BlockSize \ 8)
                Dim plainText As String = DecryptData(input)
                Return plainText
            Catch ex As System.Security.Cryptography.CryptographicException
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.DES")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function

    End Class

    ''' <summary>
    ''' Class Triple DES
    ''' </summary>
    Public Class TripleDES
#Region "Declaration"
        Private TripleDes As New TripleDESCryptoServiceProvider
#End Region

#Region "Proces Data"
        Private Function TruncateHash(
        ByVal key As String,
        ByVal length As Integer) As Byte()

            Dim sha1 As New SHA1CryptoServiceProvider

            ' Hash the key. 
            Dim keyBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(key)
            Dim hash() As Byte = sha1.ComputeHash(keyBytes)

            ' Truncate or pad the hash. 
            ReDim Preserve hash(length - 1)
            Return hash
        End Function

        Private Function EncryptData(ByVal plaintext As String) As String

            ' Convert the plaintext string to a byte array. 
            Dim plaintextBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(plaintext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream. 
            Dim encStream As New CryptoStream(ms,
                TripleDes.CreateEncryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()

            ' Convert the encrypted stream to a printable string. 
            Return Convert.ToBase64String(ms.ToArray)
        End Function

        Private Function DecryptData(
        ByVal encryptedtext As String) As String

            ' Convert the encrypted text string to a byte array. 
            Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the decoder to write to the stream. 
            Dim decStream As New CryptoStream(ms,
                TripleDes.CreateDecryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()

            ' Convert the plaintext stream to a string. 
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        End Function
#End Region

        ''' <summary>
        ''' Membuat enkripsi TripleDES
        ''' </summary>
        ''' <param name="input">Karater string yang akan di enkripsi</param>
        ''' <param name="secretKey">Input secret key value</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        Public Function Encode(input As String, secretKey As String, Optional showException As Boolean = True) As String
            Try
                TripleDes.Key = TruncateHash(secretKey, TripleDes.KeySize \ 8)
                TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
                Dim cipherText As String = EncryptData(input)
                Return cipherText
            Catch ex As Exception
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.TripleDES")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function

        ''' <summary>
        ''' Mendekripsi karakter dari enkripsi TripleDES
        ''' </summary>
        ''' <param name="input">Karater string yang akan di dekripsi</param>
        ''' <param name="secretKey">Input secret key value</param>
        ''' <param name="showException">Tampilkan log exception? Default = True</param>
        ''' <returns>String</returns>
        Public Function Decode(input As String, secretKey As String, Optional showException As Boolean = True) As String
            Try
                TripleDes.Key = TruncateHash(secretKey, TripleDes.KeySize \ 8)
                TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
                Dim plainText As String = DecryptData(input)
                Return plainText
            Catch ex As System.Security.Cryptography.CryptographicException
                If showException = True Then momolite.globals.Dev.CatchException(ex, globals.Dev.Icons.Errors, "momolite.crypt.TripleDES")
                Return Nothing
            Finally
                GC.Collect()
            End Try
        End Function

    End Class
End Namespace