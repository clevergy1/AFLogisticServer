Imports System.Net
Imports System.IO
Imports System.Runtime.Serialization.Json
Imports System.Security.Cryptography
Imports System.Net.Sockets

Public Class hsdsm
    Public Function send(t As Integer, a As String, p As String, d As String) As gmDsmResponse
        Dim retVal As gmDsmResponse = New gmDsmResponse
        If t = 2 Then
            retVal = ElaborGatemanager(p, d)
        End If

        If t = 3 Then
            retVal = ElaborTcpDirect(a, p, d)
        End If

        Return retVal
    End Function

    Private Function ElaborGatemanager(p As Integer, d As String) As gmDsmResponse
        Dim plainText As String
        Dim cipherText As String

        Dim passPhrase As String
        Dim saltValue As String
        Dim hashAlgorithm As String
        Dim passwordIterations As Integer
        Dim initVector As String
        Dim keySize As Integer

        plainText = p & ";" & d     ' original plaintext

        passPhrase = "leMejci@ver"
        saltValue = "kest@rd@vvedd@ar"
        hashAlgorithm = "SHA1"          ' can be "MD5"
        passwordIterations = 2          ' can be any number
        'initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
        initVector = "a#B6kmL9§f0kFig@"
        keySize = 256                   ' can be 192 or 128

        Console.WriteLine(String.Format("Plaintext : {0}", plainText))

        cipherText = RijndaelSimple.Encrypt _
        ( _
            plainText, _
            passPhrase, _
            saltValue, _
            hashAlgorithm, _
            passwordIterations, _
            initVector, _
            keySize _
        )



        'Dim postData As String = "p=" & p & "&d=" & d
        Dim postData As String = "x=" & cipherText

        Dim responseFromServer As String = String.Empty
        ' Create a request using a URL that can receive a post. 
        Dim request As WebRequest = WebRequest.Create("http://clevergyimpianti.it/cygmsrv/fup.ashx")
        'Dim request As WebRequest = WebRequest.Create("http://localhost:59604/fup.ashx")

        ' Set the Method property of the request to POST.
        request.Method = "POST"
        ' Create POST data and convert it to a byte array.

        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
        ' Set the ContentType property of the WebRequest.
        request.ContentType = "application/x-www-form-urlencoded"

        ' Set the ContentLength property of the WebRequest.
        request.ContentLength = byteArray.Length
        ' Get the request stream.
        Dim dataStream As Stream = request.GetRequestStream()
        ' Write the data to the request stream.
        dataStream.Write(byteArray, 0, byteArray.Length)
        ' Close the Stream object.
        dataStream.Close()
        ' Get the response.

        Dim retVal As gmDsmResponse = New gmDsmResponse

        Try
            Dim response As WebResponse = request.GetResponse()
            Dim statusCode As HttpStatusCode = CType(response, HttpWebResponse).StatusCode

            dataStream = response.GetResponseStream()
            '' Open the stream using a StreamReader for easy access.
            Dim reader As New StreamReader(dataStream)
            '' Read the content.
            responseFromServer = reader.ReadToEnd()

            Dim ser As New DataContractJsonSerializer(GetType(gmDsmResponse))
            Dim ms As New MemoryStream(Encoding.UTF8.GetBytes(responseFromServer))
            'Dim ms As New MemoryStream(Encoding.UTF8.GetBytes(j))
            Dim _obj As gmDsmResponse = DirectCast(ser.ReadObject(ms), gmDsmResponse)
            ' Clean up the streams.
            reader.Close()
            dataStream.Close()
            response.Close()


            If statusCode = HttpStatusCode.OK Then
                retVal = _obj
            End If

        Catch ex As Exception

        End Try


        Return retVal
    End Function

    Private Function ElaborTcpDirect(a As String, p As String, d As String) As gmDsmResponse
        Dim retVal As gmDsmResponse = New gmDsmResponse

        Dim tcpClient As New System.Net.Sockets.TcpClient()

        tcpClient.Connect(a, CInt(p))
        Dim networkStream As NetworkStream = tcpClient.GetStream()
        networkStream.ReadTimeout = 2000


        If networkStream.CanWrite And networkStream.CanRead Then
            ' Do a  write.
            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(d)

            networkStream.Write(sendBytes, 0, sendBytes.Length)

            ' Read the NetworkStream into a byte buffer.
            Dim buffSize As Integer = tcpClient.ReceiveBufferSize
            tcpClient.ReceiveBufferSize = 1024 ' 32768

            Try
                System.Threading.Thread.Sleep(200)
                Dim bytes(tcpClient.ReceiveBufferSize) As Byte
                networkStream.Read(bytes, 0, CInt(tcpClient.ReceiveBufferSize))
                ' Output the data received from the host to the console.
                Dim _data As String = Encoding.ASCII.GetString(bytes)
                retVal.returnErrormessage = String.Empty
                retVal.returnData = _data
                retVal.returnValue = True
            Catch ex As Exception
                retVal.returnValue = False
                retVal.returnErrormessage = "connection timeout"
            End Try

            tcpClient.Close()
        Else
            If Not networkStream.CanRead Then
                retVal.returnValue = False
                retVal.returnErrormessage = "cannot read data from this stream"
                tcpClient.Close()
            Else
                If Not networkStream.CanWrite Then
                    retVal.returnErrormessage = "cannot not write data to this stream"
                    tcpClient.Close()
                End If
            End If
        End If

        Return retVal
    End Function
End Class

Public Class gmDsmResponse
    Public Property returnValue As Boolean
    Public Property returnData As String
    Public Property returnErrormessage As String
End Class

Public Class RijndaelSimple

    ' <summary>
    ' Encrypts specified plaintext using Rijndael symmetric key algorithm
    ' and returns a base64-encoded result.
    ' </summary>
    ' <param name="plainText">
    ' Plaintext value to be encrypted.
    ' </param>
    ' <param name="passPhrase">
    ' Passphrase from which a pseudo-random password will be derived. The 
    ' derived password will be used to generate the encryption key. 
    ' Passphrase can be any string. In this example we assume that this 
    ' passphrase is an ASCII string.
    ' </param>
    ' <param name="saltValue">
    ' Salt value used along with passphrase to generate password. Salt can 
    ' be any string. In this example we assume that salt is an ASCII string.
    ' </param>
    ' <param name="hashAlgorithm">
    ' Hash algorithm used to generate password. Allowed values are: "MD5" and
    ' "SHA1". SHA1 hashes are a bit slower, but more secure than MD5 hashes.
    ' </param>
    ' <param name="passwordIterations">
    ' Number of iterations used to generate password. One or two iterations
    ' should be enough.
    ' </param>
    ' <param name="initVector">
    ' Initialization vector (or IV). This value is required to encrypt the 
    ' first block of plaintext data. For RijndaelManaged class IV must be 
    ' exactly 16 ASCII characters long.
    ' </param>
    ' <param name="keySize">
    ' Size of encryption key in bits. Allowed values are: 128, 192, and 256. 
    ' Longer keys are more secure than shorter keys.
    ' </param>
    ' <returns>
    ' Encrypted value formatted as a base64-encoded string.
    ' </returns>
    Public Shared Function Encrypt _
    ( _
        ByVal plainText As String, _
        ByVal passPhrase As String, _
        ByVal saltValue As String, _
        ByVal hashAlgorithm As String, _
        ByVal passwordIterations As Integer, _
        ByVal initVector As String, _
        ByVal keySize As Integer _
    ) _
    As String

        ' Convert strings into byte arrays.
        ' Let us assume that strings only contain ASCII codes.
        ' If strings include Unicode characters, use Unicode, UTF7, or UTF8 
        ' encoding.
        Dim initVectorBytes As Byte()
        initVectorBytes = Encoding.ASCII.GetBytes(initVector)

        Dim saltValueBytes As Byte()
        saltValueBytes = Encoding.ASCII.GetBytes(saltValue)

        ' Convert our plaintext into a byte array.
        ' Let us assume that plaintext contains UTF8-encoded characters.
        Dim plainTextBytes As Byte()
        plainTextBytes = Encoding.UTF8.GetBytes(plainText)

        ' First, we must create a password, from which the key will be derived.
        ' This password will be generated from the specified passphrase and 
        ' salt value. The password will be created using the specified hash 
        ' algorithm. Password creation can be done in several iterations.
        Dim password As PasswordDeriveBytes
        password = New PasswordDeriveBytes _
        ( _
            passPhrase, _
            saltValueBytes, _
            hashAlgorithm, _
            passwordIterations _
        )

        ' Use the password to generate pseudo-random bytes for the encryption
        ' key. Specify the size of the key in bytes (instead of bits).
        Dim keyBytes As Byte()
        keyBytes = password.GetBytes(keySize / 8)

        ' Create uninitialized Rijndael encryption object.
        Dim symmetricKey As RijndaelManaged
        symmetricKey = New RijndaelManaged()

        ' It is reasonable to set encryption mode to Cipher Block Chaining
        ' (CBC). Use default options for other symmetric key parameters.
        symmetricKey.Mode = CipherMode.CBC

        ' Generate encryptor from the existing key bytes and initialization 
        ' vector. Key size will be defined based on the number of the key 
        ' bytes.
        Dim encryptor As ICryptoTransform
        encryptor = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes)

        ' Define memory stream which will be used to hold encrypted data.
        Dim memoryStream As MemoryStream
        memoryStream = New MemoryStream()

        ' Define cryptographic stream (always use Write mode for encryption).
        Dim cryptoStream As CryptoStream
        cryptoStream = New CryptoStream _
        ( _
            memoryStream, _
            encryptor, _
            CryptoStreamMode.Write _
        )
        ' Start encrypting.
        cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)

        ' Finish encrypting.
        cryptoStream.FlushFinalBlock()

        ' Convert our encrypted data from a memory stream into a byte array.
        Dim cipherTextBytes As Byte()
        cipherTextBytes = memoryStream.ToArray()

        ' Close both streams.
        memoryStream.Close()
        cryptoStream.Close()

        ' Convert encrypted data into a base64-encoded string.
        Dim cipherText As String
        cipherText = Convert.ToBase64String(cipherTextBytes)

        ' Return encrypted string.
        Encrypt = cipherText
    End Function

    ' <summary>
    ' Decrypts specified ciphertext using Rijndael symmetric key algorithm.
    ' </summary>
    ' <param name="cipherText">
    ' Base64-formatted ciphertext value.
    ' </param>
    ' <param name="passPhrase">
    ' Passphrase from which a pseudo-random password will be derived. The 
    ' derived password will be used to generate the encryption key. 
    ' Passphrase can be any string. In this example we assume that this 
    ' passphrase is an ASCII string.
    ' </param>
    ' <param name="saltValue">
    ' Salt value used along with passphrase to generate password. Salt can 
    ' be any string. In this example we assume that salt is an ASCII string.
    ' </param>
    ' <param name="hashAlgorithm">
    ' Hash algorithm used to generate password. Allowed values are: "MD5" and
    ' "SHA1". SHA1 hashes are a bit slower, but more secure than MD5 hashes.
    ' </param>
    ' <param name="passwordIterations">
    ' Number of iterations used to generate password. One or two iterations
    ' should be enough.
    ' </param>
    ' <param name="initVector">
    ' Initialization vector (or IV). This value is required to encrypt the 
    ' first block of plaintext data. For RijndaelManaged class IV must be 
    ' exactly 16 ASCII characters long.
    ' </param>
    ' <param name="keySize">
    ' Size of encryption key in bits. Allowed values are: 128, 192, and 256. 
    ' Longer keys are more secure than shorter keys.
    ' </param>
    ' <returns>
    ' Decrypted string value.
    ' </returns>
    ' <remarks>
    ' Most of the logic in this function is similar to the Encrypt 
    ' logic. In order for decryption to work, all parameters of this function
    ' - except cipherText value - must match the corresponding parameters of 
    ' the Encrypt function which was called to generate the 
    ' ciphertext.
    ' </remarks>
    Public Shared Function Decrypt _
    ( _
        ByVal cipherText As String, _
        ByVal passPhrase As String, _
        ByVal saltValue As String, _
        ByVal hashAlgorithm As String, _
        ByVal passwordIterations As Integer, _
        ByVal initVector As String, _
        ByVal keySize As Integer _
    ) _
    As String

        ' Convert strings defining encryption key characteristics into byte
        ' arrays. Let us assume that strings only contain ASCII codes.
        ' If strings include Unicode characters, use Unicode, UTF7, or UTF8
        ' encoding.
        Dim initVectorBytes As Byte()
        initVectorBytes = Encoding.ASCII.GetBytes(initVector)

        Dim saltValueBytes As Byte()
        saltValueBytes = Encoding.ASCII.GetBytes(saltValue)

        ' Convert our ciphertext into a byte array.
        Dim cipherTextBytes As Byte()
        cipherTextBytes = Convert.FromBase64String(cipherText)

        ' First, we must create a password, from which the key will be 
        ' derived. This password will be generated from the specified 
        ' passphrase and salt value. The password will be created using
        ' the specified hash algorithm. Password creation can be done in
        ' several iterations.
        Dim password As PasswordDeriveBytes
        password = New PasswordDeriveBytes _
        ( _
            passPhrase, _
            saltValueBytes, _
            hashAlgorithm, _
            passwordIterations _
        )

        ' Use the password to generate pseudo-random bytes for the encryption
        ' key. Specify the size of the key in bytes (instead of bits).
        Dim keyBytes As Byte()
        keyBytes = password.GetBytes(keySize / 8)

        ' Create uninitialized Rijndael encryption object.
        Dim symmetricKey As RijndaelManaged
        symmetricKey = New RijndaelManaged()

        ' It is reasonable to set encryption mode to Cipher Block Chaining
        ' (CBC). Use default options for other symmetric key parameters.
        symmetricKey.Mode = CipherMode.CBC

        ' Generate decryptor from the existing key bytes and initialization 
        ' vector. Key size will be defined based on the number of the key 
        ' bytes.
        Dim decryptor As ICryptoTransform
        decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes)

        ' Define memory stream which will be used to hold encrypted data.
        Dim memoryStream As MemoryStream
        memoryStream = New MemoryStream(cipherTextBytes)

        ' Define memory stream which will be used to hold encrypted data.
        Dim cryptoStream As CryptoStream
        cryptoStream = New CryptoStream _
        ( _
            memoryStream, _
            decryptor, _
            CryptoStreamMode.Read _
        )

        ' Since at this point we don't know what the size of decrypted data
        ' will be, allocate the buffer long enough to hold ciphertext;
        ' plaintext is never longer than ciphertext.
        Dim plainTextBytes As Byte()
        ReDim plainTextBytes(cipherTextBytes.Length)

        ' Start decrypting.
        Dim decryptedByteCount As Integer
        decryptedByteCount = cryptoStream.Read _
        ( _
            plainTextBytes, _
            0, _
            plainTextBytes.Length _
        )

        ' Close both streams.
        memoryStream.Close()
        cryptoStream.Close()

        ' Convert decrypted data into a string. 
        ' Let us assume that the original plaintext string was UTF8-encoded.
        Dim plainText As String
        plainText = Encoding.UTF8.GetString _
        ( _
            plainTextBytes, _
            0, _
            decryptedByteCount _
        )

        ' Return decrypted string.
        Decrypt = plainText
    End Function
End Class
