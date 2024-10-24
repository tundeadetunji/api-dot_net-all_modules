'from Assemblies
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' Ligthweight Encryption/Decryption.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public Class Encryption
	''' <summary>
	''' Encrypts string with specified key.
	''' </summary>
	''' <param name="key"></param>
	''' <param name="data"></param>
	''' <returns>Encrypted string</returns>
	Public Shared Function Encrypt(ByVal key As String, ByVal data As String) As String
		Try
			If key.Length < 1 Or data.Length < 1 Then Exit Function

			Dim encData As String = Nothing
			Dim keys As Byte()() = GetHashKeys(key)

			Try
				encData = EncryptStringToBytes_Aes(data, keys(0), keys(1))
			Catch __unusedCryptographicException1__ As CryptographicException
			Catch __unusedArgumentNullException2__ As ArgumentNullException
			End Try

			Return encData

		Catch ex As Exception
		End Try
	End Function

	''' <summary>
	''' Decrypts string with specified key.
	''' </summary>
	''' <param name="key"></param>
	''' <param name="data"></param>
	''' <returns>Decrypted string</returns>
	Public Shared Function Decrypt(ByVal key As String, ByVal data As String) As String
		Try

			Dim decData As String = Nothing
			Dim keys As Byte()() = GetHashKeys(key)

			Try
				decData = DecryptStringFromBytes_Aes(data, keys(0), keys(1))
			Catch __unusedCryptographicException1__ As CryptographicException
			Catch __unusedArgumentNullException2__ As ArgumentNullException
			End Try

			Return decData

		Catch ex As Exception
		End Try
	End Function

	Private Shared Function GetHashKeys(ByVal key As String) As Byte()()
		Dim result As Byte()() = New Byte(1)() {}
		Dim enc As Encoding = Encoding.UTF8
		Dim sha2 As SHA256 = New SHA256CryptoServiceProvider()
		Dim rawKey As Byte() = enc.GetBytes(key)
		Dim rawIV As Byte() = enc.GetBytes(key)
		Dim hashKey As Byte() = sha2.ComputeHash(rawKey)
		Dim hashIV As Byte() = sha2.ComputeHash(rawIV)
		Array.Resize(hashIV, 16)
		result(0) = hashKey
		result(1) = hashIV
		Return result
	End Function

	Private Shared Function EncryptStringToBytes_Aes(ByVal plainText As String, ByVal Key As Byte(), ByVal IV As Byte()) As String
		If plainText Is Nothing OrElse plainText.Length <= 0 Then Throw New ArgumentNullException("plainText")
		If Key Is Nothing OrElse Key.Length <= 0 Then Throw New ArgumentNullException("Key")
		If IV Is Nothing OrElse IV.Length <= 0 Then Throw New ArgumentNullException("IV")
		Dim encrypted As Byte()

		Using aesAlg As AesManaged = New AesManaged()
			aesAlg.Key = Key
			aesAlg.IV = IV
			Dim encryptor As ICryptoTransform = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV)

			Using msEncrypt As MemoryStream = New MemoryStream()

				Using csEncrypt As CryptoStream = New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)

					Using swEncrypt As StreamWriter = New StreamWriter(csEncrypt)
						swEncrypt.Write(plainText)
					End Using

					encrypted = msEncrypt.ToArray()
				End Using
			End Using
		End Using

		Return Convert.ToBase64String(encrypted)
	End Function

	Private Shared Function DecryptStringFromBytes_Aes(ByVal cipherTextString As String, ByVal Key As Byte(), ByVal IV As Byte()) As String
		Dim cipherText As Byte() = Convert.FromBase64String(cipherTextString)
		If cipherText Is Nothing OrElse cipherText.Length <= 0 Then Throw New ArgumentNullException("cipherText")
		If Key Is Nothing OrElse Key.Length <= 0 Then Throw New ArgumentNullException("Key")
		If IV Is Nothing OrElse IV.Length <= 0 Then Throw New ArgumentNullException("IV")
		Dim plaintext As String = Nothing

		Using aesAlg As Aes = Aes.Create()
			aesAlg.Key = Key
			aesAlg.IV = IV
			Dim decryptor As ICryptoTransform = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV)

			Using msDecrypt As MemoryStream = New MemoryStream(cipherText)

				Using csDecrypt As CryptoStream = New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)

					Using srDecrypt As StreamReader = New StreamReader(csDecrypt)
						plaintext = srDecrypt.ReadToEnd()
					End Using
				End Using
			End Using
		End Using

		Return plaintext
	End Function

End Class
