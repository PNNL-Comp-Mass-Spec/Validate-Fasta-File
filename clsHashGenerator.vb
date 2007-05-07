Option Strict On

' This class will compute a hash value for a given string
' 
' -------------------------------------------------------------------------------
' Written by Ken Auberry for the Department of Energy (PNNL, Richland, WA) in 2006
' 
' E-mail: kenneth.auberry@pnl.gov or matthew.monroe@pnl.gov
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
' 
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0
'
' Notice: This computer software was prepared by Battelle Memorial Institute, 
' hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
' Department of Energy (DOE).  All rights in the computer software are reserved 
' by DOE on behalf of the United States Government and the Contractor as 
' provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
' WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
' SOFTWARE.  This notice including this sentence must appear on any copies of 
' this computer software.

Public Class clsHashGenerator
    Public Function GenerateHash(ByVal SourceText As String) As String
        Static objSHA1Provider As System.Security.Cryptography.SHA1Managed

        If objSHA1Provider Is Nothing Then
            objSHA1Provider = New System.Security.Cryptography.SHA1Managed
        End If

        'Create an encoding object to ensure the encoding standard for the source text
        Dim Ue As New System.Text.ASCIIEncoding

        'Retrieve a byte array based on the source text
        Dim ByteSourceText() As Byte = Ue.GetBytes(SourceText)

        'Compute the hash value from the source
        Dim SHA1_hash() As Byte = objSHA1Provider.ComputeHash(ByteSourceText)

        'And convert it to String format for return
        Dim SHA1string As String = ToHexString(SHA1_hash)

        Return SHA1string
    End Function

    Public Function ToHexString(ByVal bytes() As Byte) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder(bytes.Length * 2)

        For i = 0 To bytes.Length - 1
            sb.Append(bytes(i).ToString("X").PadLeft(2, "0"c))
        Next i

        Return sb.ToString.ToUpper

    End Function


End Class
