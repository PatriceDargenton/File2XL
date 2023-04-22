
' https://github.com/AutoItConsulting/text-encoding-detect

' Copyright 2015-2016 Jonathan Bennett <jon@autoitscript.com>
' 
' https://www.autoitscript.com 
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
' 
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

Option Infer On

'Namespace AutoIt.Common

Public Class TextEncodingDetect

    Private ReadOnly _utf16BeBom As Byte() = {&HFE, &HFF}

    Private ReadOnly _utf16LeBom As Byte() = {&HFF, &HFE}

    Private ReadOnly _utf8Bom As Byte() = {&HEF, &HBB, &HBF}

    Private _nullSuggestsBinary As Boolean = True
    Private _utf16ExpectedNullPercent As Double = 70
    Private _utf16UnexpectedNullPercent As Double = 10

    Public Enum Encoding

        ''' <summary>
        ''' Unknown or binary
        ''' </summary>
        None

        ''' <summary>
        ''' 0-255
        ''' </summary>
        Ansi

        ''' <summary>
        ''' 0-127
        ''' </summary>
        Ascii

        ''' <summary>
        ''' UTF8 with BOM
        ''' </summary>
        Utf8Bom

        ''' <summary>
        ''' UTF8 without BOM
        ''' </summary>
        Utf8Nobom

        ''' <summary>
        ''' UTF16 LE (Little Endian) with BOM: Unicode
        ''' </summary>
        Utf16LeBom

        ''' <summary>
        ''' UTF16 LE (Little Endian) without BOM: Unicode
        ''' </summary>
        Utf16LeNoBom

        ''' <summary>
        ''' UTF16-BE (Big Endian) with BOM
        ''' </summary>
        Utf16BeBom

        ''' <summary>
        ''' UTF16-BE (Big Endian) without BOM
        ''' </summary>
        Utf16BeNoBom

    End Enum

    ''' <summary>
    ''' Sets if the presence of nulls in a buffer indicate the buffer is binary data rather than text.
    ''' </summary>
    Public WriteOnly Property NullSuggestsBinary As Boolean
        Set(ByVal value As Boolean)
            _nullSuggestsBinary = value
        End Set
    End Property

    Public WriteOnly Property Utf16ExpectedNullPercent As Double
        Set(ByVal value As Double)
            If value > 0 AndAlso value < 100 Then
                _utf16ExpectedNullPercent = value
            End If
        End Set
    End Property

    Public WriteOnly Property Utf16UnexpectedNullPercent As Double
        Set(ByVal value As Double)
            If value > 0 AndAlso value < 100 Then
                _utf16UnexpectedNullPercent = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets the BOM length for a given Encoding mode.
    ''' </summary>
    ''' <param name="encoding"></param>
    ''' <returns>The BOM length.</returns>
    Public Shared Function GetBomLengthFromEncodingMode(ByVal encoding As Encoding) As Integer
        Dim length As Integer

        Select Case encoding
            Case Encoding.Utf16BeBom, Encoding.Utf16LeBom
                length = 2

            Case Encoding.Utf8Bom
                length = 3
            Case Else
                length = 0
        End Select

        Return length
    End Function

    ''' <summary>
    ''' Checks for a BOM sequence in a byte buffer.
    ''' </summary>
    ''' <param name="buffer"></param>
    ''' <param name="size"></param>
    ''' <returns>Encoding type or Encoding.None if no BOM.</returns>
    Public Function CheckBom(ByVal buffer As Byte(), ByVal size As Integer) As Encoding
        ' Check for BOM
        If size >= 2 AndAlso buffer(0) = _utf16LeBom(0) AndAlso buffer(1) = _utf16LeBom(1) Then
            Return Encoding.Utf16LeBom
        End If

        If size >= 2 AndAlso buffer(0) = _utf16BeBom(0) AndAlso buffer(1) = _utf16BeBom(1) Then
            Return Encoding.Utf16BeBom
        End If

        If size >= 3 AndAlso buffer(0) = _utf8Bom(0) AndAlso buffer(1) = _utf8Bom(1) AndAlso buffer(2) = _utf8Bom(2) Then
            Return Encoding.Utf8Bom
        End If

        Return Encoding.None
    End Function

    ''' <summary>
    ''' Automatically detects the Encoding type of a given byte buffer.
    ''' </summary>
    ''' <param name="buffer">The byte buffer.</param>
    ''' <param name="size">The size of the byte buffer.</param>
    ''' <returns>The Encoding type or Encoding.None if unknown.</returns>
    Public Function DetectEncoding(ByVal buffer As Byte(), ByVal size As Integer) As Encoding
        ' First check if we have a BOM and return that if so
        Dim encoding = CheckBom(buffer, size)
        If encoding <> Encoding.None Then
            Return encoding
        End If

        ' Now check for valid UTF8
        encoding = CheckUtf8(buffer, size)
        If encoding <> Encoding.None Then
            Return encoding
        End If

        ' Now try UTF16 
        encoding = CheckUtf16NewlineChars(buffer, size)
        If encoding <> Encoding.None Then
            Return encoding
        End If

        encoding = CheckUtf16Ascii(buffer, size)
        If encoding <> Encoding.None Then
            Return encoding
        End If

        ' ANSI or None (binary) then
        If Not DoesContainNulls(buffer, size) Then
            Return Encoding.Ansi
        End If

        ' Found a null, return based on the preference in null_suggests_binary_
        Return If(_nullSuggestsBinary, Encoding.None, Encoding.Ansi)
    End Function

    ''' <summary>
    ''' Checks if a buffer contains text that looks like utf16 by scanning for
    '''  newline chars that would be present even in non-english text.
    ''' </summary>
    ''' <param name="buffer">The byte buffer.</param>
    ''' <param name="size">The size of the byte buffer.</param>
    ''' <returns>Encoding.none, Encoding.Utf16LeNoBom or Encoding.Utf16BeNoBom.</returns>
    Private Shared Function CheckUtf16NewlineChars(ByVal buffer As Byte(), ByVal size As Integer) As Encoding
        If size < 2 Then
            Return Encoding.None
        End If

        ' Reduce size by 1 so we don't need to worry about bounds checking for pairs of bytes
        size -= 1

        Dim leControlChars = 0
        Dim beControlChars = 0

        'Dim pos As UInteger = 0
        Dim pos As Integer = 0
        While pos < size
            Dim ch1 = buffer(Math.Min(Threading.Interlocked.Increment(pos), pos - 1))
            Dim ch2 = buffer(Math.Min(Threading.Interlocked.Increment(pos), pos - 1))

            If ch1 = 0 Then
                If ch2 = &HA OrElse ch2 = &HD Then
                    Threading.Interlocked.Increment(beControlChars)
                End If
            ElseIf ch2 = 0 Then
                If ch1 = &HA OrElse ch1 = &HD Then
                    Threading.Interlocked.Increment(leControlChars)
                End If
            End If

            ' If we are getting both LE and BE control chars then this file is not utf16
            If leControlChars > 0 AndAlso beControlChars > 0 Then
                Return Encoding.None
            End If
        End While

        If leControlChars > 0 Then
            Return Encoding.Utf16LeNoBom
        End If

        Return If(beControlChars > 0, Encoding.Utf16BeNoBom, Encoding.None)
    End Function

    ''' <summary>
    ''' Checks if a buffer contains any nulls. Used to check for binary vs text data.
    ''' </summary>
    ''' <param name="buffer">The byte buffer.</param>
    ''' <param name="size">The size of the byte buffer.</param>
    Private Shared Function DoesContainNulls(ByVal buffer As Byte(), ByVal size As Integer) As Boolean
        'Dim pos As UInteger = 0
        Dim pos As Integer = 0
        While pos < size
            If buffer(Math.Min(Threading.Interlocked.Increment(pos), pos - 1)) = 0 Then
                Return True
            End If
        End While

        Return False
    End Function

    ''' <summary>
    ''' Checks if a buffer contains text that looks like utf16. This is done based
    '''  on the use of nulls which in ASCII/script like text can be useful to identify.
    ''' </summary>
    ''' <param name="buffer">The byte buffer.</param>
    ''' <param name="size">The size of the byte buffer.</param>
    ''' <returns>Encoding.none, Encoding.Utf16LeNoBom or Encoding.Utf16BeNoBom.</returns>
    Private Function CheckUtf16Ascii(ByVal buffer As Byte(), ByVal size As Integer) As Encoding
        Dim numOddNulls = 0
        Dim numEvenNulls = 0

        ' Get even nulls
        'Dim pos As UInteger = 0
        Dim pos As Integer = 0
        While pos < size
            If buffer(pos) = 0 Then
                numEvenNulls += 1
            End If

            pos += 2
        End While

        ' Get odd nulls
        pos = 1
        While pos < size
            If buffer(pos) = 0 Then
                numOddNulls += 1
            End If

            pos += 2
        End While

        Dim evenNullThreshold = numEvenNulls * 2.0 / size
        Dim oddNullThreshold = numOddNulls * 2.0 / size
        Dim expectedNullThreshold = _utf16ExpectedNullPercent / 100.0
        Dim unexpectedNullThreshold = _utf16UnexpectedNullPercent / 100.0

        ' Lots of odd nulls, low number of even nulls
        If evenNullThreshold < unexpectedNullThreshold AndAlso oddNullThreshold > expectedNullThreshold Then
            Return Encoding.Utf16LeNoBom
        End If

        ' Lots of even nulls, low number of odd nulls
        If oddNullThreshold < unexpectedNullThreshold AndAlso evenNullThreshold > expectedNullThreshold Then
            Return Encoding.Utf16BeNoBom
        End If

        ' Don't know
        Return Encoding.None
    End Function

    ''' <summary>
    ''' Checks if a buffer contains valid utf8.
    ''' </summary>
    ''' <param name="buffer">The byte buffer.</param>
    ''' <param name="size">The size of the byte buffer.</param>
    ''' <returns>
    '''     Encoding type of Encoding.None (invalid UTF8), Encoding.Utf8NoBom (valid utf8 multibyte strings) or
    '''     Encoding.ASCII (data in 0.127 range).
    ''' </returns>
    Private Function CheckUtf8(ByVal buffer As Byte(), ByVal size As Integer) As Encoding
        ' UTF8 Valid sequences
        ' 0xxxxxxx  ASCII
        ' 110xxxxx 10xxxxxx  2-byte
        ' 1110xxxx 10xxxxxx 10xxxxxx  3-byte
        ' 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx  4-byte
        '
        ' Width in UTF8
        ' Decimal      Width
        ' 0-127        1 byte
        ' 194-223      2 bytes
        ' 224-239      3 bytes
        ' 240-244      4 bytes
        '
        ' Subsequent chars are in the range 128-191
        Dim onlySawAsciiRange = True
        'Dim pos As UInteger = 0
        Dim pos As Integer = 0

        While pos < size
            Dim ch = buffer(Math.Min(Threading.Interlocked.Increment(pos), pos - 1))

            If ch = 0 AndAlso _nullSuggestsBinary Then
                Return Encoding.None
            End If

            Dim moreChars As Integer
            If ch <= 127 Then
                ' 1 byte
                moreChars = 0
            ElseIf ch >= 194 AndAlso ch <= 223 Then
                ' 2 Byte
                moreChars = 1
            ElseIf ch >= 224 AndAlso ch <= 239 Then
                ' 3 Byte
                moreChars = 2
            ElseIf ch >= 240 AndAlso ch <= 244 Then
                ' 4 Byte
                moreChars = 3
            Else
                Return Encoding.None ' Not utf8
            End If

            ' Check secondary chars are in range if we are expecting any
            While moreChars > 0 AndAlso pos < size
                onlySawAsciiRange = False ' Seen non-ascii chars now

                ch = buffer(Math.Min(Threading.Interlocked.Increment(pos), pos - 1))
                If ch < 128 OrElse ch > 191 Then
                    Return Encoding.None ' Not utf8
                End If

                Threading.Interlocked.Decrement(moreChars)
            End While
        End While

        ' If we get to here then only valid UTF-8 sequences have been processed

        ' If we only saw chars in the range 0-127 then we can't assume UTF8 (the caller will need to decide)
        Return If(onlySawAsciiRange, Encoding.Ascii, Encoding.Utf8Nobom)

    End Function

End Class

'End Namespace