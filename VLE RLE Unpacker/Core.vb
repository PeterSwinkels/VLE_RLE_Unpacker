'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Math

'This module contains this program's core procedures.
Public Module CoreModule
   Private Const COMPRESSED_FILE_MAXIMUM_SIZE As Integer = &HFFFFFF%       'Defines the maximum size a compressed file can be.
   Private Const COMPRESSION_MULTIPLE_PASSES_FLAG As Integer = &H80%       'Defines the flag indicating whether multiple compression passes are present.
   Private Const COMPRESSION_PASS_COUNT_MASK As Integer = &H7F%            'Defines the bits indicating the number of passes.
   Private Const COMPRESSION_TYPE_RLE As Integer = &H1%                    'Defines the run-length-encoding compression type.
   Private Const COMPRESSION_TYPE_VLE As Integer = &H2%                    'Defines the variable-length-encoding compression type.
   Private Const RLE_ESCAPE_LENGTH_MASK As Integer = &H7F%                 'Defines the bits indicating the number of escape codes.
   Private Const RLE_ESCAPE_LENGTH_NO_SEQUENCE_RUN As Integer = &H80%      'Defines the bit indicating a run is not a sequence run.
   Private Const RLE_ESCAPE_LOOKUP_TABLE_LENGTH As Integer = &H100%        'Defines the escape character lookup table size.
   Private Const RLE_ESCAPE_MAXIMUM_LENGTH As Integer = &HA%               'Defines the maximum number of escape codes.
   Private Const RLE_SECOND_ESCAPE_CODE_POSITION As Integer = &H1%         'Defines the zero-based position of the second escape code.
   Private Const VLE_ALPHABET_LENGTH As Integer = &H100%                   'Defines the number of alphabet codes.
   Private Const VLE_BYTE_MSB_MASK As Integer = &H80%                      'Defines the most significant bit in a byte.
   Private Const VLE_ESCAPE_CHARACTERS_LENGTH As Integer = &H10%           'Defines the number of escape codes.
   Private Const VLE_ESCAPE_WIDTH As Integer = &H40%                       'Defines the symbol width indicating the start of the escape sequence.
   Private Const VLE_UNKNOWN_WIDTH_LENGTH As Integer = &H80%               'Defines the flag that should not be set in the widths lengths value.
   Private Const VLE_WIDTH_LENGTH_MASK As Integer = &H7F%                  'Defines the mask for the widths lengths value.
   Private Const VLE_WIDTH_MAXIMUM_LENGTH As Integer = &HF%                'Defines the maximum widths lengths value.

   'This structure defines a data buffer.
   Private Structure DataStr
      Public Data() As Byte        'Defines the data.
      Public Position As Integer   'Defines the position inside the data.
   End Structure

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim Source As New DataStr With {.Data = {}, .Position = 0}
         Dim Success As Boolean = False
         Dim Target As New DataStr With {.Data = {}, .Position = 0}

         If GetCommandLineArgs.Count = 3 Then
            If GetCommandLineArgs(1) = GetCommandLineArgs(2) Then
               Console.WriteLine("The target file cannot be the same as the source file.")
            Else
               Success = ReadCompressedFile(GetCommandLineArgs(1), Source)
               If Success Then
                  Success = Decompress(Source, Target)
                  If Success Then
                     Console.WriteLine($"Decompressed ""{GetCommandLineArgs(1)}"".")
                     Success = WriteDecompressedFile(GetCommandLineArgs(2), Target)
                     If Success Then
                        Console.WriteLine($"Wrote ""{GetCommandLineArgs(2)}"".")
                     Else
                        Console.WriteLine($"Could not write ""{GetCommandLineArgs(2)}"".")
                     End If
                  Else
                     Console.WriteLine($"Could not decompress ""{GetCommandLineArgs(1)}""")
                  End If
               Else
                  Console.WriteLine($"Could not read ""{GetCommandLineArgs(1)}"".")
               End If
            End If
         Else
            With My.Application.Info
               Console.WriteLine($"{ .Title} v{ .Version} - by: { .CompanyName}, { .Copyright}")
               Console.WriteLine()
               Console.WriteLine(.Description)
               Console.WriteLine("Supported compressed file extensions: *.cmn, *.cod, *.dif, *.p3s, *.pes, *.pre, *.pvs")
               Console.WriteLine()
               Console.WriteLine("Usage:")
               Console.WriteLine($"""{ .AssemblyName}.exe"" SOURCE_FILE TARGET_FILE")
            End With
         End If
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure decompresses the specified data.
   Private Function Decompress(ByRef Source As DataStr, ByRef Target As DataStr) As Boolean
      Try
         Dim CompressionType As Byte = 0
         Dim PassCount As Integer = Source.Data(Source.Position)
         Dim Success As Boolean = False

         If ((PassCount And COMPRESSION_MULTIPLE_PASSES_FLAG) = COMPRESSION_MULTIPLE_PASSES_FLAG) Then
            PassCount = PassCount And COMPRESSION_PASS_COUNT_MASK
            Source.Position += 4
         Else
            PassCount = 1
         End If

         If Source.Position <= Source.Data.Length Then
            For Pass As Integer = 0 To PassCount - 1
               CompressionType = Source.Data(Source.Position)
               Source.Position += 1

               ''If Target.Data.Count > 0 Then Erase Target.Data
               ReDim Target.Data(0 To GetSubFileSize(Source) - 1)

               If Target.Data.Count > 0 Then
                  Success = False

                  Select Case CompressionType
                     Case COMPRESSION_TYPE_RLE
                        Success = RLEDecompress(Source, Target)
                     Case COMPRESSION_TYPE_VLE
                        Success = VLEDecompress(Source, Target)
                  End Select

                  If Success AndAlso Pass < (PassCount - 1) Then
                     ''Erase Source.Data
                     Source.Position = 0
                     ReDim Source.Data(0 To Target.Data.Length - 1)
                     If Source.Data.Count = 0 Then
                        Exit For
                     Else
                        Array.Copy(Target.Data, Source.Data, Source.Data.Length)
                        Target.Position = 0
                     End If
                  Else
                     Exit For
                  End If
               End If
            Next Pass
         End If

         Return Success
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure displays any errors that occur.
   Private Sub DisplayError(ExceptionO As Exception)
      Try
         Console.WriteLine()
         Console.Error.WriteLine($"ERROR: {ExceptionO.Message}")
         Console.WriteLine()
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure returns a compressed sub file's size from the specified source.
   Private Function GetSubFileSize(ByRef Source As DataStr) As Integer
      Try
         Dim SubFileSize As Integer = Source.Data(Source.Position)

         SubFileSize = SubFileSize Or (CInt(Source.Data(Source.Position + 1)) << &H8%)
         SubFileSize = SubFileSize Or (CInt(Source.Data(Source.Position + 2)) << &H10%)
         Source.Position += 3

         Return SubFileSize
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure decodes RLE sequence runs inside the specified source and writes the result to the specified target.
   Private Function RLEDecodeSequenceRuns(ByRef Source As DataStr, ByRef Target As DataStr, SECOND_ESCAPE_CODE As Byte) As Boolean
      Try
         Dim CurrentByte As Byte = 0
         Dim SequenceOffset As Integer = 0
         Dim SequenceRunLength As Integer = 0

         While Source.Position < Source.Data.Length
            CurrentByte = Source.Data(Source.Position)
            Source.Position += 1

            If CurrentByte = SECOND_ESCAPE_CODE Then
               SequenceOffset = Source.Position

               CurrentByte = Source.Data(Source.Position)
               Source.Position += 1
               While Not CurrentByte = SECOND_ESCAPE_CODE
                  If Source.Position >= Source.Data.Length Then Return False

                  Target.Data(Target.Position) = CurrentByte
                  Target.Position += 1

                  CurrentByte = Source.Data(Source.Position)
                  Source.Position += 1
               End While

               SequenceRunLength = Source.Data(Source.Position) - 1
               Source.Position += 1

               While SequenceRunLength > 0
                  SequenceRunLength -= 1

                  For Index As Integer = 0 To (Source.Position - SequenceOffset - 2) - 1
                     If Target.Position >= Target.Data.Length Then Return False

                     Target.Data(Target.Position) = Source.Data(SequenceOffset + Index)
                     Target.Position += 1
                  Next Index
               End While
            Else
               Target.Data(Target.Position) = CurrentByte
               Target.Position += 1

               If Target.Position > Target.Data.Length Then Return False
            End If
         End While

         Return True
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure decodes a single-byte run inside the specified source and writes the result to the specified target.
   Private Function RLEDecodeSingleByteRun(ByRef Source As DataStr, ByRef Target As DataStr, ByteRunLength As Integer, ByteO As Byte) As Boolean
      Try
         Dim Success As Boolean = True

         While ByteRunLength > 0
            ByteRunLength -= 1
            If Target.Position >= Target.Data.Length Then
               Success = False
               Exit While
            End If
            Target.Data(Target.Position) = ByteO
            Target.Position += 1
         End While

         Return Success
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure decodes single-byte runs inside the specified source and writes the result to the specified target.
   Private Function RLEDecodeSingleByteRuns(ByRef Source As DataStr, ByRef Target As DataStr, EscapeLookupTable() As Integer) As Boolean
      Try
         Dim ByteRunLength As Integer = 0
         Dim CurrentByte As Byte = 0
         Dim EscapeCode As Integer = Nothing
         Dim Success As Boolean = True

         While Target.Position < Target.Data.Length
            CurrentByte = Source.Data(Source.Position)
            Source.Position += 1

            EscapeCode = EscapeLookupTable(CurrentByte)
            If Not (EscapeCode And &HFF%) = &H0% Then
               Select Case EscapeCode
                  Case 1
                     ByteRunLength = Source.Data(Source.Position)
                     Source.Position += 1
                     CurrentByte = Source.Data(Source.Position)
                     Source.Position += 1
                     Success = RLEDecodeSingleByteRun(Source, Target, ByteRunLength, CurrentByte)
                  Case 3
                     ByteRunLength = Source.Data(Source.Position) Or CInt(Source.Data(Source.Position + 1)) << &H8%
                     Source.Position += 2
                     CurrentByte = Source.Data(Source.Position)
                     Source.Position += 1
                     Success = RLEDecodeSingleByteRun(Source, Target, ByteRunLength, CurrentByte)
                  Case Else
                     ByteRunLength = EscapeLookupTable(CurrentByte) - 1
                     CurrentByte = Source.Data(Source.Position)
                     Source.Position += 1
                     Success = RLEDecodeSingleByteRun(Source, Target, ByteRunLength, CurrentByte)
               End Select
            Else
               Target.Data(Target.Position) = CurrentByte
               Target.Position += 1
            End If
         End While

         Return Success
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure decompresses RLE data inside the specified source and writes the result to the specified target.
   Private Function RLEDecompress(ByRef Source As DataStr, ByRef Target As DataStr) As Boolean
      Try
         Dim DecodedRLESequenceRunsTarget As New DataStr With {.Data = Nothing, .Position = 0}
         Dim EscapeCodes(0 To RLE_ESCAPE_MAXIMUM_LENGTH) As Byte
         Dim EscapeLength As Integer = 0
         Dim EscapeLookupTable(0 To RLE_ESCAPE_LOOKUP_TABLE_LENGTH) As Integer
         Dim Success As Boolean = False

         Source.Position += 4
         EscapeLength = Source.Data(Source.Position)
         Source.Position += 1

         If (EscapeLength And RLE_ESCAPE_LENGTH_MASK) > RLE_ESCAPE_MAXIMUM_LENGTH Then Return False

         For EscapeCodeIndex As Integer = 0 To (EscapeLength And RLE_ESCAPE_LENGTH_MASK) - 1
            EscapeCodes(EscapeCodeIndex) = Source.Data(Source.Position)
            Source.Position += 1
         Next EscapeCodeIndex

         If Source.Position > Source.Data.Length Then Return False

         For EscapeCodeIndex As Integer = 0 To (EscapeLength And RLE_ESCAPE_LENGTH_MASK) - 1
            EscapeLookupTable(EscapeCodes(EscapeCodeIndex)) = EscapeCodeIndex + 1
         Next EscapeCodeIndex

         If (EscapeLength And RLE_ESCAPE_LENGTH_NO_SEQUENCE_RUN) = RLE_ESCAPE_LENGTH_NO_SEQUENCE_RUN Then
            Success = RLEDecodeSingleByteRuns(Source, Target, EscapeLookupTable)
         Else
            DecodedRLESequenceRunsTarget.Position = Target.Position
            ReDim DecodedRLESequenceRunsTarget.Data(0 To Target.Data.Length - 1)

            If Not RLEDecodeSequenceRuns(Source, DecodedRLESequenceRunsTarget, EscapeCodes(RLE_SECOND_ESCAPE_CODE_POSITION)) Then
               Erase DecodedRLESequenceRunsTarget.Data
               Success = False
            Else
               ReDim Preserve DecodedRLESequenceRunsTarget.Data(0 To DecodedRLESequenceRunsTarget.Position)
               DecodedRLESequenceRunsTarget.Position = 0

               Success = RLEDecodeSingleByteRuns(DecodedRLESequenceRunsTarget, Target, EscapeLookupTable)
            End If
         End If

         Return Success
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure reads the specified compressed file.
   Private Function ReadCompressedFile(SourceFile As String, ByRef Source As DataStr) As Boolean
      Try
         Dim Success As Boolean = False

         If New FileInfo(SourceFile).Length <= COMPRESSED_FILE_MAXIMUM_SIZE Then
            Source.Data = File.ReadAllBytes(SourceFile)
            Success = True
         End If

         Return Success
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure decodes VLE compression codes inside the specified source and writes the result to the specified target.
   Private Function VLEDecode(ByRef Source As DataStr, ByRef Target As DataStr, ByRef Alphabet As Byte(), ByRef Symbols As Byte(), ByRef Widths As Byte(), ByRef EscapeCharacters1 As Integer(), ByRef EscapeCharacters2 As Integer()) As Boolean
      Try
         Dim CurrentSymbol As Integer = 0
         Dim CurrentWidth As Integer = 8
         Dim CurrentWord As Integer = 0
         Dim EscapeSequenceComplete As Boolean = False
         Dim EscapeSequenceIndex As Integer = 0
         Dim MSBBitSet As Boolean = False
         Dim NextWidth As Integer = 0

         CurrentWord = (CInt(Source.Data(Source.Position)) << &H8%) And &HFFFF%
         Source.Position += 1
         CurrentWord = (CurrentWord Or Source.Data(Source.Position)) And &HFFFF%
         Source.Position += 1

         While Target.Position < Target.Data.Length
            CurrentSymbol = (CurrentWord And &HFF00%) >> &H8%
            NextWidth = Widths(CurrentSymbol)

            If NextWidth > 8 Then
               If Not NextWidth = VLE_ESCAPE_WIDTH Then Return False

               CurrentSymbol = (CurrentWord And &HFF%)
               CurrentWord = (CurrentWord >> &H8%) And &HFFFF%
               EscapeSequenceIndex = 7
               EscapeSequenceComplete = False

               While Not EscapeSequenceComplete
                  If CurrentWidth = 0 Then
                     CurrentSymbol = Source.Data(Source.Position)
                     Source.Position += 1
                     CurrentWidth = 8
                  End If

                  MSBBitSet = ((CurrentSymbol And VLE_BYTE_MSB_MASK) = VLE_BYTE_MSB_MASK)
                  CurrentWord = ((CurrentWord << &H1%) Or Abs(CInt(MSBBitSet))) And &HFFFF%
                  CurrentSymbol <<= &H1%
                  CurrentWidth -= 1
                  EscapeSequenceIndex += 1

                  If EscapeSequenceIndex >= VLE_ESCAPE_CHARACTERS_LENGTH Then Return False

                  If (CurrentWord < EscapeCharacters2(EscapeSequenceIndex)) Then
                     CurrentWord = (CurrentWord + EscapeCharacters1(EscapeSequenceIndex)) And &HFFFF%
                     If (CurrentWord > &HFF%) Then Return False

                     Target.Data(Target.Position) = Alphabet(CurrentWord)
                     Target.Position += 1

                     EscapeSequenceComplete = True
                  End If
               End While


               If Source.Position < Source.Data.Length Then
                  CurrentWord = ((CurrentSymbol << CurrentWidth) Or Source.Data(Source.Position) And &HFFFF%)
               End If

               Source.Position += 1
               NextWidth = 8 - CurrentWidth
               CurrentWidth = 8
            Else
               Target.Data(Target.Position) = Symbols(CurrentSymbol)
               Target.Position += 1

               If CurrentWidth < NextWidth Then
                  CurrentWord = (CurrentWord << CurrentWidth) And &HFFFF%
                  NextWidth -= CurrentWidth
                  CurrentWidth = 8
                  If Source.Position >= Source.Data.Length Then Exit While
                  CurrentWord = (CurrentWord Or Source.Data(Source.Position)) And &HFFFF%
                  Source.Position += 1
               End If
            End If

            CurrentWord = (CurrentWord << NextWidth) And &HFFFF%
            CurrentWidth -= NextWidth

            If (Source.Position - 1) > Source.Data.Length AndAlso Target.Position < Target.Data.Length Then Return False
         End While

         Return True
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure decompresses VLE Data inside the specified source and writes the result to the specified target.
   Private Function VLEDecompress(ByRef Source As DataStr, ByRef Target As DataStr) As Boolean
      Try
         Dim Alphabet(0 To VLE_ALPHABET_LENGTH) As Byte
         Dim AlphabetLength As Integer = 0
         Dim CodesOffset As Integer = 0
         Dim EscapeCharacters1(0 To VLE_ESCAPE_CHARACTERS_LENGTH) As Integer
         Dim EscapeCharacters2(0 To VLE_ESCAPE_CHARACTERS_LENGTH) As Integer
         Dim Success As Boolean = False
         Dim Symbols(0 To VLE_ALPHABET_LENGTH) As Byte
         Dim Widths(0 To VLE_ALPHABET_LENGTH) As Byte
         Dim WidthsLengths As Byte = Source.Data(Source.Position)
         Dim WidthsOffset As Integer = 0

         Source.Position += 1
         WidthsOffset = Source.Position

         If Not ((WidthsLengths And VLE_UNKNOWN_WIDTH_LENGTH) = VLE_UNKNOWN_WIDTH_LENGTH) OrElse ((WidthsLengths And VLE_WIDTH_LENGTH_MASK) > VLE_WIDTH_MAXIMUM_LENGTH) Then
            AlphabetLength = VLEGenerateEscapeTable(Source, EscapeCharacters1, EscapeCharacters2, WidthsLengths)
            If AlphabetLength <= VLE_ALPHABET_LENGTH Then
               For Letter As Integer = 0 To AlphabetLength - 1
                  Alphabet(Letter) = Source.Data(Source.Position)
                  Source.Position += 1
               Next Letter

               CodesOffset = Source.Position
               Source.Position = WidthsOffset

               VLEGenerateLookupTable(Source, WidthsLengths, Alphabet, Symbols, Widths)

               Source.Position = CodesOffset

               Success = VLEDecode(Source, Target, Alphabet, Symbols, Widths, EscapeCharacters1, EscapeCharacters2)
            End If
         End If

         Return Success
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure generates the VLE escape table and returns the VLE alphabet length.
   Private Function VLEGenerateEscapeTable(ByRef Source As DataStr, ByRef EscapeCharacters1() As Integer, ByRef EscapeCharacters2() As Integer, WidthsLength As Integer) As Integer
      Try
         Dim AlphabetLength As Integer = 0
         Dim CurrentByte As Byte
         Dim WidthSum As Integer = 0

         For EscapeCharacter As Integer = 0 To WidthsLength - 1
            WidthSum *= 2
            EscapeCharacters1(EscapeCharacter) = AlphabetLength - WidthSum
            CurrentByte = Source.Data(Source.Position)
            Source.Position += 1
            WidthSum += CurrentByte
            AlphabetLength += CurrentByte
            EscapeCharacters2(EscapeCharacter) = WidthSum
         Next EscapeCharacter

         Return AlphabetLength
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure generates the VLE look up table.
   Private Sub VLEGenerateLookupTable(ByRef Source As DataStr, WidthsLengths As Integer, ByRef Alphabet() As Byte, ByRef Symbols() As Byte, ByRef Widths() As Byte)
      Try
         Dim AlphabetIndex As Integer = 0
         Dim SymbolWidthIndex As Integer = 0
         Dim SymbolsPerWidth As Byte = VLE_BYTE_MSB_MASK
         Dim WidthsDistributionLengths As Integer = If(WidthsLengths >= 8, 8, WidthsLengths)

         For Width As Integer = 1 To WidthsDistributionLengths
            For SymbolWidth As Integer = Source.Data(Source.Position) To 1 Step -1
               For SymbolsRemaining As Integer = SymbolsPerWidth To 1 Step -1
                  Symbols(SymbolWidthIndex) = Alphabet(AlphabetIndex)
                  Widths(SymbolWidthIndex) = CByte(Width)
                  SymbolWidthIndex += 1
               Next SymbolsRemaining
               AlphabetIndex += 1
            Next SymbolWidth
            Source.Position += 1
            SymbolsPerWidth >>= &H1%
         Next Width

         For RemainingWidthsIndex As Integer = SymbolWidthIndex To VLE_ALPHABET_LENGTH - 1
            Widths(RemainingWidthsIndex) = VLE_ESCAPE_WIDTH
         Next RemainingWidthsIndex
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure writes the specified data to the specified file.
   Private Function WriteDecompressedFile(TargetFile As String, Target As DataStr) As Boolean
      Try
         File.WriteAllBytes(TargetFile, Target.Data)

         Return True
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return False
   End Function
End Module