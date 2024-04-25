Imports System

Module Program
    Sub Main()
    End Sub
    Sub Paper2023()
        Dim IsValid As Boolean = False
        While Not IsValid
            Console.Write("Enter a valid string: ")
            Dim Input As String = Console.ReadLine()
            If Input.Length > 7 OrElse Input.Length < 5 Then
                Console.WriteLine("Error: string is not between 5 and 7 letters long (inclusive).")
                Continue While
            End If
            Dim Chars As New List(Of String)
            Dim Sum As Integer = 0
            For Each letter In Input
                If Asc(letter) < 65 OrElse Asc(letter) > 90 Then
                    Console.WriteLine("Error: string contains characters other than capital letters.")
                    Continue While
                ElseIf Chars.Contains(letter) Then
                    Console.WriteLine("Error: string contains duplicate letters.")
                    Continue While
                Else
                    Chars.Add(letter)
                    Sum += Asc(letter)
                End If
            Next
            If Sum > 600 OrElse Sum < 420 Then
                Console.WriteLine("Error: string sum of ASCII values is not between 420 and 600 inclusive.")
                Continue While
            End If
            IsValid = True
            Console.WriteLine("Valid string entered.")
        End While
    End Sub
    Function Paper2022(Input As String) As String
        Dim Vowels As String() = New String() {"A", "E", "I", "O", "U"}
        Dim InputStr As String = Input.ToUpper
        Dim Result(Input.Length - 1) As Char
        Dim VowelDict As New List(Of Tuple(Of Integer, String))() 'first value is index, second is actual vowel

        For i As Integer = 0 To Input.Length - 1
            If Vowels.Contains(InputStr(i)) Then
                VowelDict.Add(New Tuple(Of Integer, String)(i, Input(i)))
            Else
                Result(i) = Input(i)
            End If
        Next
        'have a full dict of vowels and their indexes now
        If VowelDict.Count > 1 Then 'dont need to flip any vowels if less than 2 vowels
            For i As Integer = 0 To (VowelDict.Count - 1) \ 2
                Dim TempVowel As String = VowelDict(i).Item2
                VowelDict(i) = New Tuple(Of Integer, String)(VowelDict(i).Item1, VowelDict(VowelDict.Count - i - 1).Item2)
                VowelDict(VowelDict.Count - 1 - i) = New Tuple(Of Integer, String)(VowelDict(VowelDict.Count - 1 - i).Item1, TempVowel)
            Next
            'swapping each vowel with its opposite "pair" vowel, hence for loop only going up to the floor of the midpoint of the list, as odd-numbered centre vowels dont switch positions
        End If

        If VowelDict.Count > 0 Then
            For i As Integer = 0 To VowelDict.Count - 1
                Result(VowelDict(i).Item1) = VowelDict(i).Item2
            Next
        End If

        Return New String(Result)
    End Function
    Function Paper2021(N As Integer) As Integer
        Dim Index As Integer = 0
        Dim Number As Integer = 0
        While Index < N
            Number += 1
            Dim DigitSum As Integer = 0
            Dim NumberStr As String = Number.ToString()
            For i As Integer = 0 To NumberStr.Length - 1
                DigitSum += CInt(NumberStr(i).ToString())
            Next
            If Number Mod DigitSum = 0 Then
                Index += 1
            End If
        End While
        Return Number
    End Function
    Sub Paper2020(NoOfInputs As Integer)
        Dim Inputs(NoOfInputs - 1) As Integer
        Dim OccurrenceDict As New Dictionary(Of Integer, Integer)
        For i As Integer = 0 To NoOfInputs - 1
            Console.Write($"Enter digit number {i + 1}: ")
            Inputs(i) = CInt(Console.ReadLine())
            If OccurrenceDict.ContainsKey(Inputs(i)) Then
                OccurrenceDict(Inputs(i)) += 1
            Else
                OccurrenceDict.Add(Inputs(i), 1)
            End If
        Next

        Dim MaxOccurrences As Integer = -1
        Dim MultiModal As Boolean = False
        Dim Mode As Integer = -1

        For Each kvPair In OccurrenceDict
            If kvPair.Value > MaxOccurrences Then
                MaxOccurrences = kvPair.Value
                Mode = kvPair.Key
                MultiModal = False
            ElseIf kvPair.Value = MaxOccurrences Then
                MultiModal = True
            End If
        Next

        For i As Integer = 0 To OccurrenceDict.Count - 1

        Next

        If MultiModal Then
            Console.WriteLine("Data was multimodal.")
        Else
            Console.WriteLine($"Mode: {Mode}, Frequency: {MaxOccurrences}")
        End If
    End Sub
    Function Paper2019(NewWord As String, Original As String) As Boolean
        Dim AbleTo As Boolean = True
        Dim Parts As List(Of Char) = NewWord.ToList()
        Dim Whole As List(Of Char) = Original.ToList()
        If Parts.Count > Whole.Count Then
            AbleTo = False
        End If
        While AbleTo And Parts.Count > 0
            Dim TempChar As Char = Parts(0)
            If Whole.Contains(TempChar) Then
                Whole.Remove(TempChar)
                Parts.RemoveAt(0)
            Else
                AbleTo = False
            End If
        End While
        Return AbleTo
    End Function
    Function Paper2018(N As Integer) As Boolean
        If N < 2 Then
            Console.WriteLine("Not greater than 1.")
            Return False
        End If
        If N = 2 Then
            Console.WriteLine("Is a prime.")
            Return True
        End If
        For i As Integer = 2 To Math.Min(Math.Sqrt(N) + 1, N - 1)
            If N Mod i = 0 Then
                Console.WriteLine("Is not a prime.")
                Return False
            End If
        Next
        Console.WriteLine("Is a prime.")
        Return True
    End Function
    Function Paper2017(Data As String) As String
        Dim Result As New List(Of String)
        If Data = "" Then
            Return ""
        End If
        For I As Integer = 1 To Data.Length 'looping through a valid input string
            Dim LetterCount As Integer = 1
            While Mid(Data, I, 1) = Mid(Data, I + 1, 1)
                LetterCount += 1
                I += 1
            End While
            If Asc(Mid(Data, I, 1)) = 32 Then
                Result.Add(""" """ & CStr(LetterCount))
            ElseIf Asc(Mid(Data, I, 1)) >= 48 Or Asc(Mid(Data, I, 1)) <= 57 Then
                Result.Add(Mid(Data, I, 1) & CStr(LetterCount))
            Else
                Result.Add(Mid(Data, I, 1).ToUpper & CStr(LetterCount))
            End If
        Next
        Return String.Join(", ", Result)
    End Function
End Module
