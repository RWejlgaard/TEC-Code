globalAnswer = vbyes
Do Until globalAnswer = vbno
	Set list = CreateObject("Scripting.FileSystemObject")
	Set file = list.OpenTextFile("C:\Users\HVF-E308\Desktop\test.txt",1)
	count = 0
	ReDim arrList(0)
	Do Until file.AtEndOfStream
	    nxtLine = UCase(file.ReadLine)
	    arrList(count) = nxtLine
	    count = count + 1
	    ReDim Preserve arrList(UBound(arrList)+1)
	Loop
	word = arrList(rand(0,UBound(arrList)))
	ReDim ltrRevealed(Len(word))
	ReDim ltrHidden(Len(word))
	
	used = ""
	score = 0
	combo = 0
	lives = 5
	isSolved = False
	setDefaultValues()
	
	do while isSolved = False
		temp = ""
		For Each x In ltrHidden
			temp = (temp & " " & x)
		next
		answer = InputBox("Score: " & score & vbNewLine & vbNewLine & temp & vbNewLine & vbNewLine & "Lives: " & lives & vbNewLine & vbNewLine & "Used Words: " & used, "HANGMAN")
		isInWord(answer)
		used = used & " " & UCase(answer)
		If lives < 0 Then
			MsgBox("YOU ARE DEAD" & vbNewLine & vbNewLine & "The word was: " & word)
			isSolved = true			
		End if
		If isDone = true Then
			MsgBox("Congratulations!!!!!" & vbNewLine & vbNewLine & "The word was: " & word & vbNewLine & "Final Score: " & Score)
			isSolved = true
		End if
	Loop
	
	globalAnswer = MsgBox("Run Again?", vbYesNo)

Loop

Function isInWord (input)
		If answer = "quit" Then 
		WScript.Quit
	End if
	for i=1 to len(word)
		if(mid(word,i,1) = UCase(input)) Then
			ltrHidden(i) = ltrRevealed(i)
			combo = combo + 1
			score = score + 1 * combo
		Else If (InStr(word, UCase(answer)) = False) then
			lives = lives - 1
			combo = 0
			Exit For 
		end If
		End if
	next
End Function

Function setDefaultValues ()
	For i=1 to len(word)
		ltrRevealed(i) = Mid(word,i,1)
	Next
	
	for i=1 to len(word)
		ltrHidden(i) = "_"
	Next
End Function

Function isDone ()	
	For Each x In ltrHidden
		If x = "_" Then
			Exit function
		End If
	Next
	isDone = True
End Function

Function rand(min, max)
    Randomize
    Rand = (Int((max-min+1)*Rnd+min))
End Function

