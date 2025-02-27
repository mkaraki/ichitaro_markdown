option explicit

declare variable c
declare variable i
declare variable %isBold
declare variable %isItalic
declare variable %lastLine
declare variable %paragEnd(3)
declare variable %paragLength
declare variable %isBoldNext
declare variable %isItalicNext
declare variable %markdownFile
declare variable %markdownHeadingStartLevel
declare variable %cursorChar

%markdownHeadingStartLevel = 2

!! Set Edit type Screen to Image Edit mode
EditingScreenType(1)

!! If cursor not in edit area, error
if GetCursorArea() <> 1 then
	Message("Cursor not in edit area")
	stop
end if

!! Open Markdown file with SharedenyWrite (No write share), Text mode, Unicode (UTF-16?)
set %markdownFile = CreateFile("TestMarkdown.md", &H0010, 1, 5)

!! Set cursor Start of Document
JumpStart()

%lastLine = 0

do
	!! If get end of content. Break this.
	if GetRow() = %lastLine then
		exit do
	end if
	
	EndOfParagraph()
	%paragEnd(1) = GetPage()
	%paragEnd(2) = GetRow()
	%paragEnd(3) = GetColumn()
	StartOfParagraph()
	
	!! ToDo: Get outline level
	
	select case GetParagraphStyleName()
		case "文書ﾀｲﾄﾙ(ｵｰﾄｽﾀｲﾙ)"
			if %markdownHeadingStartLevel >= 2 then
				for i = 1 to %markdownHeadingStartLevel - 1
					%markdownFile.Write("#")
				next
				%markdownFile.Write(" ")
			end if
		case "大見出し(ｵｰﾄｽﾀｲﾙ)"
			for i = 1 to %markdownHeadingStartLevel - 1 + 1
				%markdownFile.Write("#")
			next
			%markdownFile.Write(" ")
		case "中見出し(ｵｰﾄｽﾀｲﾙ)"
			for i = 1 to %markdownHeadingStartLevel - 1 + 2
				%markdownFile.Write("#")
			next
			%markdownFile.Write(" ")
		case "小見出し(ｵｰﾄｽﾀｲﾙ)"
			for i = 1 to %markdownHeadingStartLevel - 1 + 3
				%markdownFile.Write("#")
			next
			%markdownFile.Write(" ")
		case "小見出し2(ｵｰﾄｽﾀｲﾙ)"
			for i = 1 to %markdownHeadingStartLevel - 1 + 4
				%markdownFile.Write("#")
			next
			%markdownFile.Write(" ")
		case "小見出し3(ｵｰﾄｽﾀｲﾙ)"
			for i = 1 to %markdownHeadingStartLevel - 1 + 5
				%markdownFile.Write("#")
			next
			%markdownFile.Write(" ")
		case "小見出し4(ｵｰﾄｽﾀｲﾙ)"
			for i = 1 to %markdownHeadingStartLevel - 1 + 5
				%markdownFile.Write("#")
			next
			%markdownFile.Write(" ")
		case "小見出し5(ｵｰﾄｽﾀｲﾙ)"
			for i = 1 to %markdownHeadingStartLevel - 1 + 6
				%markdownFile.Write("#")
			next
			%markdownFile.Write(" ")
		case "箇条書き(ｵｰﾄｽﾀｲﾙ)"
			%markdownFile.Write("- ")
		case "箇条書きの説明(ｵｰﾄｽﾀｲﾙ)"
			%markdownFile.Write("  ")
	end select
			

	%isBold = false
	%isItalic = false

	do
		%cursorChar = GetCharacter()
	
		%isBoldNext = IsBold()
		
		if %isBold <> %isBoldNext then
			!! Previous char and current char's bold status changed.
			!!  => Switch bold by **
			%markdownFile.Write("**")
		end if
		
		%isBold = %isBoldNext
		
		%isItalicNext = IsSlant()
		
		if %isItalic <> %isItalicNext then
			%markdownFile.Write("*")
		end if
		
		%isItalic = %isItalicNext
		
		%markdownFile.Write(%cursorChar)
		
		if GetPage() = %paragEnd(1) and GetRow() = %paragEnd(2) and GetColumn() = %paragEnd(3) then
			exit do
		else
			CursorRight()
		end if
	loop
	
	%markdownFile.Write(Char(10, 5))
	%markdownFile.Write(Char(10, 5))
	
	!! Next paragraph
	NextParagraph(1)
	
	if GetPage() = %paragEnd(1) and GetRow() = %paragEnd(2) and GetColumn() = %paragEnd(3) then
		!! Nothing new.
		Message("Ended")
		exit do
	end if
loop

%markdownFile.Close()

end
