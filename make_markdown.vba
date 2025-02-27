option explicit

declare variable c
declare variable %isBold
declare variable %isItalic
declare variable %lastLine
declare variable %lineContent
declare variable %lineLength
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
	
	if GetPage() > 2 then
		Message("Page limit exceeded")
		exit do
	end if
	
	EndOfLine()
	%lineLength = GetColumn()
	StartOfLine()
	
	!! ToDo: Get outline level

	%isBold = false
	%isItalic = false
	
	if GetRow() = 1 then
		Message(%lineLength)
	end if
	
	for c = 1 to %lineLength step 1
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
		
		CursorRight()
	next
	
	%markdownFile.Write(Char(10, 5))
	
	!! Next line (line++)
	!! Last cursor right cause line skip. So, we don't need it.
	!! NextLine(1)
loop

%markdownFile.Close()

end
