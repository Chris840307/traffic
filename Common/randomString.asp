<%
Function gen_key(digits) 
	dim char_array(80) 
	For i = 0 To 9 
		char_array(i) = CStr(i) 
	Next 
	For i = 10 To 35 
		char_array(i) = Chr(i + 55) 
	Next 
	For i = 36 To 61 
		char_array(i) = Chr(i + 61) 
	Next 

	Randomize 

	do while len(output) < digits 
		num = char_array(Int((62 - 0 + 1) * Rnd + 0)) 
		output = output & num 
	loop 

	gen_key = output 

End Function 
%>