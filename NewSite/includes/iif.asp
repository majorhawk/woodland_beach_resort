<%
function iif(trueFalse, a, b)
	if cbool(trueFalse) then
		iif = a
	else
		iif = b
	end if
end function 


dim Random, smPhoto
Function RandomNumber(intHighestNumber)
	Randomize
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function

Random = RandomNumber(8)

smPhoto = "albm_photo_0" & Random
%>