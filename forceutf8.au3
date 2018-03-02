#include-once

#cs
   ForceUTF8
   http://github.com/jesobreira/forceutf8_au3

   Based on the work by Sebastian Grignoli (https://github.com/neitanod/forceutf8)
   Ported to AutoIt3 by Jefrey S. Santos <jefrey[at]jefrey.ml>
#ce

#Region Constants
Global $win1252ToUtf8 = ObjCreate("Scripting.Dictionary")
$win1252ToUtf8.Add(128, 0xe2 & 0x82 & 0xac)
$win1252ToUtf8.Add(130, 0xe2 & 0x80 & 0x9a)
$win1252ToUtf8.Add(131, 0xc6 & 0x92)
$win1252ToUtf8.Add(132, 0xe2 & 0x80 & 0x9e)
$win1252ToUtf8.Add(133, 0xe2 & 0x80 & 0xa6)
$win1252ToUtf8.Add(134, 0xe2 & 0x80 & 0xa0)
$win1252ToUtf8.Add(135, 0xe2 & 0x80 & 0xa1)
$win1252ToUtf8.Add(136, 0xcb & 0x86)
$win1252ToUtf8.Add(137, 0xe2 & 0x80 & 0xb0)
$win1252ToUtf8.Add(138, 0xc5 & 0xa0)
$win1252ToUtf8.Add(139, 0xe2 & 0x80 & 0xb9)
$win1252ToUtf8.Add(140, 0xc5 & 0x92)
$win1252ToUtf8.Add(142, 0xc5 & 0xbd)
$win1252ToUtf8.Add(145, 0xe2 & 0x80 & 0x98)
$win1252ToUtf8.Add(146, 0xe2 & 0x80 & 0x99)
$win1252ToUtf8.Add(147, 0xe2 & 0x80 & 0x9c)
$win1252ToUtf8.Add(148, 0xe2 & 0x80 & 0x9d)
$win1252ToUtf8.Add(149, 0xe2 & 0x80 & 0xa2)
$win1252ToUtf8.Add(150, 0xe2 & 0x80 & 0x93)
$win1252ToUtf8.Add(151, 0xe2 & 0x80 & 0x94)
$win1252ToUtf8.Add(152, 0xcb & 0x9c)
$win1252ToUtf8.Add(153, 0xe2 & 0x84 & 0xa2)
$win1252ToUtf8.Add(154, 0xc5 & 0xa1)
$win1252ToUtf8.Add(155, 0xe2 & 0x80 & 0xba)
$win1252ToUtf8.Add(156, 0xc5 & 0x93)
$win1252ToUtf8.Add(158, 0xc5 & 0xbe)
$win1252ToUtf8.Add(159, 0xc5 & 0xb8)

$brokenUtf8ToUtf8 = ObjCreate("Scripting.Dictionary")
$brokenUtf8ToUtf8.Add(0xc2 & 0x80, 0xe2 & 0x82 & 0xac)
$brokenUtf8ToUtf8.Add(0xc2 & 0x82, 0xe2 & 0x80 & 0x9a)
$brokenUtf8ToUtf8.Add(0xc2 & 0x83, 0xc6 & 0x92)
$brokenUtf8ToUtf8.Add(0xc2 & 0x84, 0xe2 & 0x80 & 0x9e)
$brokenUtf8ToUtf8.Add(0xc2 & 0x85, 0xe2 & 0x80 & 0xa6)
$brokenUtf8ToUtf8.Add(0xc2 & 0x86, 0xe2 & 0x80 & 0xa0)
$brokenUtf8ToUtf8.Add(0xc2 & 0x87, 0xe2 & 0x80 & 0xa1)
$brokenUtf8ToUtf8.Add(0xc2 & 0x88, 0xcb & 0x86)
$brokenUtf8ToUtf8.Add(0xc2 & 0x89, 0xe2 & 0x80 & 0xb0)
$brokenUtf8ToUtf8.Add(0xc2 & 0x8a, 0xc5 & 0xa0)
$brokenUtf8ToUtf8.Add(0xc2 & 0x8b, 0xe2 & 0x80 & 0xb9)
$brokenUtf8ToUtf8.Add(0xc2 & 0x8c, 0xc5 & 0x92)
$brokenUtf8ToUtf8.Add(0xc2 & 0x8e, 0xc5 & 0xbd)
$brokenUtf8ToUtf8.Add(0xc2 & 0x91, 0xe2 & 0x80 & 0x98)
$brokenUtf8ToUtf8.Add(0xc2 & 0x92, 0xe2 & 0x80 & 0x99)
$brokenUtf8ToUtf8.Add(0xc2 & 0x93, 0xe2 & 0x80 & 0x9c)
$brokenUtf8ToUtf8.Add(0xc2 & 0x94, 0xe2 & 0x80 & 0x9d)
$brokenUtf8ToUtf8.Add(0xc2 & 0x95, 0xe2 & 0x80 & 0xa2)
$brokenUtf8ToUtf8.Add(0xc2 & 0x96, 0xe2 & 0x80 & 0x93)
$brokenUtf8ToUtf8.Add(0xc2 & 0x97, 0xe2 & 0x80 & 0x94)
$brokenUtf8ToUtf8.Add(0xc2 & 0x98, 0xcb & 0x9c)
$brokenUtf8ToUtf8.Add(0xc2 & 0x99, 0xe2 & 0x84 & 0xa2)
$brokenUtf8ToUtf8.Add(0xc2 & 0x9a, 0xc5 & 0xa1)
$brokenUtf8ToUtf8.Add(0xc2 & 0x9b, 0xe2 & 0x80 & 0xba)
$brokenUtf8ToUtf8.Add(0xc2 & 0x9c, 0xc5 & 0x93)
$brokenUtf8ToUtf8.Add(0xc2 & 0x9e, 0xc5 & 0xbe)
$brokenUtf8ToUtf8.Add(0xc2 & 0x9f, 0xc5 & 0xb8)

$utf8ToWin1252 = ObjCreate("Scripting.Dictionary")
$utf8ToWin1252.Add(0xe2 & 0x82 & 0xac, 0x80)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0x9a, 0x82)
$utf8ToWin1252.Add(0xc6 & 0x92, 0x83)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0x9e, 0x84)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0xa6, 0x85)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0xa0, 0x86)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0xa1, 0x87)
$utf8ToWin1252.Add(0xcb & 0x86, 0x88)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0xb0, 0x89)
$utf8ToWin1252.Add(0xc5 & 0xa0, 0x8a)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0xb9, 0x8b)
$utf8ToWin1252.Add(0xc5 & 0x92, 0x8c)
$utf8ToWin1252.Add(0xc5 & 0xbd, 0x8e)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0x98, 0x91)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0x99, 0x92)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0x9c, 0x93)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0x9d, 0x94)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0xa2, 0x95)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0x93, 0x96)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0x94, 0x97)
$utf8ToWin1252.Add(0xcb & 0x9c, 0x98)
$utf8ToWin1252.Add(0xe2 & 0x84 & 0xa2, 0x99)
$utf8ToWin1252.Add(0xc5 & 0xa1, 0x9a)
$utf8ToWin1252.Add(0xe2 & 0x80 & 0xba, 0x9b)
$utf8ToWin1252.Add(0xc5 & 0x93, 0x9c)
$utf8ToWin1252.Add(0xc5 & 0xbe, 0x9e)
$utf8ToWin1252.Add(0xc5 & 0xb8, 0x9f)

Global $__encoding_equivalences = ObjCreate("Scripting.Dictionary")
$__encoding_equivalences.Add("ISO88591", "ISO-8859-1")
$__encoding_equivalences.Add("ISO8859", "ISO-8859-1")
$__encoding_equivalences.Add("ISO", "ISO-8859-1")
$__encoding_equivalences.Add("LATIN1", "ISO-8859-1")
$__encoding_equivalences.Add("LATIN", "ISO-8859-1")
$__encoding_equivalences.Add("UTF8", "UTF-8")
$__encoding_equivalences.Add("UTF", "UTF-8")
$__encoding_equivalences.Add("WIN1252", "ISO-8859-1")
$__encoding_equivalences.Add("WINDOWS1252", "ISO-8859-1")
#EndRegion

Func toUTF8($text)
   If Not IsString($text) Then
	  Return SetError(1, 0, False)
   EndIf

   Local $buf,$max = StringLen($text)
   Local $c1,$c2,$c3,$c4,$cc1,$cc2

   For $i = 1 To $max
	  $c1 = StringMid($text, $i, 1)
	  If $c1 >= 0xc0 Then ; Should be converted to UTF8, if it's not UTF8 already
		 $c2 = $i+1 > $max ? 0x00 : StringMid($text, $i+1, 1)
		 $c3 = $i+2 > $max ? 0x00 : StringMid($text, $i+2, 1)
		 $c4 = $i+3 > $max ? 0x00 : StringMid($text, $i+3, 1)
		 If $c1 >= 0xc0 And $c1 <= 0xdf Then ; looks like 2 bytes UTF8
			If $c2 >= 0x80 And $c2 <= 0xbf Then ; yeah, almost sure it's UTF8 already
			   $buf &= $c1 & $c2
			   $i += 1
			Else ; not valid UTF8. Convert it.
			   $cc1 = BitOR(Chr(Asc($c1) / 64), 0xc0)
			   $cc2 = BitOR(BitAND($c1, 0x3f), 0x80)
			   $buf &= $cc1 & $cc2
			EndIf
		 ElseIf $c1 >= 0xe0 And $c1 <= 0xef Then ; looks like 3 bytes UTF8
			If $c2 >= 0x80 And $c2 <= 0xbf And $c3 >= 0x80 And $c3 <= 0xbf Then ; yeah, almost sure it's UTF8 already
			   $buf &= $c1 & $c2 & $c3
			   $i += 2
			Else ; not valid UTF8. Convert it.
			   $cc1 = BitOR(Chr(Asc($c1) / 64), 0xc0)
			   $cc2 = BitOR(BitAND($c1, 0x3f), 0x80)
			   $buf &= $cc1 & $cc2
			EndIf
		 ElseIf $c1 >= 0xf0 And $c1 <= 0xf7 Then ; looks like 4 bytes UTF8
			If $c2 >= 0x80 And $c2 <= 0xbf And $c3 >= 0x80 And $c3 <= 0xbf And $c4 >= 0x80 And $c4 <= 0xbf Then ; yeah, almost sure it's UTF8 already
			   $buf &= $c1 & $c2 & $c3 & $c4
			   $i += 3
			Else ; not valid UTF8. Convert it.
			   $cc1 = BitOR(Chr(Asc($c1) / 64), 0xc0)
			   $cc2 = BitOR(BitAND($c1, 0x3f), 0x80)
			   $buf &= $cc1 & $cc2
			EndIf
		 Else ; doesn't look like UTF8, but should be converted
			$cc1 = BitOR(Chr(Asc($c1) / 64), 0xc0)
			$cc2 = BitOR(BitAND($c1, 0x3f), 0x80)
			$buf &= $cc1 & $cc2
		 EndIf
	  ElseIf BitAND($c1, 0xc0) = 0x80 Then ; needs conversion
		 If $win1252ToUtf8.Exists(Asc($c1)) Then
			$buf &= $win1252ToUtf8.Item(Asc($c1))
		 Else
			$cc1 = BitOR(Chr(Asc($c1) / 64), 0xc0)
			$cc2 = BitOR(BitAND($c1, 0x3f), 0x80)
			$buf &= $cc1 & $cc2
		 EndIf
	  Else ; it doesn't need conversion
		 $buf &= $c1
	  EndIf
   Next
   Return $buf
EndFunc

Func toWin1252($text)
   Return utf8_decode($text)
EndFunc

Func toISO8859($text)
   Return toWin1252($text)
EndFunc

Func toLatin1($text)
   Return toWin1252($text)
EndFunc

Func fixUTF8($text)
   Local $last
   While $Last <> $text
	  $last = $text
	  $text = toUTF8($text)
   WEnd
   $text = toUTF8($text)
   Return $text
EndFunc

Func UTF8FixWin1252Chars($text)
   ; If you received an UTF-8 string that was converted from Windows-1252 as it was ISO8859-1
   ; (ignoring Windows-1252 chars from 80 to 9F) use this function to fix it.
   ; See: http://en.wikipedia.org/wiki/Windows-1252

   Local $value,$o
   $keys = $brokenUtf8ToUtf8.Keys
   $items = $brokenUtf8ToUtf8.Items
   For $key In $keys
	  $value = $brokenUtf8ToUtf8.Item($key)
	  $o = StringReplace($text, $key, $value)
   Next
   Return $o
EndFunc

Func removeBOM($str)
   If StringLeft($str, 3) = 0xef & 0xbb & 0xbf Then
	  $str = StringTrimLeft($str, 3)
   EndIf
   Return $str
EndFunc

Func normalizeEncoding($encodingLabel)
   $encoding = StringUpper($encodingLabel)
   $encoding = StringRegExpReplace($encoding, "[^a-zA-Z0-9\s]", "")
   If $__encoding_equivalences.Exists($encoding) Then
	  Return $__encoding_equivalences.Item($encoding)
   Else
	  Return 'UTF-8'
   EndIf
EndFunc

Func encode($encodingLabel, $text)
   $encodingLabel = normalizeEncoding($encodingLabel)
   If $encodingLabel = "ISO-8859-1" Then Return toLatin1($text)
   Return toUTF8($text)
EndFunc

Func utf8_decode($text)
   Local $value,$o
   $keys = $utf8ToWin1252.Keys
   $items = $utf8ToWin1252.Items
   For $key In $keys
	  $value = $utf8ToWin1252.Item($key)
	  $o = StringReplace(toUTF8($text), $key, $value)
   Next
   Return $o
EndFunc
