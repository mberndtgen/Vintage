' taken from http://itknowledgeexchange.techtarget.com/vbscript-systems-administrator/very-simple-encryption-example-with-vbscript/

temp = encrypt("was#CADMGR", "huasHIYhkasdho1")
wscript.echo temp
temp = Decrypt("º¼®·œ |ÛÌØ", "huasHIYhkasdho1")
wscript.echo temp

Function encrypt(Str, key)
 Dim lenKey, KeyPos, LenStr, x, Newstr, EncCharNum

 Newstr = ""
 lenKey = Len(key)
 KeyPos = 1
 LenStr = Len(Str)
 str = StrReverse(str)
 For x = 1 To LenStr
      EncCharNum = Asc (Mid (str, x, 1)) + Asc (Mid (key, KeyPos, 1))
      Newstr = Newstr & chr(EncCharNum Mod 256)
      KeyPos = keypos+1
      If KeyPos > lenKey Then KeyPos = 1
 Next
 encrypt = Newstr
End Function

Function Decrypt(str,key)
 Dim lenKey, KeyPos, LenStr, x, Newstr, DecCharNum

 Newstr = ""
 lenKey = Len(key)
 KeyPos = 1
 LenStr = Len(Str)

 str=StrReverse(str)
 For x = LenStr To 1 Step -1
      DecCharNum = Asc (Mid (str, x, 1)) - Asc (Mid (key,KeyPos, 1)) + 256
      Newstr = Newstr & chr(DecCharNum Mod 256)
      KeyPos = KeyPos+1
      If KeyPos > lenKey Then KeyPos = 1
      Next
      Newstr=StrReverse(Newstr)
      Decrypt = Newstr
End Function