Attribute VB_Name = "UniConvertModule"
Function UniConvert(Text As String, InputMethod As String) As String
  Dim VNI_Type, Telex_Type, CharCode, Temp, i As Long
  UniConvert = Text
  VNI_Type = Array("a81", "a82", "a83", "a84", "a85", "a61", "a62", "a63", "a64", "a65", "e61", _
      "e62", "e63", "e64", "e65", "o61", "o62", "o63", "o64", "o65", "o71", "o72", "o73", "o74", _
      "o75", "u71", "u72", "u73", "u74", "u75", "a1", "a2", "a3", "a4", "a5", "a8", "a6", "d9", _
      "e1", "e2", "e3", "e4", "e5", "e6", "i1", "i2", "i3", "i4", "i5", "o1", "o2", "o3", "o4", _
      "o5", "o6", "o7", "u1", "u2", "u3", "u4", "u5", "u7", "y1", "y2", "y3", "y4", "y5")
  Telex_Type = Array("aws", "awf", "awr", "awx", "awj", "aas", "aaf", "aar", "aax", "aaj", "ees", _
      "eef", "eer", "eex", "eej", "oos", "oof", "oor", "oox", "ooj", "ows", "owf", "owr", "owx", _
      "owj", "uws", "uwf", "uwr", "uwx", "uwj", "as", "af", "ar", "ax", "aj", "aw", "aa", "dd", _
      "es", "ef", "er", "ex", "ej", "ee", "is", "if", "ir", "ix", "ij", "os", "of", "or", "ox", _
      "oj", "oo", "ow", "us", "uf", "ur", "ux", "uj", "uw", "ys", "yf", "yr", "yx", "yj")
  CharCode = Array(ChrW(7855), ChrW(7857), ChrW(7859), ChrW(7861), ChrW(7863), ChrW(7845), ChrW(7847), _
      ChrW(7849), ChrW(7851), ChrW(7853), ChrW(7871), ChrW(7873), ChrW(7875), ChrW(7877), ChrW(7879), _
      ChrW(7889), ChrW(7891), ChrW(7893), ChrW(7895), ChrW(7897), ChrW(7899), ChrW(7901), ChrW(7903), _
      ChrW(7905), ChrW(7907), ChrW(7913), ChrW(7915), ChrW(7917), ChrW(7919), ChrW(7921), ChrW(225), _
      ChrW(224), ChrW(7843), ChrW(227), ChrW(7841), ChrW(259), ChrW(226), ChrW(273), ChrW(233), ChrW(232), _
      ChrW(7867), ChrW(7869), ChrW(7865), ChrW(234), ChrW(237), ChrW(236), ChrW(7881), ChrW(297), ChrW(7883), _
      ChrW(243), ChrW(242), ChrW(7887), ChrW(245), ChrW(7885), ChrW(244), ChrW(417), ChrW(250), ChrW(249), _
      ChrW(7911), ChrW(361), ChrW(7909), ChrW(432), ChrW(253), ChrW(7923), ChrW(7927), ChrW(7929), ChrW(7925))
  Select Case InputMethod
    Case Is = "VNI": Temp = VNI_Type
    Case Is = "Telex": Temp = Telex_Type
  End Select
  For i = 0 To UBound(CharCode)
    UniConvert = Replace(UniConvert, Temp(i), CharCode(i))
    UniConvert = Replace(UniConvert, UCase(Temp(i)), UCase(CharCode(i)))
  Next i
End Function

