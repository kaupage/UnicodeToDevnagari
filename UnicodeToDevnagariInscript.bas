Sub UnicodeToDevnagariInscript()
    'converts Unicode Characters to Devnagari Inscript
    'Use in conjuction with Shree - Unicode version 2 Excel
    
    Set RngTxt = Selection.Range
    Set RngFnd = RngTxt.Duplicate
    
    Dim splChr_Devnagari, splChr_Uni, plusMinus_Devnagari, plusMinus_Uni, rKreem_Devnagari, rKreem_Uni, rKree_Devnagari, rKree_Uni, rKrim_Devnagari, rKrim_Uni, rKri_Devnagari, rKri_Uni As Variant
    Dim rKaim_Devnagari, rKaim_Uni, rKai_Devnagari, rKai_Uni, rKem_Devnagari, rKem_Uni, rKe_Devnagari, rKe_Uni As Variant
    Dim rKam_Devnagari, rKam_Uni, rKa_Devnagari, rKa_Uni, rKm_Devnagari, rKm_Uni, rK_Devnagari, rK_Uni, k_Devnagari, k_Uni As Variant
    
    
        '094D position change
    Dim buff() As String
    ReDim buff(Len(RngFnd) - 1)
    For i = 1 To Len(RngFnd)
        buff(i - 1) = Mid$(RngFnd, i, 1)
    Next
    
    '2381 velanti replace
    For i = LBound(buff) To UBound(buff)
    
    If (AscW(buff(i)) = 2367) Then
    If (i - 2 > 0) Then
        If (AscW(buff(i - 2)) = 2381) Then
        'jodakshar aahe
            If (i - 4 > 0 & AscW(buff(i - 4)) = 2381) Then
            'double jodakshar
            'akshar | 094D | akshar | 094D | akshar | velanti
            Dim zemp1 As String
            Dim zemp2 As String
            Dim zemp3 As String
            Dim zemp4 As String
            Dim zemp5 As String
            Dim zemp6 As String
            zemp6 = buff(i)
            zemp5 = buff(i - 1)
            zemp4 = buff(i - 2)
            zemp3 = buff(i - 3)
            zemp2 = buff(i - 4)
            zemp1 = buff(i - 5)
            buff(i) = zemp5
            buff(i - 1) = zemp4
            buff(i - 2) = zemp3
            buff(i - 3) = zemp2
            buff(i - 4) = zemp1
            buff(i - 5) = zemp6
            
            Else
            'single dokakshar
            'akshar | 094D | akshar | velanti
            Dim temp1 As String
            Dim temp2 As String
            Dim temp3 As String
            Dim temp4 As String
            temp4 = buff(i)
            temp3 = buff(i - 1)
            temp2 = buff(i - 2)
            temp1 = buff(i - 3)
            buff(i) = temp3
            buff(i - 1) = temp2
            buff(i - 2) = temp1
            buff(i - 3) = temp4
            End If
        Else
        'jodakshar naahi
        Dim temp5 As String
        temp5 = buff(i - 1)
        buff(i - 1) = buff(i)
        buff(i) = temp5
        End If
    Else
    Dim temp As String 'Pahilya aksharala velanti
    temp = buff(i - 1)
    buff(i - 1) = buff(i)
    buff(i) = temp
    End If
    End If
    Next
    
    RngFnd = Join(buff, "")
    
    'ü   û   f   Ÿ   ú   Äæ  ÂÂà õ   ù   ÷   ö   ¨¾à X   V   U   ¹Ü  T   S   ¹ì  ®îÈà    ®î«à    ®îªà    ®î©à    ®î§ ññ  ò   ï¼  Øàî ØZààè   ØZààç   ØZàÞ    ØZàÝ
    splChr_Devnagari = Array(ChrW(8250), ChrW(8482), ChrW(200), ChrW(201), ChrW(202), ChrW(203), ChrW(204), ChrW(205), ChrW(206), ChrW(207), ChrW(208), ChrW(209), ChrW(210), ChrW(211), ChrW(212), ChrW(213), ChrW(214), ChrW(215), ChrW(216), ChrW(217), ChrW(218), ChrW(219), ChrW(220), ChrW(221), ChrW(222), ChrW(223), ChrW(224), ChrW(225), ChrW(226), ChrW(227), ChrW(228), ChrW(229), ChrW(39), ChrW(34), ChrW(34), ChrW(39))
    splChr_Uni = Array(ChrW(2404), ChrW(2365), ChrW(2406), ChrW(2407), ChrW(2408), ChrW(2409), ChrW(2410), ChrW(2411), ChrW(2412), ChrW(2413), ChrW(2414), ChrW(2415), ChrW(48), ChrW(49), ChrW(50), ChrW(51), ChrW(52), ChrW(53), ChrW(54), ChrW(55), ChrW(56), ChrW(57), ChrW(40), ChrW(41), ChrW(42), ChrW(43), ChrW(45), ChrW(739), ChrW(8763), ChrW(47), ChrW(36), ChrW(59), ChrW(8216), ChrW(8221), ChrW(8220), ChrW(8217))
    
    For i = LBound(splChr_Uni) To UBound(splChr_Uni)
        RngFnd = Replace(RngFnd, splChr_Uni(i), splChr_Devnagari(i))
    Next
    
    '+   =   )   and numbers 1-9
    plusMinus_Devnagari = Array(ChrW(112), ChrW(48), ChrW(49), ChrW(50), ChrW(51), ChrW(52), ChrW(53), ChrW(54), ChrW(55), ChrW(56), ChrW(57), ChrW(48), ChrW(49), ChrW(50), ChrW(51), ChrW(52), ChrW(53), ChrW(54), ChrW(55), ChrW(56), ChrW(57))
    plusMinus_Uni = Array(ChrW(2365), ChrW(2406), ChrW(2407), ChrW(2408), ChrW(2409), ChrW(2410), ChrW(2411), ChrW(2412), ChrW(2413), ChrW(2414), ChrW(2415), ChrW(48), ChrW(49), ChrW(50), ChrW(51), ChrW(52), ChrW(53), ChrW(54), ChrW(55), ChrW(56), ChrW(57))
    
    For i = LBound(plusMinus_Uni) To UBound(plusMinus_Uni)
        RngFnd = Replace(RngFnd, plusMinus_Uni(i), plusMinus_Devnagari(i))
    Next
    
        '§ýK ©àK ªàK iK  ®K  °àK ²K  ³àK ´àK ¸àK ¹K  ºK  "K  ¼K  ½àK ¾àK ¿àK ÀK  ÁàK ÂàK ïÂàK    qK  ÄK  ÆàK ÇàK ÈàK ÉàK ÊK  ïÊK vK  ÍàK ÎàK ÏàK ÐàK ÑK  ÓK  ïÓK ÕàK ÖàK
    rKreem_Devnagari = Array( _
    ChrW(167) & ChrW(253) & ChrW(75), ChrW(169) & ChrW(224) & ChrW(75), ChrW(170) & ChrW(224) & ChrW(75), ChrW(105) & ChrW(75), ChrW(174) & ChrW(75), ChrW(176) & ChrW(224) & ChrW(75), ChrW(178) & ChrW(75), ChrW(179) & ChrW(224) & ChrW(75), ChrW(180) & ChrW(224) & ChrW(75), ChrW(184) & ChrW(224) & ChrW(75), ChrW(185) & ChrW(75), ChrW(186) & ChrW(75), ChrW(187) & ChrW(75), ChrW(188) & ChrW(75), ChrW(189) & ChrW(224) & ChrW(75), ChrW(190) & ChrW(224) & ChrW(75), ChrW(191) & ChrW(224) & ChrW(75), ChrW(192) & ChrW(75), ChrW(193) & ChrW(224) & ChrW(75), ChrW(194) & ChrW(224) & ChrW(75), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(75), ChrW(113) & ChrW(75), ChrW(196) & ChrW(75), ChrW(198) & ChrW(224) & ChrW(75), ChrW(199) & ChrW(224) & ChrW(75), ChrW(200) & ChrW(224) & ChrW(75), ChrW(201) & ChrW(224) & ChrW(75), ChrW(202) & ChrW(75), ChrW(239) & ChrW(202) & ChrW(75), ChrW(118) & ChrW(75), ChrW(205) & ChrW(224) & ChrW(75), ChrW(206) & ChrW(224) & ChrW(75), ChrW(207) & ChrW(224) & ChrW(75), _
    ChrW(208) & ChrW(224) & ChrW(75), ChrW(209) & ChrW(75), ChrW(211) & ChrW(75), ChrW(239) & ChrW(211) & ChrW(75), ChrW(213) & ChrW(224) & ChrW(75), ChrW(214) & ChrW(224) & ChrW(75))
    
    rKreem_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2368) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2368) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2368) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2368) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(23812359) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(23812334) & ChrW(2368) & ChrW(2306))
    
    For i = LBound(rKreem_Devnagari) To UBound(rKreem_Devnagari)
        RngFnd = Replace(RngFnd, rKreem_Uni(i), rKreem_Devnagari(i))
    Next
    
    '§J  ©àJ ªàJ iJ  ®J  °àJ ²J  ³àJ ´àJ ¸àJ ¹J  ºJ  "J  ¼J  ½àJ ¾àJ ¿àJ ÀJ  ÁàJ ÂàJ ïÂàJ    qJ  ÄJ  ÆàJ ÇàJ ÈàJ ÉàJ ÊJ  ïÊJ vJ  ÍàJ ÎàJ ÏàJ ÐàJ ÑJ  ÓJ  ïÓJ ÕàJ ÖàJ
    rKree_Devnagari = Array( _
    ChrW(167) & ChrW(253) & ChrW(74), ChrW(169) & ChrW(224) & ChrW(74), ChrW(170) & ChrW(224) & ChrW(74), ChrW(105) & ChrW(74), ChrW(174) & ChrW(74), ChrW(176) & ChrW(224) & ChrW(74), ChrW(178) & ChrW(74), ChrW(179) & ChrW(224) & ChrW(74), _
    ChrW(180) & ChrW(224) & ChrW(74), ChrW(184) & ChrW(224) & ChrW(74), ChrW(185) & ChrW(74), ChrW(186) & ChrW(74), ChrW(187) & ChrW(74), ChrW(188) & ChrW(74), ChrW(189) & ChrW(224) & ChrW(74), ChrW(190) & ChrW(224) & ChrW(74), ChrW(191) & ChrW(224) & ChrW(74), _
    ChrW(192) & ChrW(74), ChrW(193) & ChrW(224) & ChrW(74), ChrW(194) & ChrW(224) & ChrW(74), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(74), ChrW(113) & ChrW(74), ChrW(196) & ChrW(253) & ChrW(74), ChrW(198) & ChrW(224) & ChrW(74), ChrW(199) & ChrW(224) & ChrW(74), _
    ChrW(200) & ChrW(224) & ChrW(74), ChrW(201) & ChrW(224) & ChrW(74), ChrW(202) & ChrW(74), ChrW(239) & ChrW(202) & ChrW(74), ChrW(118) & ChrW(74), ChrW(205) & ChrW(224) & ChrW(74), ChrW(206) & ChrW(224) & ChrW(74), ChrW(207) & ChrW(224) & ChrW(74), _
    ChrW(208) & ChrW(224) & ChrW(74), ChrW(209) & ChrW(74), ChrW(211) & ChrW(74), ChrW(239) & ChrW(211) & ChrW(74), ChrW(213) & ChrW(224) & ChrW(74), ChrW(214) & ChrW(224) & ChrW(74))
    
    rKree_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2368), _
    ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2368), _
    ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2368), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2368), _
    ChrW(2352) & ChrW(23812325) & ChrW(2381) & ChrW(2359) & ChrW(2368), ChrW(2352) & ChrW(23812332) & ChrW(2381) & ChrW(2334) & ChrW(2368))
        
    For i = LBound(rKree_Devnagari) To UBound(rKree_Devnagari)
        RngFnd = Replace(RngFnd, rKree_Uni(i), rKree_Devnagari(i))
    Next
        
    'â§ê â©àê    âªàê    âiê â®ê â°àê    â²ê â³àê    â´àê    â¸àê    â¹ê âºê â"ê â¼ê â½àê    â¾àê    â¿àê    âÀê âÁàê    âÂàê    âïÂàê   âqê âÄê âÆàê    âÇàê    âÈàê    âÉàê    âÊê âïÊê    âvê âÍàê    âÎàê    âÏàê    âÐàê    âÑê âÓê âïÓê    âÕàê    âÖàê
    rKri_Devnagari = Array( _
    ChrW(226) & ChrW(167) & ChrW(234), ChrW(226) & ChrW(169) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(170) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(105) & ChrW(234), ChrW(226) & ChrW(174) & ChrW(234), ChrW(226) & ChrW(176) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(178) & ChrW(234), ChrW(226) & ChrW(179) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(180) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(184) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(185) & ChrW(234), ChrW(226) & ChrW(186) & ChrW(234), ChrW(226) & ChrW(187) & ChrW(234), _
    ChrW(226) & ChrW(188) & ChrW(234), ChrW(226) & ChrW(189) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(190) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(191) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(192) & ChrW(234), ChrW(226) & ChrW(193) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(194) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(239) & ChrW(194) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(113) & ChrW(234), ChrW(226) & ChrW(196) & ChrW(234), ChrW(226) & ChrW(198) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(199) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(200) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(201) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(202) & ChrW(234), ChrW(226) & ChrW(239) & ChrW(202) & ChrW(234), ChrW(226) & ChrW(118) & ChrW(234), ChrW(226) & ChrW(205) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(206) & ChrW(224) & ChrW(234), _
    ChrW(226) & ChrW(207) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(208) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(209) & ChrW(234), ChrW(226) & ChrW(211) & ChrW(234), ChrW(226) & ChrW(239) & ChrW(211) & ChrW(234), _
    ChrW(226) & ChrW(213) & ChrW(224) & ChrW(234), ChrW(226) & ChrW(214) & ChrW(224) & ChrW(234))
    
    rKri_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2367), _
    ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2367), _
    ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2367), _
    ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2367), _
    ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2367), _
    ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2367), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2367), ChrW(2352) & ChrW(23812325) & ChrW(2381) & ChrW(2359) & ChrW(2367), ChrW(2352) & ChrW(23812332) & ChrW(2381) & ChrW(2334) & ChrW(2367))

    For i = LBound(rKri_Devnagari) To UBound(rKri_Devnagari)
        RngFnd = Replace(RngFnd, rKree_Uni(i), rKree_Devnagari(i))
    Next
    
    'â§ë â©àë    âªàë    âië â®ë â°àë    â²ë â³àë    â´àë    â¸àë    â¹ë âºë â"ë â¼ë â½àë    â¾àë    â¿àë    âÀë âÁàë    âÂàë    âïÂàë   âqë âÄë âÆàë    âÇàë    âÈàë    âÉàë    âÊë âïÊë    âvë âÍàë    âÎàë    âÏàë    âÐàë    âÑë âÓë âïÓë    âÕàë    âÖàë
    rKrim_Devnagari = Array( _
    ChrW(226) & ChrW(167) & ChrW(235), ChrW(226) & ChrW(169) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(170) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(105) & ChrW(235), ChrW(226) & ChrW(174) & ChrW(235), ChrW(226) & ChrW(176) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(178) & ChrW(235), ChrW(226) & ChrW(179) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(180) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(184) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(185) & ChrW(235), ChrW(226) & ChrW(186) & ChrW(235), ChrW(226) & ChrW(187) & ChrW(235), ChrW(226) & ChrW(188) & ChrW(235), ChrW(226) & ChrW(189) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(190) & ChrW(224) & ChrW(235), _
    ChrW(226) & ChrW(191) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(192) & ChrW(235), ChrW(226) & ChrW(193) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(194) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(239) & ChrW(194) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(113) & ChrW(235), ChrW(226) & ChrW(196) & ChrW(235), ChrW(226) & ChrW(198) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(199) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(200) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(201) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(202) & ChrW(235), ChrW(226) & ChrW(239) & ChrW(202) & ChrW(235), ChrW(226) & ChrW(118) & ChrW(235), ChrW(226) & ChrW(205) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(206) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(207) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(208) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(209) & ChrW(235), ChrW(226) & ChrW(211) & ChrW(235), ChrW(226) & ChrW(239) & ChrW(211) & ChrW(235), _
    ChrW(226) & ChrW(213) & ChrW(224) & ChrW(235), ChrW(226) & ChrW(214) & ChrW(224) & ChrW(235))
    
    rKrim_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(23812359) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(23812334) & ChrW(2367) & ChrW(2306))
        
    For i = LBound(rKrim_Devnagari) To UBound(rKrim_Devnagari)
        RngFnd = Replace(RngFnd, rKrim_Uni(i), rKrim_Devnagari(i))
    Next
        
    '§íB RàB ªàíB    iíB ®D  °àíB    ²D  ³àíB    ´àD ¸àíB    ¹D  ºD  "D  ¼D  ½àíB    ØàB ¿àíB    ÀíB ÁàíB    ÂàíB    ïÂàíB   qíB ÄíB ÆàíB    ÇàíB    ÈàíB    ÉàíB    ÊêB ïÊêB    víB ÍàíB    óB  ÏàíB    ÐàíB    gB  ÓD  ïÓD ÕàD ÖàD
    kru_Devnagari = Array( _
    ChrW(167) & ChrW(237) & ChrW(66), ChrW(82) & ChrW(224) & ChrW(66), ChrW(170) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(105) & ChrW(237) & ChrW(66), ChrW(174) & ChrW(68), ChrW(176) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(178) & ChrW(68), ChrW(179) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(180) & ChrW(224) & ChrW(68), ChrW(184) & ChrW(224) & ChrW(237) & ChrW(66), _
    ChrW(185) & ChrW(68), ChrW(186) & ChrW(68), ChrW(187) & ChrW(68), ChrW(188) & ChrW(68), ChrW(189) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(216) & ChrW(224) & ChrW(66), ChrW(191) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(192) & ChrW(237) & ChrW(66), ChrW(193) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(194) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(113) & ChrW(237) & ChrW(66), _
    ChrW(196) & ChrW(237) & ChrW(66), ChrW(198) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(199) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(200) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(201) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(202) & ChrW(234) & ChrW(66), _
    ChrW(239) & ChrW(202) & ChrW(234) & ChrW(66), ChrW(118) & ChrW(237) & ChrW(66), ChrW(205) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(243) & ChrW(66), ChrW(207) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(208) & ChrW(224) & ChrW(237) & ChrW(66), ChrW(103) & ChrW(66), ChrW(211) & ChrW(68), ChrW(239) & ChrW(211) & ChrW(68), ChrW(213) & ChrW(224) & ChrW(68), ChrW(214) & ChrW(224) & ChrW(68))
    
    kru_Uni = Array( _
    ChrW(2325) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2326) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2327) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2328) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2329) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2330) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2331) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2332) & ChrW(2381) & ChrW(2352) & ChrW(2369), _
    ChrW(2333) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2334) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2335) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2336) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2337) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2338) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2339) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2340) & ChrW(2381) & ChrW(2352) & ChrW(2369), _
    ChrW(2341) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2342) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2343) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2344) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2345) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2346) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2347) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2348) & ChrW(2381) & ChrW(2352) & ChrW(2369), _
    ChrW(2349) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2350) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2351) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2353) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2354) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2357) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2358) & ChrW(2381) & ChrW(2352) & ChrW(2369), _
    ChrW(2359) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2360) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2361) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2355) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2356) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2325) & ChrW(23812359) & ChrW(2381) & ChrW(2352) & ChrW(2369), ChrW(2332) & ChrW(23812334) & ChrW(2381) & ChrW(2352) & ChrW(2369))
    
    For i = LBound(kru_Devnagari) To UBound(kru_Devnagari)
        RngFnd = Replace(RngFnd, rKrim_Uni(i), rKrim_Devnagari(i))
    Next
    
    '§íC RàC ªàíC    iíC ®E  °àíC    ²E  ³àE ´àE ¸àE ¹E  ºE  "E  ¼E  ½àíC    ØàC ¿àíC    ÀíC ÁàíC    ÂàíC    ïÂàE    qíC ÄíC ÆàíC    ÇàíC    ÈàíC    ÉàíC    ÊêC ïÊêC    vE  ÍàíC    óC  ÏàíC    ÐàE ÑE  ÓE  ïÓE ÕàE ÖàE
    kroo_Devnagari = Array( _
    ChrW(167) & ChrW(237) & ChrW(67), ChrW(82) & ChrW(224) & ChrW(67), ChrW(170) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(105) & ChrW(237) & ChrW(67), ChrW(174) & ChrW(69), ChrW(176) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(178) & ChrW(69), ChrW(179) & ChrW(224) & ChrW(69), ChrW(180) & ChrW(224) & ChrW(69), _
    ChrW(184) & ChrW(224) & ChrW(69), ChrW(185) & ChrW(69), ChrW(186) & ChrW(69), ChrW(187) & ChrW(69), ChrW(188) & ChrW(69), ChrW(189) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(216) & ChrW(224) & ChrW(67), ChrW(191) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(192) & ChrW(237) & ChrW(67), ChrW(193) & ChrW(224) & ChrW(237) & ChrW(67), _
    ChrW(194) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(69), ChrW(113) & ChrW(237) & ChrW(67), ChrW(196) & ChrW(237) & ChrW(67), ChrW(198) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(199) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(200) & ChrW(224) & ChrW(237) & ChrW(67), _
    ChrW(201) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(202) & ChrW(234) & ChrW(67), ChrW(239) & ChrW(202) & ChrW(234) & ChrW(67), ChrW(118) & ChrW(69), ChrW(205) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(243) & ChrW(67), ChrW(207) & ChrW(224) & ChrW(237) & ChrW(67), ChrW(208) & ChrW(224) & ChrW(69), _
    ChrW(209) & ChrW(69), ChrW(211) & ChrW(69), ChrW(239) & ChrW(211) & ChrW(69), ChrW(213) & ChrW(224) & ChrW(69), ChrW(214) & ChrW(224) & ChrW(69))
    
    kroo_Uni = Array( _
    ChrW(2325) & ChrW(2381) & ChrW(2352) & ChrW(2370), ChrW(2326) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2327) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2328) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2329) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2330) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2331) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2332) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), _
    ChrW(2333) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2334) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2335) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2336) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2337) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2338) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2339) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2340) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2341) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2342) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2343) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2344) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), _
    ChrW(2345) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2346) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2347) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2348) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2349) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), _
    ChrW(2350) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2351) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2353) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2354) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), _
    ChrW(2357) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2358) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2359) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2360) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2361) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), _
    ChrW(2355) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2356) & ChrW(2381) & ChrW(2352) & ChrW(2370) & ChrW(2306), ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(23812352) & ChrW(2370) & ChrW(2306), ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(23812352) & ChrW(2370) & ChrW(2306))

    For i = LBound(kroo_Devnagari) To UBound(kroo_Devnagari)
        RngFnd = Replace(RngFnd, kroo_Uni(i), kroo_Devnagari(i))
    Next

    '§I  ©àI ªàI iI  ®I  °àI ²I  ³àI ´àI ¸àI ¹I  ºI  "I  ¼I  ½àI ¾àI ¿àI ÀI  ÁàI ÂàI ïÂàI    qI  ÄI  ÆàI ÇàI ÈàI ÉàI ÊI  ïÊI vI  ÍàI ÎàI ÏàI ÐàI ÑI  ÓI  ïÓI ÕàI ÖàI
    rKaim_Devnagari = Array( _
    ChrW(167) & ChrW(73), ChrW(169) & ChrW(224) & ChrW(73), ChrW(170) & ChrW(224) & ChrW(73), ChrW(105) & ChrW(73), ChrW(174) & ChrW(73), ChrW(176) & ChrW(224) & ChrW(73), ChrW(178) & ChrW(73), ChrW(179) & ChrW(224) & ChrW(73), ChrW(180) & ChrW(224) & ChrW(73), ChrW(184) & ChrW(224) & ChrW(73), ChrW(185) & ChrW(73), _
    ChrW(186) & ChrW(73), ChrW(187) & ChrW(73), ChrW(188) & ChrW(73), ChrW(189) & ChrW(224) & ChrW(73), ChrW(190) & ChrW(224) & ChrW(73), ChrW(191) & ChrW(224) & ChrW(73), ChrW(192) & ChrW(73), ChrW(193) & ChrW(224) & ChrW(73), _
    ChrW(194) & ChrW(224) & ChrW(73), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(73), ChrW(113) & ChrW(73), ChrW(196) & ChrW(73), ChrW(198) & ChrW(224) & ChrW(73), ChrW(199) & ChrW(224) & ChrW(73), ChrW(200) & ChrW(224) & ChrW(73), ChrW(201) & ChrW(224) & ChrW(73), ChrW(202) & ChrW(73), ChrW(239) & ChrW(202) & ChrW(73), _
    ChrW(118) & ChrW(73), ChrW(205) & ChrW(224) & ChrW(73), ChrW(206) & ChrW(224) & ChrW(73), ChrW(207) & ChrW(224) & ChrW(73), ChrW(208) & ChrW(224) & ChrW(73), ChrW(209) & ChrW(73), ChrW(211) & ChrW(73), ChrW(239) & ChrW(211) & ChrW(73), _
    ChrW(213) & ChrW(224) & ChrW(73), ChrW(214) & ChrW(224) & ChrW(73))

    rKaim_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2376) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2376) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2376) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2376) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(23812359) & ChrW(2376) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(23812334) & ChrW(2376) & ChrW(2306))

    For i = LBound(rKaim_Devnagari) To UBound(rKaim_Devnagari)
        RngFnd = Replace(RngFnd, rKaim_Uni(i), rKaim_Devnagari(i))
    Next

    '§H  ©àH ªàH iH  ®H  °àH ²H  ³àH ´àH ¸àH ¹H  ºH  "H  ¼H  ½àH ¾àH ¿àH ÀH  ÁàH ÂàH ïÂàH    qH  ÄH  ÆàH ÇàH ÈàH ÉàH ÊH  ïÊH vH  ÍàH ÎàH ÏàH ÐàH ÑH  ÓH  ïÓH ÕàH ÖàH
    rKai_Devnagari = Array( _
    ChrW(167) & ChrW(72), ChrW(169) & ChrW(224) & ChrW(72), ChrW(170) & ChrW(224) & ChrW(72), ChrW(105) & ChrW(72), ChrW(174) & ChrW(72), ChrW(176) & ChrW(224) & ChrW(72), ChrW(178) & ChrW(72), ChrW(179) & ChrW(224) & ChrW(72), ChrW(180) & ChrW(224) & ChrW(72), ChrW(184) & ChrW(224) & ChrW(72), ChrW(185) & ChrW(72), _
    ChrW(186) & ChrW(72), ChrW(187) & ChrW(72), ChrW(188) & ChrW(72), ChrW(189) & ChrW(224) & ChrW(72), ChrW(190) & ChrW(224) & ChrW(72), ChrW(191) & ChrW(224) & ChrW(72), ChrW(192) & ChrW(72), ChrW(193) & ChrW(224) & ChrW(72), ChrW(194) & ChrW(224) & ChrW(72), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(72), _
    ChrW(113) & ChrW(72), ChrW(196) & ChrW(72), ChrW(198) & ChrW(224) & ChrW(72), ChrW(199) & ChrW(224) & ChrW(72), ChrW(200) & ChrW(224) & ChrW(72), ChrW(201) & ChrW(224) & ChrW(72), ChrW(202) & ChrW(72), ChrW(239) & ChrW(202) & ChrW(72), _
    ChrW(118) & ChrW(72), ChrW(205) & ChrW(224) & ChrW(72), ChrW(206) & ChrW(224) & ChrW(72), ChrW(207) & ChrW(224) & ChrW(72), ChrW(208) & ChrW(224) & ChrW(72), ChrW(209) & ChrW(72), ChrW(211) & ChrW(72), ChrW(239) & ChrW(211) & ChrW(72), ChrW(213) & ChrW(224) & ChrW(72), ChrW(214) & ChrW(224) & ChrW(72))

    rKai_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2376), _
    ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2376), _
    ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2376), _
    ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2376), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2376), _
    ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2376), ChrW(2352) & ChrW(23812325) & ChrW(2381) & ChrW(2359) & ChrW(2376), ChrW(2352) & ChrW(23812332) & ChrW(2381) & ChrW(2334) & ChrW(2376))

    For i = LBound(rKai_Devnagari) To UBound(rKai_Devnagari)
        RngFnd = Replace(RngFnd, rKai_Uni(i), rKai_Devnagari(i))
    Next

    

    '§G  ©àG ªàG iG  ®G  °àG ²G  ³àG ´àG ¸àG ¹G  ºG  »G  ¼G  ½àG ¾àG ¿àG ÀG  ÁàG ÂàG ïÂàG    qG  ÄG  ÆàG ÇàG ÈàG ÉàG ÊG  ïÊG vG  ÍàG ÎàG ÏàG ÐàG ÑG  ÓG  ïÓG ÕàG ÖàG
    rKem_Devnagari = Array( _
    ChrW(167) & ChrW(71), ChrW(169) & ChrW(224) & ChrW(71), ChrW(170) & ChrW(224) & ChrW(71), ChrW(105) & ChrW(71), ChrW(174) & ChrW(71), ChrW(176) & ChrW(224) & ChrW(71), ChrW(178) & ChrW(71), ChrW(179) & ChrW(224) & ChrW(71), ChrW(180) & ChrW(224) & ChrW(71), ChrW(184) & ChrW(224) & ChrW(71), ChrW(185) & ChrW(71), _
    ChrW(186) & ChrW(71), ChrW(187) & ChrW(71), ChrW(188) & ChrW(71), ChrW(189) & ChrW(224) & ChrW(71), ChrW(190) & ChrW(224) & ChrW(71), ChrW(191) & ChrW(224) & ChrW(71), ChrW(192) & ChrW(71), ChrW(193) & ChrW(224) & ChrW(71), ChrW(194) & ChrW(224) & ChrW(71), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(71), _
    ChrW(113) & ChrW(71), ChrW(196) & ChrW(71), ChrW(198) & ChrW(224) & ChrW(71), ChrW(199) & ChrW(224) & ChrW(71), ChrW(200) & ChrW(224) & ChrW(71), ChrW(201) & ChrW(224) & ChrW(71), ChrW(202) & ChrW(71), ChrW(239) & ChrW(202) & ChrW(71), ChrW(118) & ChrW(71), ChrW(205) & ChrW(224) & ChrW(71), ChrW(206) & ChrW(224) & ChrW(71), ChrW(207) & ChrW(224) & ChrW(71), ChrW(208) & ChrW(224) & ChrW(71), _
    ChrW(209) & ChrW(71), ChrW(211) & ChrW(71), ChrW(239) & ChrW(211) & ChrW(71), ChrW(213) & ChrW(224) & ChrW(71), ChrW(214) & ChrW(224) & ChrW(71))

    rKem_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2375) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2375) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2375) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2375) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2375) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2375) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2375) & ChrW(2306))

    For i = LBound(rKem_Devnagari) To UBound(rKem_Devnagari)
        RngFnd = Replace(RngFnd, rKem_Uni(i), rKem_Devnagari(i))
    Next

    '§F  ©àF ªàF iF  ®F  °àF ²F  ³àF ´àF ¸àF ¹F  ºF  »F  ¼F  ½àF ¾àF ¿àF ÀF  ÁàF ÂàF ïÂàF    qF  ÄF  ÆàF ÇàF ÈàF ÉàF ÊF  ïÊF vF  ÍàF ÎàF ÏàF ÐàF ÑF  ÓF  ïÓF ÕàF ÖàF
    rKe_Devnagari = Array( _
    ChrW(167) & ChrW(70), ChrW(169) & ChrW(224) & ChrW(70), ChrW(170) & ChrW(224) & ChrW(70), ChrW(105) & ChrW(70), ChrW(174) & ChrW(70), ChrW(176) & ChrW(224) & ChrW(70), ChrW(178) & ChrW(70), ChrW(179) & ChrW(224) & ChrW(70), ChrW(180) & ChrW(224) & ChrW(70), ChrW(184) & ChrW(224) & ChrW(70), ChrW(185) & ChrW(70), ChrW(186) & ChrW(70), ChrW(187) & ChrW(70), ChrW(188) & ChrW(70), ChrW(189) & ChrW(224) & ChrW(70), ChrW(190) & ChrW(224) & ChrW(70), ChrW(191) & ChrW(224) & ChrW(70), _
    ChrW(192) & ChrW(70), ChrW(193) & ChrW(224) & ChrW(70), ChrW(194) & ChrW(224) & ChrW(70), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(70), ChrW(113) & ChrW(70), ChrW(196) & ChrW(70), ChrW(198) & ChrW(224) & ChrW(70), ChrW(199) & ChrW(224) & ChrW(70), ChrW(200) & ChrW(224) & ChrW(70), ChrW(201) & ChrW(224) & ChrW(70), _
    ChrW(202) & ChrW(70), ChrW(239) & ChrW(202) & ChrW(70), ChrW(118) & ChrW(70), ChrW(205) & ChrW(224) & ChrW(70), ChrW(206) & ChrW(224) & ChrW(70), ChrW(207) & ChrW(224) & ChrW(70), ChrW(208) & ChrW(224) & ChrW(70), _
    ChrW(209) & ChrW(70), ChrW(211) & ChrW(70), ChrW(239) & ChrW(211) & ChrW(70), ChrW(213) & ChrW(224) & ChrW(70), ChrW(214) & ChrW(224) & ChrW(70))

    rKe_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2375), _
    ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2375), _
    ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2375), _
    ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2375), _
    ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2375), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2375), _
    ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2375))

    For i = LBound(rKe_Devnagari) To UBound(rKe_Devnagari)
        RngFnd = Replace(RngFnd, rKe_Uni(i), rKe_Devnagari(i))
    Next
    
    '
    rKam_Devnagari = Array( _
    ChrW(167) & ChrW(253) & ChrW(224) & ChrW(235), ChrW(169) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(170) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(105) & ChrW(224) & ChrW(235), ChrW(174) & ChrW(224) & ChrW(235), ChrW(176) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(178) & ChrW(224) & ChrW(235), ChrW(179) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(180) & ChrW(224) & ChrW(224) & ChrW(235), _
    ChrW(184) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(185) & ChrW(224) & ChrW(235), ChrW(186) & ChrW(224) & ChrW(235), ChrW(187) & ChrW(224) & ChrW(235), ChrW(188) & ChrW(224) & ChrW(235), ChrW(189) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(190) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(191) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(192) & ChrW(224) & ChrW(235), ChrW(193) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(194) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(224) & ChrW(235), _
    ChrW(113) & ChrW(224) & ChrW(235), ChrW(196) & ChrW(253) & ChrW(224) & ChrW(235), ChrW(198) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(199) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(200) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(201) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(202) & ChrW(224) & ChrW(235), _
    ChrW(239) & ChrW(202) & ChrW(224) & ChrW(235), ChrW(118) & ChrW(224) & ChrW(235), ChrW(205) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(206) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(207) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(208) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(209) & ChrW(224) & ChrW(235), _
    ChrW(211) & ChrW(224) & ChrW(235), ChrW(239) & ChrW(211) & ChrW(224) & ChrW(235), ChrW(213) & ChrW(224) & ChrW(224) & ChrW(235), ChrW(214) & ChrW(224) & ChrW(224) & ChrW(235))

    rKam_Desi = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2367) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2367) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2367) & ChrW(2306))

    For i = LBound(rKam_Devnagari) To UBound(rKam_Devnagari)
        RngFnd = Replace(RngFnd, rKam_Desi(i), rKam_Devnagari(i))
    Next

    '§ýàê    ©ààê    ªààê    iàê ®àê °ààê    ²àê ³ààê    ´ààê    ¸ààê    ¹àê ºàê »àê ¼àê ½ààê    ¾ààê    ¿ààê    Ààê Áààê    Âààê    ïÂààê   qàê Äýàê    Æààê    Çààê    Èààê    Éààê    Êàê ïÊàê    vàê Íààê    Îààê    Ïààê    Ðààê    Ñàê Óàê ïÓàê    Õààê    Öààê
    rKa_Devnagari = Array( _
    ChrW(167) & ChrW(253) & ChrW(224) & ChrW(234), ChrW(169) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(170) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(105) & ChrW(224) & ChrW(234), ChrW(174) & ChrW(224) & ChrW(234), ChrW(176) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(178) & ChrW(224) & ChrW(234), ChrW(179) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(180) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(184) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(185) & ChrW(224) & ChrW(234), _
    ChrW(186) & ChrW(224) & ChrW(234), ChrW(187) & ChrW(224) & ChrW(234), ChrW(188) & ChrW(224) & ChrW(234), ChrW(189) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(190) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(191) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(192) & ChrW(224) & ChrW(234), ChrW(193) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(194) & ChrW(224) & ChrW(224) & ChrW(234), _
    ChrW(239) & ChrW(194) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(113) & ChrW(224) & ChrW(234), ChrW(196) & ChrW(253) & ChrW(224) & ChrW(234), ChrW(198) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(199) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(200) & ChrW(224) & ChrW(224) & ChrW(234), _
    ChrW(201) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(202) & ChrW(224) & ChrW(234), ChrW(239) & ChrW(202) & ChrW(224) & ChrW(234), ChrW(118) & ChrW(224) & ChrW(234), ChrW(205) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(206) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(207) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(208) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(209) & ChrW(224) & ChrW(234), _
    ChrW(211) & ChrW(224) & ChrW(234), ChrW(239) & ChrW(211) & ChrW(224) & ChrW(234), ChrW(213) & ChrW(224) & ChrW(224) & ChrW(234), ChrW(214) & ChrW(224) & ChrW(224) & ChrW(234))

    rKa_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2366), _
    ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2366), _
    ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2366), _
    ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2366), _
    ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2366), _
    ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2366))

    For i = LBound(rKa_Devnagari) To UBound(rKa_Devnagari)
        RngFnd = Replace(RngFnd, rKa_Uni(i), rKa_Devnagari(i))
    Next

    '
    rKm_Devnagari = Array( _
    ChrW(167) & ChrW(235), ChrW(169) & ChrW(224) & ChrW(235), ChrW(170) & ChrW(224) & ChrW(235), ChrW(105) & ChrW(235), ChrW(174) & ChrW(235), ChrW(176) & ChrW(224) & ChrW(235), ChrW(178) & ChrW(235), ChrW(179) & ChrW(224) & ChrW(235), ChrW(180) & ChrW(224) & ChrW(235), ChrW(184) & ChrW(224) & ChrW(235), ChrW(185) & ChrW(235), _
    ChrW(186) & ChrW(235), ChrW(187) & ChrW(235), ChrW(188) & ChrW(235), ChrW(189) & ChrW(224) & ChrW(235), ChrW(190) & ChrW(224) & ChrW(235), ChrW(191) & ChrW(224) & ChrW(235), ChrW(192) & ChrW(235), ChrW(193) & ChrW(224) & ChrW(235), ChrW(194) & ChrW(224) & ChrW(235), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(235), _
    ChrW(113) & ChrW(235), ChrW(196) & ChrW(235), ChrW(198) & ChrW(224) & ChrW(235), ChrW(199) & ChrW(224) & ChrW(235), ChrW(200) & ChrW(224) & ChrW(235), ChrW(201) & ChrW(224) & ChrW(235), ChrW(202) & ChrW(235), ChrW(239) & ChrW(202) & ChrW(235), ChrW(118) & ChrW(235), ChrW(205) & ChrW(224) & ChrW(235), _
    ChrW(206) & ChrW(224) & ChrW(235), ChrW(207) & ChrW(224) & ChrW(235), ChrW(208) & ChrW(224) & ChrW(235), ChrW(209) & ChrW(235), ChrW(211) & ChrW(235), ChrW(239) & ChrW(211) & ChrW(235), ChrW(213) & ChrW(224) & ChrW(235), ChrW(214) & ChrW(224) & ChrW(235))

    rKm_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2306))
    
    For i = LBound(rKm_Devnagari) To UBound(rKm_Devnagari)
        RngFnd = Replace(RngFnd, rKm_Uni(i), rKm_Devnagari(i))
    Next
    
    '
    rK_Devnagari = Array( _
    ChrW(167) & ChrW(234), ChrW(169) & ChrW(224) & ChrW(234), ChrW(170) & ChrW(224) & ChrW(234), ChrW(105) & ChrW(234), ChrW(174) & ChrW(234), ChrW(176) & ChrW(224) & ChrW(234), ChrW(178) & ChrW(234), ChrW(179) & ChrW(224) & ChrW(234), ChrW(180) & ChrW(224) & ChrW(234), ChrW(184) & ChrW(224) & ChrW(234), ChrW(185) & ChrW(234), ChrW(186) & ChrW(234), ChrW(187) & ChrW(234), ChrW(188) & ChrW(234), ChrW(189) & ChrW(224) & ChrW(234), _
    ChrW(190) & ChrW(224) & ChrW(234), ChrW(191) & ChrW(224) & ChrW(234), ChrW(192) & ChrW(234), ChrW(193) & ChrW(224) & ChrW(234), ChrW(194) & ChrW(224) & ChrW(234), ChrW(239) & ChrW(194) & ChrW(224) & ChrW(234), ChrW(113) & ChrW(234), ChrW(196) & ChrW(234), ChrW(198) & ChrW(224) & ChrW(234), ChrW(199) & ChrW(224) & ChrW(234), ChrW(200) & ChrW(224) & ChrW(234), ChrW(201) & ChrW(224) & ChrW(234), ChrW(202) & ChrW(234), ChrW(239) & ChrW(202) & ChrW(234), ChrW(118) & ChrW(234), ChrW(205) & ChrW(224) & ChrW(234), _
    ChrW(206) & ChrW(224) & ChrW(234), ChrW(207) & ChrW(224) & ChrW(234), ChrW(208) & ChrW(224) & ChrW(234), ChrW(209) & ChrW(234), ChrW(211) & ChrW(234), ChrW(239) & ChrW(211) & ChrW(234), ChrW(213) & ChrW(224) & ChrW(234), ChrW(214) & ChrW(224) & ChrW(234))
    
    rK_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325), ChrW(2352) & ChrW(2381) & ChrW(2326), ChrW(2352) & ChrW(2381) & ChrW(2327), ChrW(2352) & ChrW(2381) & ChrW(2328), ChrW(2352) & ChrW(2381) & ChrW(2329), ChrW(2352) & ChrW(2381) & ChrW(2330), ChrW(2352) & ChrW(2381) & ChrW(2331), ChrW(2352) & ChrW(2381) & ChrW(2332), ChrW(2352) & ChrW(2381) & ChrW(2333), ChrW(2352) & ChrW(2381) & ChrW(2334), ChrW(2352) & ChrW(2381) & ChrW(2335), ChrW(2352) & ChrW(2381) & ChrW(2336), ChrW(2352) & ChrW(2381) & ChrW(2337), _
    ChrW(2352) & ChrW(2381) & ChrW(2338), ChrW(2352) & ChrW(2381) & ChrW(2339), ChrW(2352) & ChrW(2381) & ChrW(2340), ChrW(2352) & ChrW(2381) & ChrW(2341), ChrW(2352) & ChrW(2381) & ChrW(2342), ChrW(2352) & ChrW(2381) & ChrW(2343), ChrW(2352) & ChrW(2381) & ChrW(2344), ChrW(2352) & ChrW(2381) & ChrW(2345), ChrW(2352) & ChrW(2381) & ChrW(2346), ChrW(2352) & ChrW(2381) & ChrW(2347), ChrW(2352) & ChrW(2381) & ChrW(2348), _
    ChrW(2352) & ChrW(2381) & ChrW(2349), ChrW(2352) & ChrW(2381) & ChrW(2350), ChrW(2352) & ChrW(2381) & ChrW(2351), ChrW(2352) & ChrW(2381) & ChrW(2352), ChrW(2352) & ChrW(2381) & ChrW(2353), ChrW(2352) & ChrW(2381) & ChrW(2354), ChrW(2352) & ChrW(2381) & ChrW(2357), ChrW(2352) & ChrW(2381) & ChrW(2358), ChrW(2352) & ChrW(2381) & ChrW(2359), ChrW(2352) & ChrW(2381) & ChrW(2360), _
    ChrW(2352) & ChrW(2381) & ChrW(2361), ChrW(2352) & ChrW(2381) & ChrW(2355), ChrW(2352) & ChrW(2381) & ChrW(2356), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334))

    For i = LBound(rK_Devnagari) To UBound(rK_Devnagari)
        RngFnd = Replace(RngFnd, rK_Uni(i), rK_Devnagari(i))
    Next
    
    '
    k_Devnagari = Array( _
    ChrW(167), ChrW(169) & ChrW(224), ChrW(170) & ChrW(224), ChrW(105), ChrW(174), ChrW(176) & ChrW(224), ChrW(178), ChrW(179) & ChrW(224), ChrW(180) & ChrW(224), ChrW(184) & ChrW(224), ChrW(185), ChrW(186), ChrW(187), ChrW(188), ChrW(189) & ChrW(224), ChrW(190) & ChrW(224), ChrW(191) & ChrW(224), ChrW(192), ChrW(193) & ChrW(224), _
    ChrW(194) & ChrW(224), ChrW(239) & ChrW(194) & ChrW(224), ChrW(113), ChrW(196), ChrW(198) & ChrW(224), ChrW(199) & ChrW(224), ChrW(200) & ChrW(224), ChrW(201) & ChrW(224), ChrW(202), ChrW(239) & ChrW(202), ChrW(118), ChrW(205) & ChrW(224), ChrW(206) & ChrW(224), ChrW(207) & ChrW(224), ChrW(208) & ChrW(224), ChrW(209), ChrW(211), ChrW(239) & ChrW(211), ChrW(213) & ChrW(224), ChrW(214) & ChrW(224))

    k_Uni = Array( _
    ChrW(2325), ChrW(2326), ChrW(2327), ChrW(2328), ChrW(2329), ChrW(2330), ChrW(2331), ChrW(2332), ChrW(2333), ChrW(2334), ChrW(2335), ChrW(2336), ChrW(2337), ChrW(2338), ChrW(2339), ChrW(2340), ChrW(2341), ChrW(2342), ChrW(2343), ChrW(2344), ChrW(2345), ChrW(2346), ChrW(2347), ChrW(2348), _
    ChrW(2349), ChrW(2350), ChrW(2351), ChrW(2352), ChrW(2353), ChrW(2354), ChrW(2357), ChrW(2358), ChrW(2359), ChrW(2360), ChrW(2361), ChrW(2355), ChrW(2356), ChrW(2325) & ChrW(2381) & ChrW(2359), ChrW(2332) & ChrW(2381) & ChrW(2334))

    For i = LBound(k_Devnagari) To UBound(k_Devnagari)
        RngFnd = Replace(RngFnd, k_Uni(i), k_Devnagari(i))
    Next
    
    '¡àè ¡àç ¡àé ¡à  ¡Þ  ¡Ý  ¡ß  ¡   ¢ê  ¢   ¤   £   ¥ç  ¥   ¦æ  ¦   ×æ  ×z
    Dim aAA_Devnagari() As Variant
    aAA_Devnagari = Array(ChrW(161) & ChrW(224) & ChrW(232), ChrW(161) & ChrW(224) & ChrW(231), ChrW(161) & ChrW(224) & ChrW(233), ChrW(161) & ChrW(224), ChrW(161) & ChrW(222), ChrW(161) & ChrW(221), ChrW(161) & ChrW(223), ChrW(161), ChrW(162) & ChrW(234), ChrW(162), ChrW(164), ChrW(163), ChrW(165) & ChrW(231), ChrW(165), ChrW(166) & ChrW(230), ChrW(166), ChrW(215) & ChrW(230), ChrW(215) & ChrW(122))
    Dim aAA_Uni() As Variant
    aAA_Uni = Array(ChrW(2324), ChrW(2323), ChrW(2321), ChrW(2310), ChrW(2310), ChrW(2309) & ChrW(2305), ChrW(2309) & ChrW(2307), ChrW(2309), ChrW(2312), ChrW(2311), ChrW(2314), ChrW(2313), ChrW(2320), ChrW(2319), ChrW(2315) & ChrW(2371), ChrW(2315), ChrW(2316), ChrW(2401))
    
    For i = LBound(aAA_Devnagari) To UBound(aAA_Devnagari)
        RngFnd = Replace(RngFnd, aAA_Uni(i), aAA_Devnagari(i))
    Next
    
    '_   -   |   \   E   e   u   U   I   i   O   o   a   :   <   ,   >   .   ›   .
    Dim aakar_ukar_Devnagari() As Variant
    aakar_ukar_Devnagari = Array(ChrW(124), ChrW(240), ChrW(45), ChrW(240), ChrW(238), ChrW(232), ChrW(231), ChrW(228), ChrW(229), ChrW(227), ChrW(226), ChrW(224) & ChrW(232), ChrW(224) & ChrW(231), ChrW(224), ChrW(223), ChrW(230), ChrW(44), ChrW(221), ChrW(222), ChrW(241), ChrW(233), ChrW(224) & ChrW(233))
    Dim aakar_ukar_Uni() As Variant
    aakar_ukar_Uni = Array(ChrW(124), ChrW(2364), ChrW(45), ChrW(46), ChrW(2381), ChrW(2376), ChrW(2375), ChrW(2369), ChrW(2370), ChrW(2368), ChrW(2367), ChrW(2380), ChrW(2379), ChrW(2366), ChrW(2307), ChrW(2371), ChrW(44), ChrW(2305), ChrW(2306), ChrW(2404), ChrW(2373), ChrW(2377))
    
    For i = LBound(aakar_ukar_Devnagari) To UBound(aakar_ukar_Devnagari)
        RngFnd = Replace(RngFnd, aakar_ukar_Uni(i), aakar_ukar_Devnagari(i))
    Next
    
    
    'fndList = mergeArrays(level1_Uni, level2_Uni)
    'fndList = mergeArrays(fndList, level3_Uni)
    'fndList = mergeArrays(fndList, level4_Uni)
    'fndList = mergeArrays(fndList, level5_Uni)
    'fndList = mergeArrays(fndList, level6_Uni)
    'fndList = mergeArrays(fndList, level7_Uni)
    'fndList = mergeArrays(fndList, level8_Uni)
    'fndList = mergeArrays(fndList, aAA_Uni)
    'fndList = mergeArrays(fndList, aakar_ukar_Uni)

    'rplcList = mergeArrays(level1_Devnagari, level2_Devnagari)
    'rplcList = mergeArrays(rplcList, level3_Devnagari)
    'rplcList = mergeArrays(rplcList, level4_Devnagari)
    'rplcList = mergeArrays(rplcList, level5_Devnagari)
    'rplcList = mergeArrays(rplcList, level6_Devnagari)
    'rplcList = mergeArrays(rplcList, level7_Devnagari)
    'rplcList = mergeArrays(rplcList, level8_Devnagari)
    'rplcList = mergeArrays(rplcList, aAA_Devnagari)
    'rplcList = mergeArrays(rplcList, aakar_ukar_Devnagari)
    

    

    'For i = LBound(fndList) To UBound(fndList)
    '    RngFnd = Replace(RngFnd, fndList(i), rplcList(i))
    'Next


    
    RngTxt.Text = RngFnd
    Selection.ClearFormatting
    

    
End Sub
