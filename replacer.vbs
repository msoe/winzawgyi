Sub Replacer()
'
' Replace Macro
' Macro recorded 10/2/2006 by Minn Myat Soe
'
    Selection.HomeKey Unit:=wdStory
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    ReDim WinZawgyi(255, 2) As Long
        
    WinZawgyi(0, 0) = 117
    WinZawgyi(1, 0) = 99
    WinZawgyi(2, 0) = 42
    WinZawgyi(3, 0) = 67
    WinZawgyi(4, 0) = 105
    WinZawgyi(5, 0) = 112
    WinZawgyi(6, 0) = 113
    WinZawgyi(7, 0) = 90
    WinZawgyi(8, 0) = 218
    WinZawgyi(9, 0) = 110
    WinZawgyi(10, 0) = 35
    WinZawgyi(11, 0) = 88
    WinZawgyi(12, 0) = 33
    WinZawgyi(13, 0) = 161
    WinZawgyi(14, 0) = 80
    WinZawgyi(15, 0) = 119
    WinZawgyi(16, 0) = 120
    WinZawgyi(17, 0) = 39
    WinZawgyi(18, 0) = 34
    WinZawgyi(19, 0) = 101
    WinZawgyi(20, 0) = 121
    WinZawgyi(21, 0) = 122
    WinZawgyi(22, 0) = 65
    WinZawgyi(23, 0) = 98
    WinZawgyi(24, 0) = 114
    WinZawgyi(25, 0) = 44
    WinZawgyi(26, 0) = 38
    WinZawgyi(27, 0) = 118
    WinZawgyi(28, 0) = 48
    WinZawgyi(29, 0) = 111
    WinZawgyi(30, 0) = 91
    WinZawgyi(31, 0) = 86
    WinZawgyi(32, 0) = 116
    WinZawgyi(33, 0) = 49
    WinZawgyi(34, 0) = 50
    WinZawgyi(35, 0) = 51
    WinZawgyi(36, 0) = 52
    WinZawgyi(37, 0) = 53
    WinZawgyi(38, 0) = 54
    WinZawgyi(39, 0) = 55
    WinZawgyi(40, 0) = 56
    WinZawgyi(41, 0) = 57
    WinZawgyi(42, 0) = 48
    WinZawgyi(43, 0) = 163
    WinZawgyi(44, 0) = 254
    WinZawgyi(45, 0) = 79
    WinZawgyi(46, 0) = 123
    WinZawgyi(47, 0) = 100
    WinZawgyi(48, 0) = 68
    WinZawgyi(49, 0) = 107
    WinZawgyi(50, 0) = 75
    WinZawgyi(51, 0) = 108
    WinZawgyi(52, 0) = 73
    WinZawgyi(53, 0) = 76
    WinZawgyi(54, 0) = 97
    WinZawgyi(55, 0) = 74
    WinZawgyi(56, 0) = 109
    WinZawgyi(57, 0) = 103
    WinZawgyi(58, 0) = 58
    WinZawgyi(59, 0) = 59
    WinZawgyi(60, 0) = 72
    WinZawgyi(61, 0) = 104
    WinZawgyi(62, 0) = 85
    WinZawgyi(63, 0) = 89
    WinZawgyi(64, 0) = 115
    WinZawgyi(65, 0) = 106
    WinZawgyi(66, 0) = 77
    WinZawgyi(67, 0) = 78
    WinZawgyi(68, 0) = 66
    WinZawgyi(69, 0) = 126
    WinZawgyi(70, 0) = 96
    WinZawgyi(71, 0) = 71
    WinZawgyi(72, 0) = 83
    WinZawgyi(73, 0) = 84
    WinZawgyi(74, 0) = 252
    WinZawgyi(75, 0) = 237
    WinZawgyi(76, 0) = 164
    WinZawgyi(77, 0) = 92
    WinZawgyi(78, 0) = 250
    WinZawgyi(79, 0) = 169
    WinZawgyi(80, 0) = 190
    WinZawgyi(81, 0) = 162
    WinZawgyi(82, 0) = 246
    WinZawgyi(83, 0) = 228
    WinZawgyi(84, 0) = 198
    WinZawgyi(85, 0) = 209
    WinZawgyi(86, 0) = 241
    WinZawgyi(87, 0) = 205
    WinZawgyi(88, 0) = 165
    WinZawgyi(89, 0) = 179
    WinZawgyi(90, 0) = 178
    WinZawgyi(91, 0) = 124
    WinZawgyi(92, 0) = 215
    WinZawgyi(93, 0) = 64
    WinZawgyi(94, 0) = 185
    WinZawgyi(95, 0) = 214
    WinZawgyi(96, 0) = 197
    WinZawgyi(97, 0) = 229
    WinZawgyi(98, 0) = 166
    WinZawgyi(99, 0) = 172
    WinZawgyi(100, 0) = 180
    WinZawgyi(101, 0) = 168
    WinZawgyi(102, 0) = 69
    WinZawgyi(103, 0) = 233
    WinZawgyi(104, 0) = 220
    WinZawgyi(105, 0) = 230
    WinZawgyi(106, 0) = 193
    WinZawgyi(107, 0) = 199
    WinZawgyi(108, 0) = 174
    WinZawgyi(109, 0) = 189
    WinZawgyi(110, 0) = 243
    WinZawgyi(111, 0) = 167
    WinZawgyi(112, 0) = 70
    WinZawgyi(113, 0) = 208
    WinZawgyi(114, 0) = 216
    WinZawgyi(115, 0) = 248
    WinZawgyi(116, 0) = 240
    WinZawgyi(117, 0) = 201
    WinZawgyi(118, 0) = 223
    WinZawgyi(119, 0) = 102
    WinZawgyi(120, 0) = 47
    WinZawgyi(121, 0) = 63
    WinZawgyi(122, 0) = 94
    WinZawgyi(123, 0) = 224
    WinZawgyi(124, 0) = 196
            
            
    WinZawgyi(0, 1) = &H1000
    WinZawgyi(1, 1) = &H1001
    WinZawgyi(2, 1) = &H1002
    WinZawgyi(3, 1) = &H1003
    WinZawgyi(4, 1) = &H1004
    WinZawgyi(5, 1) = &H1005
    WinZawgyi(6, 1) = &H1006
    WinZawgyi(7, 1) = &H1007
    WinZawgyi(8, 1) = &H1009
    WinZawgyi(9, 1) = &H100A
    WinZawgyi(10, 1) = &H100B
    WinZawgyi(11, 1) = &H100C
    WinZawgyi(12, 1) = &H100D
    WinZawgyi(13, 1) = &H100E
    WinZawgyi(14, 1) = &H100F
    WinZawgyi(15, 1) = &H1010
    WinZawgyi(16, 1) = &H1011
    WinZawgyi(17, 1) = &H1012
    WinZawgyi(18, 1) = &H1013
    WinZawgyi(19, 1) = &H1014
    WinZawgyi(20, 1) = &H1015
    WinZawgyi(21, 1) = &H1016
    WinZawgyi(22, 1) = &H1017
    WinZawgyi(23, 1) = &H1018
    WinZawgyi(24, 1) = &H1019
    WinZawgyi(25, 1) = &H101A
    WinZawgyi(26, 1) = &H101B
    WinZawgyi(27, 1) = &H101C
    WinZawgyi(28, 1) = &H101D
    WinZawgyi(29, 1) = &H101E
    WinZawgyi(30, 1) = &H101F
    WinZawgyi(31, 1) = &H1020
    WinZawgyi(32, 1) = &H1021
    WinZawgyi(33, 1) = &H1041
    WinZawgyi(34, 1) = &H1042
    WinZawgyi(35, 1) = &H1043
    WinZawgyi(36, 1) = &H1044
    WinZawgyi(37, 1) = &H1045
    WinZawgyi(38, 1) = &H1046
    WinZawgyi(39, 1) = &H1047
    WinZawgyi(40, 1) = &H1048
    WinZawgyi(41, 1) = &H1049
    WinZawgyi(42, 1) = &H1040
    WinZawgyi(43, 1) = &H1023
    WinZawgyi(44, 1) = &H1024
    WinZawgyi(45, 1) = &H1025
    WinZawgyi(46, 1) = &H1027
    WinZawgyi(47, 1) = &H102D
    WinZawgyi(48, 1) = &H102E
    WinZawgyi(49, 1) = &H102F
    WinZawgyi(50, 1) = &H1033
    WinZawgyi(51, 1) = &H1030
    WinZawgyi(52, 1) = &H1088
    WinZawgyi(53, 1) = &H1034
    WinZawgyi(54, 1) = &H1031
    WinZawgyi(55, 1) = &H1032
    WinZawgyi(56, 1) = &H102C
    WinZawgyi(57, 1) = &H102B
    WinZawgyi(58, 1) = &H105A
    WinZawgyi(59, 1) = &H1038
    WinZawgyi(60, 1) = &H1036
    WinZawgyi(61, 1) = &H1037
    WinZawgyi(62, 1) = &H1095
    WinZawgyi(63, 1) = &H1094
    WinZawgyi(64, 1) = &H103A
    WinZawgyi(65, 1) = &H103B
    WinZawgyi(66, 1) = &H107E
    WinZawgyi(67, 1) = &H107F
    WinZawgyi(68, 1) = &H1080
    WinZawgyi(69, 1) = &H1082
    WinZawgyi(70, 1) = &H1081
    WinZawgyi(71, 1) = &H103C
    WinZawgyi(72, 1) = &H103D
    WinZawgyi(73, 1) = &H108A
    WinZawgyi(74, 1) = &H104C
    WinZawgyi(75, 1) = &H104D
    WinZawgyi(76, 1) = &H104E
    WinZawgyi(77, 1) = &H104F
    WinZawgyi(78, 1) = &H1060
    WinZawgyi(79, 1) = &H1061
    WinZawgyi(80, 1) = &H1062
    WinZawgyi(81, 1) = &H1063
    WinZawgyi(82, 1) = &H1065
    WinZawgyi(83, 1) = &H1066
    WinZawgyi(84, 1) = &H1067
    WinZawgyi(85, 1) = &H1069
    WinZawgyi(86, 1) = &H106B
    WinZawgyi(87, 1) = &H106A
    WinZawgyi(88, 1) = &H1097
    WinZawgyi(89, 1) = &H106C
    WinZawgyi(90, 1) = &H106D
    WinZawgyi(91, 1) = &H1092
    WinZawgyi(92, 1) = &H106E
    WinZawgyi(93, 1) = &H1091
    WinZawgyi(94, 1) = &H106F
    WinZawgyi(95, 1) = &H1070
    WinZawgyi(96, 1) = &H1072
    WinZawgyi(97, 1) = &H1071
    WinZawgyi(98, 1) = &H1073
    WinZawgyi(99, 1) = &H1074
    WinZawgyi(100, 1) = &H1075
    WinZawgyi(101, 1) = &H1076
    WinZawgyi(102, 1) = &H108F
    WinZawgyi(103, 1) = &H1077
    WinZawgyi(104, 1) = &H1078
    WinZawgyi(105, 1) = &H1079
    WinZawgyi(106, 1) = &H107A
    WinZawgyi(107, 1) = &H107B
    WinZawgyi(108, 1) = &H107C
    WinZawgyi(109, 1) = &H1090
    WinZawgyi(110, 1) = &H1086
    WinZawgyi(111, 1) = &H1087
    WinZawgyi(112, 1) = &H1064
    WinZawgyi(113, 1) = &H108C
    WinZawgyi(114, 1) = &H108B
    WinZawgyi(115, 1) = &H108D
    WinZawgyi(116, 1) = &H108E
    WinZawgyi(117, 1) = &H1096
    WinZawgyi(118, 1) = &H107D
    WinZawgyi(118, 1) = &H107D
    WinZawgyi(119, 1) = &H1039
    WinZawgyi(120, 1) = &H104B
    WinZawgyi(121, 1) = &H104A
    WinZawgyi(122, 1) = &H2F
    WinZawgyi(123, 1) = &H2666
    WinZawgyi(124, 1) = &H66D
        
    ReDim WinZawgyiPhrases(0 To 16, 0 To 1) As String
    
    WinZawgyiPhrases(0, 0) = Chr(77) & Chr(111)
    WinZawgyiPhrases(0, 1) = ChrW(&H1029)
    WinZawgyiPhrases(1, 0) = Chr(97) & Chr(77) & Chr(111) & Chr(109) & Chr(102)
    WinZawgyiPhrases(1, 1) = ChrW(&H102A)
    WinZawgyiPhrases(2, 0) = Chr(52) & Chr(105) & Chr(102) & Chr(59)
    WinZawgyiPhrases(2, 1) = ChrW(&H104E)
    WinZawgyiPhrases(3, 0) = Chr(164) & Chr(105) & Chr(102) & Chr(59)
    WinZawgyiPhrases(3, 1) = ChrW(&H104E)
    WinZawgyiPhrases(4, 0) = Chr(93) & Chr(93)
    WinZawgyiPhrases(4, 1) = ChrW(&H22)
    WinZawgyiPhrases(5, 0) = Chr(93)
    WinZawgyiPhrases(5, 1) = ChrW(&H27)
    WinZawgyiPhrases(6, 0) = Chr(125) & Chr(125)
    WinZawgyiPhrases(6, 1) = ChrW(&H22)
    WinZawgyiPhrases(7, 0) = Chr(125)
    WinZawgyiPhrases(7, 1) = ChrW(&H27)
    WinZawgyiPhrases(8, 0) = "<u"
    WinZawgyiPhrases(8, 1) = ChrW(&H1082) & ChrW(&H1000) & ChrW(&H103C)
    WinZawgyiPhrases(9, 0) = "<x"
    WinZawgyiPhrases(9, 1) = ChrW(&H1082) & ChrW(&H1010) & ChrW(&H103C)
    WinZawgyiPhrases(10, 0) = Chr(62) & Chr(99)
    WinZawgyiPhrases(10, 1) = ChrW(&H1081) & ChrW(&H1001) & ChrW(&H103C)
    WinZawgyiPhrases(11, 0) = Chr(62) & Chr(114)
    WinZawgyiPhrases(11, 1) = ChrW(&H1081) & ChrW(&H1019) & ChrW(&H103C)
    WinZawgyiPhrases(12, 0) = "ps"
    WinZawgyiPhrases(12, 1) = ChrW(&H1008)
    WinZawgyiPhrases(13, 0) = Chr(82)
    WinZawgyiPhrases(13, 1) = ChrW(&H103C) & ChrW(&H107D)
    WinZawgyiPhrases(14, 0) = Chr(81)
    WinZawgyiPhrases(14, 1) = ChrW(&H103D) & ChrW(&H103A)
    WinZawgyiPhrases(15, 0) = Chr(87)
    WinZawgyiPhrases(15, 1) = ChrW(&H108A) & ChrW(&H107D)
    WinZawgyiPhrases(16, 0) = Chr(211)
    WinZawgyiPhrases(16, 1) = ChrW(&H1009) & ChrW(&H102C)
    
                
    With Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
        .Font.Name = "-Win---Researcher"
        .Replacement.Font.Name = "Zawgyi-One"
    End With
    
    For i = 0 To UBound(WinZawgyiPhrases)
        Selection.Find.Text = WinZawgyiPhrases(i, 0)
        Selection.Find.Replacement.Text = WinZawgyiPhrases(i, 1)
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
    
    For i = 0 To UBound(WinZawgyi)
        Selection.Find.Text = Chr(WinZawgyi(i, 0))
        Selection.Find.Replacement.Text = ChrW(WinZawgyi(i, 1))
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
    
    Selection.WholeStory
    Selection.Font.Size = 10
    
    Selection.WholeStory
    Selection.Cut
    
    Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    
    Selection.PasteAndFormat (wdPasteDefault)
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    
End Sub

