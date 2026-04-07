Attribute VB_Name = "mデータクリア_集計表"
Sub 集計表データクリアマクロ()

    ' 対象範囲を結合して内容をクリア
    Union(Range("D46:D53"), Range("F4:F8,F10,F12:F16,F18,F20:F22,F24,F26:F29,F31,F33:F35,F37,F39:F42,F46:F53"), _
          Range("H4:H8,H10,H12:H16,H18,H20:H22,H24,H26:H29,H33:H35,H37,H39:H42,H46:H53"), _
          Range("J4:J8,J10,J12:J16,J33:J35,J37,J39:J42,J46:J53"), _
          Range("L18,L31,L46:L53"), _
          Range("M33,M35"), _
          Range("N4:N8,N10,N12:N16,N35,N37,N40,N46:N53"), _
          Range("P33:P35,P37,P39:P42,P46:P53"), _
          Range("T4:T9,T11:T16,T20:T25,T30:T35,T37:T42")).ClearContents

End Sub
