#include "..\AutoItTest.au3"
#include "..\..\utility\ArrayUtility.au3"

#region Constant_Define
Const $two_dimensional_array[5][4] = [ _
		["1-1", "1-2", "1-3", "1-4"], _
		["2-1", "2-2", "2-3", "2-4"], _
		["3-1", "3-2", "3-3", "3-4"], _
		["4-1", "4-2", "4-3", "4-4"], _
		["5-1", "5-2", "5-3", "5-4"] _
		]
Const $ArrayUtility_ExtractionRow_Test_Answer[4] = ["1-1", "1-2", "1-3", "1-4"]
Const $ArrayUtility_ExtractionColumn_Test_Answer[5] = ["1-1", "2-1", "3-1", "4-1", "5-1"]
#endregion Constant_Define

Local $ArrayUtilityTest[2][5] = [ _
		["", "ArrayUtility_ExtractionRow_Test", "AutoItTest_AssertArrayEquals", $ArrayUtility_ExtractionRow_Test_Answer, ""], _
		["", "ArrayUtility_ExtractionColumn_Test", "AutoItTest_AssertArrayEquals", $ArrayUtility_ExtractionColumn_Test_Answer, ""] _
		]
AutoItTest_Runner($ArrayUtilityTest)

#region ArrayUtility_ExtractionRow_Test
Func ArrayUtility_ExtractionRow_Test()
	Return ArrayUtility_ExtractionRow($two_dimensional_array, 0)
EndFunc   ;==>ArrayUtility_ExtractionRow_Test
#endregion ArrayUtility_ExtractionRow_Test

#region ArrayUtility_ExtractionColumn_Test
Func ArrayUtility_ExtractionColumn_Test()
	Return ArrayUtility_ExtractionColumn($two_dimensional_array, 0)
EndFunc   ;==>ArrayUtility_ExtractionColumn_Test
#endregion ArrayUtility_ExtractionColumn_Test