# faniry-robot

OA Robot command collection with Faniry's LAMBDAs and VBA macros.

## Install

1. In Excel, open the OA Robot palette (`Ctrl+Shift+O`).
2. Type `+@` and add this file as a synchable GitHub command collection: `https://github.com/Faniry14/faniry-robot/blob/main/FaniryBot.xlsm` (the link to the file, not to the repository).
3. Open the case workbook and run **Import My Lambdas** (`FIL`).

## Editing

Keep your working copy under a different name (for example `Lambda.xlsm`): OA Robot keeps `FaniryBot.xlsm` open in the background, and Excel cannot open two workbooks with the same name. To publish a new version, copy your working file over `FaniryBot.xlsm` in File Explorer (not with Save As), then commit and push.

## Commands

| Code | Command | What it does |
|---|---|---|
| `FIL` | Import My Lambdas | Copies all LAMBDAs of this collection into the active workbook. Same-named lambdas are updated; names the case already uses for something else are left alone. |
| `FLS` | Create Level Sheet | Creates a LevelN sheet from the selected row (XLOOKUP on B2). |
| `FDT` | Data Table VBA Simulation | Runs each value of B5:B into B2 and writes C4's result in column C. |
| `FCC` | Turn Selected Cells Into Color Code | Writes each selected cell's fill color code into the cell. |
| `FRC` | Replace Based On Color | With two areas selected, fills area 1 with the values of area 2 that share the same color. |
| `FRGB` | RGB Color | Writes the RGB code of each selected cell's fill. |
| `FHEX` | RGB Color In Hex | Writes the hex code of each selected cell's fill. |
| `FXC` | XLookup Color | Looks up values by fill color between two ranges. |
| `FUC` | Unique Colors RGB Hex | Lists the unique fill colors of a range with RGB, hex and count. |

## LAMBDAs (213)

`_addresses`, `_cardnum`, `_countOcc`, `_distfrom`, `_MostOf`, `_neighborOp`, `_next`, `_pip`, `_PNext`, `_Pprev`, `_prev`, `Ad_To_Contents`, `Add_to_RC_ref`, `Add2Kyear`, `Addres`, `AddressOfCells`, `AddressOfRectangle`, `ALevel`, `Alphabet`, `ArrayShift`, `ArrLookup`, `Arrow_to_RC_ref`, `Arrows`, `Arwscomb`, `Bcol`, `Best_Path`, `Between`, `BiCol`, `Biggest`, `BigMath`, `BiRow`, `BlockOutMap`, `Bmap`, `BoardGameMove`, `Bround`, `Brow`, `BscanC`, `BscanR`, `BscanT`, `ByAll`, `ByRectangle`, `Candy_Crush`, `CardDetails`, `Cartesian_Product`, `Case_Lookup`, `CaseQuestions`, `ChessMoves`, `ChessNext`, `ChessPiece`, `Cipher`, `Closest_RC_ref`, `ClosestOnMap`, `ClosestOnMapPro`, `cm_to_inches`, `Col_Letter`, `Coll_lab`, `ColNum`, `Combinations`, `Consecutive`, `Consecutives`, `Contain`, `Content_to_add`, `ContentOnMap`, `ConversionTable`, `Count_Word`, `CountOcc`, `CurrentRegion`, `Cyclic_Gap`, `Dart`, `Devise_conversion`, `Diag_to_Row`, `Dice`, `Dir_To_End`, `Dist`, `Distance`, `Div`, `Exists`, `Extract_emojis`, `ExtractWords`, `Fall`, `Find_In_Map`, `FindMatches`, `Flat_Map`, `Flatten`, `Flip_letters`, `FlipHV`, `FordFulkerson`, `French`, `Gby`, `HEXA_to_RGB`, `Hungarian`, `Idirecrt`, `inches_to_cm`, `Is_palindromic`, `ISBETWEEN`, `IsMultiple`, `ISPRIME`, `IsPrimeBetween`, `kmh_to_ms`, `Kth_Value`, `LeagueStat`, `Level_Lookup`, `LevenshteinDist`, `Line_in_map`, `LIST.COMBIN`, `LIST.PERMUTATIONA`, `Longaddition`, `Map_dijkstra`, `map_to_RC_ref`, `MapBlocker`, `MapGroup`, `MapRC`, `MapV`, `Mark_Map`, `Max_Lookup`, `Maze`, `Maze_old`, `Min_lookup`, `MinCoins`, `MinDisFrom`, `Money`, `Month_Trad`, `MostOf`, `Mov`, `Move_On_Map_Array`, `MoveOnMap`, `ms_to_kmh`, `Mulg`, `Multiset_Combinations`, `N_of_kind`, `Neighbor`, `NetWorkHours`, `NT.BigMod`, `NT.Collatz`, `NT.CountWithDigit`, `NT.Digits`, `NT.Divisors`, `NT.Fib`, `NT.FirstPal`, `NT.InSeries`, `NT.MissingDigit`, `NT.MostDivisors`, `NT.Prime`, `NT.SameDiv`, `NT.Sk`, `NT.SumMultiples`, `packCol_helper`, `Palindrome`, `Pascal_Triangle`, `Percom`, `Plus`, `PokerHand`, `POWERLIFT`, `Prev`, `Primes`, `PriorityMovement`, `RC_Ref_Convert`, `RC_ref_to_Add`, `Rectangle_Possible_to_row`, `RegexFromExample`, `RemoveChar`, `RemoveCharacters`, `Repeat_array`, `Repeat_Elements`, `Replace_Block`, `ResolveBoard`, `Reverse`, `reverse_array_2D`, `RGB_to_HEX`, `Rlookup`, `RoCo`, `Rotate90`, `RotateC`, `Row_lab`, `RowNum`, `runMask_helper`, `RwClNm`, `Scan_RC`, `sequence_expand_range`, `sheet_content`, `Sheet_name`, `Shift_text`, `ShiftText`, `ShortPath`, `Snake`, `SnakeG`, `solveBoard_helper`, `SpellNo`, `Spiral`, `Split_add`, `SplitCards`, `SplitText`, `Sudoku`, `Suite_card`, `SwapOnMap`, `Take_diagonals`, `Take_Text`, `TEXTBETWEEN`, `TextPT`, `TimerN`, `Tlookup`, `To_ref`, `TopCombos`, `Tslate`, `UnitCombine`, `UnitSepG`, `Unpivot`, `VStacks`, `Week_Day_Trad`, `WeekMonth`, `winEq_helper`, `WordMaker`, `WordValue`

## Credits

Some lambdas are adapted from the Excel community (Excel Robot, MEWC competitors and others), and some were written with the help of Claude (Anthropic).
