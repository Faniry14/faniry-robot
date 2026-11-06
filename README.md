# faniry-robot

OA Robot command collection with Faniry's LAMBDAs and VBA macros.

## Install

1. In Excel, open the OA Robot palette (`Ctrl+Shift+O`).
2. Type `+@` and add this file as a synchable GitHub command collection: `https://github.com/Faniry14/faniry-robot/blob/main/Lambda.xlsm` (the link to the file, not to the repository).
3. Open the case workbook and run **Import My Lambdas** (`FIL`).

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

## LAMBDAs (198)

`_addresses`, `_cardnum`, `_countOcc`, `_distfrom`, `_MostOf`, `_neighborOp`, `_next`, `_packCol`, `_pip`, `_PNext`, `_Pprev`, `_prev`, `_runMask`, `_solveBoard`, `_split`, `_winEq`, `Ad_To_Contents`, `Add_to_RC_ref`, `Add2Kyear`, `Addres`, `AddressOfCells`, `ALevel`, `Alphabet`, `ArrayShift`, `Arrow_to_RC_ref`, `Arrows`, `Arwscomb`, `Bcol`, `Best_Path`, `Between`, `BiCol`, `Biggest`, `BiRow`, `BlockOutMap`, `Bmap`, `BoardGameMove`, `Bround`, `Brow`, `BscanC`, `BscanR`, `BscanT`, `ByAll`, `Candy_Crush`, `CardDetails`, `Case_Lookup`, `CaseQuestions`, `ChessMoves`, `ChessNext`, `ChessPiece`, `Cipher`, `Closest_RC_ref`, `ClosestOnMap`, `ClosestOnMapPro`, `cm_to_inches`, `Col_Letter`, `Coll_lab`, `ColNum`, `Combinations`, `Consecutive`, `Consecutives`, `Contain`, `Content_to_add`, `ContentOnMap`, `ConversionTable`, `Count_Word`, `CountOcc`, `CurrentRegion`, `Dart`, `Diag_to_Row`, `Dice`, `Dir_To_End`, `Dist`, `Distance`, `Div`, `Exists`, `Extract_emojis`, `ExtractNumbers`, `ExtractWords`, `Fall`, `Find_In_Map`, `FindMatches`, `Flat_Map`, `Flatten`, `Flip_letters`, `FlipHV`, `FordFulkerson`, `French`, `Gby`, `HEXA_to_RGB`, `Hungarian`, `iColor`, `Idirecrt`, `inches_to_cm`, `Is_palindromic`, `ISBETWEEN`, `IsMultiple`, `ISPRIME`, `IsPrimeBetween`, `kmh_to_ms`, `LeagueStat`, `Level_Lookup`, `LevenshteinDist`, `Line_in_map`, `LIST.COMBIN`, `LIST.PERMUTATIONA`, `Longaddition`, `Map_dijkstra`, `map_to_RC_ref`, `MapBlocker`, `MapGroup`, `MapRC`, `MapV`, `Mark_Map`, `Max_Lookup`, `Maze`, `Maze_Map`, `MazeDistance`, `Min_lookup`, `MinCoins`, `MinDis2`, `MinDisFrom`, `Money`, `MostOf`, `Mov`, `MoveOnMap`, `MovPath`, `MovPathC`, `MovPathRC`, `MovPathZp`, `ms_to_kmh`, `Mulg`, `N_of_kind`, `Neighbor`, `NetWorkHours`, `Palindrome`, `Pascal_Triangle`, `Percom`, `Plus`, `PokerHand`, `POWERLIFT`, `Prev`, `Primes`, `PriorityMovement`, `RC_Ref_Convert`, `RC_ref_to_Add`, `Rectangle_Possible_to_row`, `RegexFromExample`, `RemoveChar`, `RemoveCharacters`, `Repeat_array`, `ResolveBoard`, `Reverse`, `reverse_array_2D`, `RGB_to_HEX`, `RoCo`, `Rotate90`, `RotateC`, `Row_lab`, `RowNum`, `RwClNm`, `Scan_RC`, `sequence_expand_range`, `Sheet_name`, `Shift_text`, `ShiftText`, `ShortPath`, `Snake`, `SnakeG`, `SpellNo`, `Spiral`, `Split_add`, `Split_upper`, `SplitAll`, `SplitArray`, `SplitCards`, `SplitCombo`, `SplitText`, `Sudoku`, `Suite_card`, `SwapOnMap`, `Take_diagonals`, `Take_Digits`, `Take_LowCase_Letter`, `take_number`, `Take_UpCase_Letter`, `TEXTBETWEEN`, `TextPT`, `TimerN`, `Tlookup`, `To_ref`, `TopCombos`, `Tslate`, `UnitCombine`, `UnitSepG`, `Unpivot`, `VStacks`, `WordMaker`, `WordValue`

## Credits

Some lambdas are adapted from the Excel community (Excel Robot, MEWC competitors and others), and some were written with the help of Claude (Anthropic).
