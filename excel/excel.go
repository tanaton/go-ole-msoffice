package excel

import (
	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"time"
	"unsafe"
)

const (
	Xl24HourClock                            = 33
	Xl3DArea                                 = -4098
	Xl3DAreaStacked                          = 78
	Xl3DAreaStacked100                       = 79
	Xl3DBarClustered                         = 60
	Xl3DBarStacked                           = 61
	Xl3DBarStacked100                        = 62
	Xl3DColumn                               = -4100
	Xl3DColumnClustered                      = 54
	Xl3DColumnStacked                        = 55
	Xl3DColumnStacked100                     = 56
	Xl3DLine                                 = -4101
	Xl3DPie                                  = -4102
	Xl3DPieExploded                          = 70
	Xl4DigitYears                            = 43
	XlA1                                     = 1
	XlADORecordset                           = 7
	XlAboveAverageCondition                  = 12
	XlAbsRowRelColumn                        = 2
	XlAbsolute                               = 1
	XlAddIn8                                 = 18
	XlAllAtOnce                              = 2
	XlAllChanges                             = 2
	XlAllFaces                               = 7
	XlAllTables                              = 2
	XlAlternateArraySeparator                = 16
	XlAlways                                 = 1
	XlAnd                                    = 1
	XlAnyGallery                             = 23
	XlAnyKey                                 = 2
	XlArabicBothStrict                       = 3
	XlArabicNone                             = 0
	XlArabicStrictAlefHamza                  = 1
	XlArabicStrictFinalYaa                   = 2
	XlArea                                   = 1
	XlAreaStacked                            = 76
	XlAreaStacked100                         = 77
	XlArrangeStyleCascade                    = 7
	XlArrangeStyleHorizontal                 = -4128
	XlArrangeStyleTiled                      = 1
	XlArrangeStyleVertical                   = -4166
	XlArrowHeadLengthLong                    = 3
	XlArrowHeadLengthMedium                  = -4138
	XlArrowHeadLengthShort                   = 1
	XlArrowHeadStyleClosed                   = 3
	XlArrowHeadStyleDoubleClosed             = 5
	XlArrowHeadStyleDoubleOpen               = 4
	XlArrowHeadStyleNone                     = -4142
	XlArrowHeadStyleOpen                     = 2
	XlArrowHeadWidthMedium                   = -4138
	XlArrowHeadWidthNarrow                   = 1
	XlArrowHeadWidthWide                     = 3
	XlAsRequired                             = 0
	XlAscending                              = 1
	XlAtBottom                               = 2
	XlAtTop                                  = 1
	XlAutoActivate                           = 3
	XlAutoClose                              = 2
	XlAutoDeactivate                         = 4
	XlAutoFill                               = 4
	XlAutoOpen                               = 1
	XlAutomaticScale                         = -4105
	XlAutomaticUpdate                        = 4
	XlAverage                                = -4106
	XlAxis                                   = 21
	XlAxisCrossesAutomatic                   = -4105
	XlAxisCrossesCustom                      = -4114
	XlAxisCrossesMaximum                     = 2
	XlAxisCrossesMinimum                     = 4
	XlAxisTitle                              = 17
	XlBIFF                                   = 2
	XlBMP                                    = 1
	XlBackgroundAutomatic                    = -4105
	XlBackgroundOpaque                       = 3
	XlBackgroundTransparent                  = 2
	XlBarClustered                           = 57
	XlBarOfPie                               = 71
	XlBarStacked                             = 58
	XlBarStacked100                          = 59
	XlBetween                                = 1
	XlBitmap                                 = 2
	XlBlanks                                 = 4
	XlBlanksCondition                        = 10
	XlBottom10Items                          = 4
	XlBottom10Percent                        = 6
	XlBox                                    = 0
	XlBubble                                 = 15
	XlBubble3DEffect                         = 87
	XlBuiltIn                                = 21
	XlButton                                 = 15
	XlButtonControl                          = 0
	XlButtonOnly                             = 2
	XlByColumns                              = 2
	XlByRows                                 = 1
	XlCGM                                    = 7
	XlCSV                                    = 6
	XlCSVMSDOS                               = 24
	XlCSVMac                                 = 22
	XlCSVWindows                             = 23
	XlCalculatedMeasure                      = 2
	XlCalculatedMember                       = 0
	XlCalculatedSet                          = 1
	XlCalculating                            = 1
	XlCalculationAutomatic                   = -4105
	XlCalculationManual                      = -4135
	XlCalculationSemiautomatic               = 2
	XlCancel                                 = 1
	XlCap                                    = 1
	XlCategory                               = 1
	XlCategoryScale                          = 2
	XlCellTypeAllFormatConditions            = -4172
	XlCellTypeAllValidation                  = -4174
	XlCellTypeBlanks                         = 4
	XlCellTypeComments                       = -4144
	XlCellTypeConstants                      = 2
	XlCellTypeFormulas                       = -4123
	XlCellTypeLastCell                       = 11
	XlCellTypeSameFormatConditions           = -4173
	XlCellTypeSameValidation                 = -4175
	XlCellTypeVisible                        = 12
	XlCellValue                              = 1
	XlChangeAttributes                       = 6
	XlChart                                  = -4109
	XlChartArea                              = 2
	XlChartAsWindow                          = 5
	XlChartElementPositionAutomatic          = -4105
	XlChartElementPositionCustom             = -4114
	XlChartInPlace                           = 4
	XlChartTitle                             = 4
	XlCheckBox                               = 1
	XlChronological                          = 3
	XlClipboard                              = 3
	XlClipboardFormatBIFF                    = 8
	XlClipboardFormatBIFF12                  = 63
	XlClipboardFormatBIFF2                   = 18
	XlClipboardFormatBIFF3                   = 20
	XlClipboardFormatBIFF4                   = 30
	XlClipboardFormatBinary                  = 15
	XlClipboardFormatBitmap                  = 9
	XlClipboardFormatCGM                     = 13
	XlClipboardFormatCSV                     = 5
	XlClipboardFormatDIF                     = 4
	XlClipboardFormatDspText                 = 12
	XlClipboardFormatEmbedSource             = 22
	XlClipboardFormatEmbeddedObject          = 21
	XlClipboardFormatLink                    = 11
	XlClipboardFormatLinkSource              = 23
	XlClipboardFormatLinkSourceDesc          = 32
	XlClipboardFormatMovie                   = 24
	XlClipboardFormatNative                  = 14
	XlClipboardFormatObjectDesc              = 31
	XlClipboardFormatObjectLink              = 19
	XlClipboardFormatOwnerLink               = 17
	XlClipboardFormatPICT                    = 2
	XlClipboardFormatPrintPICT               = 3
	XlClipboardFormatRTF                     = 7
	XlClipboardFormatSYLK                    = 6
	XlClipboardFormatScreenPICT              = 29
	XlClipboardFormatStandardFont            = 28
	XlClipboardFormatStandardScale           = 27
	XlClipboardFormatTable                   = 16
	XlClipboardFormatText                    = 0
	XlClipboardFormatToolFace                = 25
	XlClipboardFormatToolFacePICT            = 26
	XlClipboardFormatVALU                    = 1
	XlClipboardFormatWK1                     = 10
	XlCmdCube                                = 1
	XlCmdDAX                                 = 8
	XlCmdDefault                             = 4
	XlCmdExcel                               = 7
	XlCmdList                                = 5
	XlCmdSql                                 = 2
	XlCmdTable                               = 3
	XlCmdTableCollection                     = 6
	XlCodePage                               = 2
	XlColorIndexAutomatic                    = -4105
	XlColorIndexNone                         = -4142
	XlColorScale                             = 3
	XlColumnClustered                        = 51
	XlColumnField                            = 2
	XlColumnHeader                           = -4110
	XlColumnItem                             = 5
	XlColumnLabels                           = 2
	XlColumnSeparator                        = 14
	XlColumnStacked                          = 52
	XlColumnStacked100                       = 53
	XlColumnThenRow                          = 2
	XlColumns                                = 2
	XlCommand                                = 2
	XlCommandUnderlinesAutomatic             = -4105
	XlCommandUnderlinesOff                   = -4146
	XlCommandUnderlinesOn                    = 1
	XlCommentAndIndicator                    = 1
	XlCommentIndicatorOnly                   = -1
	XlComments                               = -4144
	XlConeBarClustered                       = 102
	XlConeBarStacked                         = 103
	XlConeBarStacked100                      = 104
	XlConeCol                                = 105
	XlConeColClustered                       = 99
	XlConeColStacked                         = 100
	XlConeColStacked100                      = 101
	XlConeToMax                              = 5
	XlConeToPoint                            = 4
	XlConsolidation                          = 3
	XlConstant                               = 1
	XlContinuous                             = 1
	XlCopy                                   = 1
	XlCorners                                = 6
	XlCount                                  = -4112
	XlCountNums                              = -4113
	XlCountryCode                            = 1
	XlCountrySetting                         = 2
	XlCreatorCode                            = 1480803660
	XlCurrencyBefore                         = 37
	XlCurrencyCode                           = 25
	XlCurrencyDigits                         = 27
	XlCurrencyLeadingZeros                   = 40
	XlCurrencyMinusSign                      = 38
	XlCurrencyNegative                       = 28
	XlCurrencySpaceBefore                    = 36
	XlCurrencyTrailingZeros                  = 39
	XlCurrentPlatformText                    = -4158
	XlCut                                    = 2
	XlCylinder                               = 3
	XlCylinderBarClustered                   = 95
	XlCylinderBarStacked                     = 96
	XlCylinderBarStacked100                  = 97
	XlCylinderCol                            = 98
	XlCylinderColClustered                   = 92
	XlCylinderColStacked                     = 93
	XlCylinderColStacked100                  = 94
	XlDAORecordset                           = 2
	XlDBF2                                   = 7
	XlDBF3                                   = 8
	XlDBF4                                   = 11
	XlDIF                                    = 9
	XlDMYFormat                              = 4
	XlDRW                                    = 4
	XlDXF                                    = 5
	XlDYMFormat                              = 7
	XlDash                                   = -4115
	XlDashDot                                = 4
	XlDashDotDot                             = 5
	XlDataAndLabel                           = 0
	XlDataField                              = 4
	XlDataHeader                             = 3
	XlDataItem                               = 7
	XlDataLabel                              = 0
	XlDataLabelSeparatorDefault              = 1
	XlDataLabelsShowBubbleSizes              = 6
	XlDataLabelsShowLabel                    = 4
	XlDataLabelsShowLabelAndPercent          = 5
	XlDataLabelsShowNone                     = -4142
	XlDataLabelsShowPercent                  = 3
	XlDataLabelsShowValue                    = 2
	XlDataOnly                               = 2
	XlDataSeriesLinear                       = -4132
	XlDataTable                              = 7
	XlDatabar                                = 4
	XlDatabase                               = 1
	XlDate                                   = 2
	XlDateOrder                              = 32
	XlDateSeparator                          = 17
	XlDay                                    = 1
	XlDayCode                                = 21
	XlDayLeadingZero                         = 42
	XlDays                                   = 0
	XlDecimalSeparator                       = 3
	XlDefault                                = -4143
	XlDelimited                              = 1
	XlDescending                             = 2
	XlDiagonalDown                           = 5
	XlDiagonalUp                             = 6
	XlDialogActivate                         = 103
	XlDialogActiveCellFont                   = 476
	XlDialogAddChartAutoformat               = 390
	XlDialogAddinManager                     = 321
	XlDialogAlignment                        = 43
	XlDialogAppMove                          = 170
	XlDialogAppSize                          = 171
	XlDialogApplyNames                       = 133
	XlDialogApplyStyle                       = 212
	XlDialogArrangeAll                       = 12
	XlDialogAssignToObject                   = 213
	XlDialogAssignToTool                     = 293
	XlDialogAttachText                       = 80
	XlDialogAttachToolbars                   = 323
	XlDialogAutoCorrect                      = 485
	XlDialogAxes                             = 78
	XlDialogBorder                           = 45
	XlDialogCalculation                      = 32
	XlDialogCellProtection                   = 46
	XlDialogChangeLink                       = 166
	XlDialogChartAddData                     = 392
	XlDialogChartLocation                    = 527
	XlDialogChartOptionsDataLabelMultiple    = 724
	XlDialogChartOptionsDataLabels           = 505
	XlDialogChartOptionsDataTable            = 506
	XlDialogChartSourceData                  = 540
	XlDialogChartTrend                       = 350
	XlDialogChartType                        = 526
	XlDialogChartWizard                      = 288
	XlDialogCheckboxProperties               = 435
	XlDialogClear                            = 52
	XlDialogColorPalette                     = 161
	XlDialogColumnWidth                      = 47
	XlDialogCombination                      = 73
	XlDialogConditionalFormatting            = 583
	XlDialogConsolidate                      = 191
	XlDialogCopyChart                        = 147
	XlDialogCopyPicture                      = 108
	XlDialogCreateList                       = 796
	XlDialogCreateNames                      = 62
	XlDialogCreatePublisher                  = 217
	XlDialogCreateRelationship               = 1272
	XlDialogCustomViews                      = 493
	XlDialogCustomizeToolbar                 = 276
	XlDialogDataDelete                       = 36
	XlDialogDataLabel                        = 379
	XlDialogDataLabelMultiple                = 723
	XlDialogDataSeries                       = 40
	XlDialogDataValidation                   = 525
	XlDialogDefineName                       = 61
	XlDialogDefineStyle                      = 229
	XlDialogDeleteFormat                     = 111
	XlDialogDeleteName                       = 110
	XlDialogDemote                           = 203
	XlDialogDisplay                          = 27
	XlDialogDocumentInspector                = 862
	XlDialogEditColor                        = 223
	XlDialogEditDelete                       = 54
	XlDialogEditSeries                       = 228
	XlDialogEditboxProperties                = 438
	XlDialogEditionOptions                   = 251
	XlDialogErrorChecking                    = 732
	XlDialogErrorbarX                        = 463
	XlDialogErrorbarY                        = 464
	XlDialogEvaluateFormula                  = 709
	XlDialogExternalDataProperties           = 530
	XlDialogExtract                          = 35
	XlDialogFileDelete                       = 6
	XlDialogFileSharing                      = 481
	XlDialogFillGroup                        = 200
	XlDialogFillWorkgroup                    = 301
	XlDialogFilter                           = 447
	XlDialogFilterAdvanced                   = 370
	XlDialogFindFile                         = 475
	XlDialogFont                             = 26
	XlDialogFontProperties                   = 381
	XlDialogFormatAuto                       = 269
	XlDialogFormatChart                      = 465
	XlDialogFormatCharttype                  = 423
	XlDialogFormatFont                       = 150
	XlDialogFormatLegend                     = 88
	XlDialogFormatMain                       = 225
	XlDialogFormatMove                       = 128
	XlDialogFormatNumber                     = 42
	XlDialogFormatOverlay                    = 226
	XlDialogFormatSize                       = 129
	XlDialogFormatText                       = 89
	XlDialogFormulaFind                      = 64
	XlDialogFormulaGoto                      = 63
	XlDialogFormulaReplace                   = 130
	XlDialogFunctionWizard                   = 450
	XlDialogGallery3dArea                    = 193
	XlDialogGallery3dBar                     = 272
	XlDialogGallery3dColumn                  = 194
	XlDialogGallery3dLine                    = 195
	XlDialogGallery3dPie                     = 196
	XlDialogGallery3dSurface                 = 273
	XlDialogGalleryArea                      = 67
	XlDialogGalleryBar                       = 68
	XlDialogGalleryColumn                    = 69
	XlDialogGalleryCustom                    = 388
	XlDialogGalleryDoughnut                  = 344
	XlDialogGalleryLine                      = 70
	XlDialogGalleryPie                       = 71
	XlDialogGalleryRadar                     = 249
	XlDialogGalleryScatter                   = 72
	XlDialogGoalSeek                         = 198
	XlDialogGridlines                        = 76
	XlDialogImportTextFile                   = 666
	XlDialogInsert                           = 55
	XlDialogInsertHyperlink                  = 596
	XlDialogInsertNameLabel                  = 496
	XlDialogInsertObject                     = 259
	XlDialogInsertPicture                    = 342
	XlDialogInsertTitle                      = 380
	XlDialogLabelProperties                  = 436
	XlDialogListboxProperties                = 437
	XlDialogMacroOptions                     = 382
	XlDialogMailEditMailer                   = 470
	XlDialogMailLogon                        = 339
	XlDialogMailNextLetter                   = 378
	XlDialogMainChart                        = 85
	XlDialogMainChartType                    = 185
	XlDialogManageRelationships              = 1271
	XlDialogMenuEditor                       = 322
	XlDialogMove                             = 262
	XlDialogMyPermission                     = 834
	XlDialogNameManager                      = 977
	XlDialogNew                              = 119
	XlDialogNewName                          = 978
	XlDialogNewWebQuery                      = 667
	XlDialogNote                             = 154
	XlDialogObjectProperties                 = 207
	XlDialogObjectProtection                 = 214
	XlDialogOpen                             = 1
	XlDialogOpenLinks                        = 2
	XlDialogOpenMail                         = 188
	XlDialogOpenText                         = 441
	XlDialogOptionsCalculation               = 318
	XlDialogOptionsChart                     = 325
	XlDialogOptionsEdit                      = 319
	XlDialogOptionsGeneral                   = 356
	XlDialogOptionsListsAdd                  = 458
	XlDialogOptionsME                        = 647
	XlDialogOptionsTransition                = 355
	XlDialogOptionsView                      = 320
	XlDialogOutline                          = 142
	XlDialogOverlay                          = 86
	XlDialogOverlayChartType                 = 186
	XlDialogPageSetup                        = 7
	XlDialogParse                            = 91
	XlDialogPasteNames                       = 58
	XlDialogPasteSpecial                     = 53
	XlDialogPatterns                         = 84
	XlDialogPermission                       = 832
	XlDialogPhonetic                         = 656
	XlDialogPivotCalculatedField             = 570
	XlDialogPivotCalculatedItem              = 572
	XlDialogPivotClientServerSet             = 689
	XlDialogPivotFieldGroup                  = 433
	XlDialogPivotFieldProperties             = 313
	XlDialogPivotFieldUngroup                = 434
	XlDialogPivotShowPages                   = 421
	XlDialogPivotSolveOrder                  = 568
	XlDialogPivotTableOptions                = 567
	XlDialogPivotTableSlicerConnections      = 1183
	XlDialogPivotTableWhatIfAnalysisSettings = 1153
	XlDialogPivotTableWizard                 = 312
	XlDialogPlacement                        = 300
	XlDialogPrint                            = 8
	XlDialogPrintPreview                     = 222
	XlDialogPrinterSetup                     = 9
	XlDialogPromote                          = 202
	XlDialogProperties                       = 474
	XlDialogPropertyFields                   = 754
	XlDialogProtectDocument                  = 28
	XlDialogProtectSharing                   = 620
	XlDialogPublishAsWebPage                 = 653
	XlDialogPushbuttonProperties             = 445
	XlDialogRecommendedPivotTables           = 1258
	XlDialogReplaceFont                      = 134
	XlDialogRoutingSlip                      = 336
	XlDialogRowHeight                        = 127
	XlDialogRun                              = 17
	XlDialogSaveAs                           = 5
	XlDialogSaveCopyAs                       = 456
	XlDialogSaveNewObject                    = 208
	XlDialogSaveWorkbook                     = 145
	XlDialogSaveWorkspace                    = 285
	XlDialogScale                            = 87
	XlDialogScenarioAdd                      = 307
	XlDialogScenarioCells                    = 305
	XlDialogScenarioEdit                     = 308
	XlDialogScenarioMerge                    = 473
	XlDialogScenarioSummary                  = 311
	XlDialogScrollbarProperties              = 420
	XlDialogSearch                           = 731
	XlDialogSelectSpecial                    = 132
	XlDialogSendMail                         = 189
	XlDialogSeriesAxes                       = 460
	XlDialogSeriesOptions                    = 557
	XlDialogSeriesOrder                      = 466
	XlDialogSeriesShape                      = 504
	XlDialogSeriesX                          = 461
	XlDialogSeriesY                          = 462
	XlDialogSetBackgroundPicture             = 509
	XlDialogSetMDXEditor                     = 1208
	XlDialogSetManager                       = 1109
	XlDialogSetPrintTitles                   = 23
	XlDialogSetTupleEditorOnColumns          = 1108
	XlDialogSetTupleEditorOnRows             = 1107
	XlDialogSetUpdateStatus                  = 159
	XlDialogSheet                            = -4116
	XlDialogShowDetail                       = 204
	XlDialogShowToolbar                      = 220
	XlDialogSize                             = 261
	XlDialogSlicerCreation                   = 1182
	XlDialogSlicerPivotTableConnections      = 1184
	XlDialogSlicerSettings                   = 1179
	XlDialogSort                             = 39
	XlDialogSortSpecial                      = 192
	XlDialogSparklineInsertColumn            = 1134
	XlDialogSparklineInsertLine              = 1133
	XlDialogSparklineInsertWinLoss           = 1135
	XlDialogSplit                            = 137
	XlDialogStandardFont                     = 190
	XlDialogStandardWidth                    = 472
	XlDialogStyle                            = 44
	XlDialogSubscribeTo                      = 218
	XlDialogSubtotalCreate                   = 398
	XlDialogTabOrder                         = 394
	XlDialogTable                            = 41
	XlDialogTextToColumns                    = 422
	XlDialogUnhide                           = 94
	XlDialogUpdateLink                       = 201
	XlDialogVbaInsertFile                    = 328
	XlDialogVbaMakeAddin                     = 478
	XlDialogVbaProcedureDefinition           = 330
	XlDialogView3d                           = 197
	XlDialogWebOptionsBrowsers               = 773
	XlDialogWebOptionsEncoding               = 686
	XlDialogWebOptionsFiles                  = 684
	XlDialogWebOptionsFonts                  = 687
	XlDialogWebOptionsGeneral                = 683
	XlDialogWebOptionsPictures               = 685
	XlDialogWindowMove                       = 14
	XlDialogWindowSize                       = 13
	XlDialogWorkbookAdd                      = 281
	XlDialogWorkbookCopy                     = 283
	XlDialogWorkbookInsert                   = 354
	XlDialogWorkbookMove                     = 282
	XlDialogWorkbookName                     = 386
	XlDialogWorkbookNew                      = 302
	XlDialogWorkbookOptions                  = 284
	XlDialogWorkbookProtect                  = 417
	XlDialogWorkbookTabSplit                 = 415
	XlDialogWorkbookUnhide                   = 384
	XlDialogWorkgroup                        = 199
	XlDialogWorkspace                        = 95
	XlDialogZoom                             = 256
	XlDifferenceFrom                         = 2
	XlDisabled                               = 0
	XlDisplayNone                            = 1
	XlDisplayShapes                          = -4104
	XlDisplayUnitLabel                       = 30
	XlDistinctCount                          = 11
	XlDoNotSaveChanges                       = 2
	XlDone                                   = 0
	XlDot                                    = -4118
	XlDouble                                 = -4119
	XlDoughnut                               = -4120
	XlDoughnutExploded                       = 80
	XlDown                                   = -4121
	XlDownBars                               = 20
	XlDownThenOver                           = 1
	XlDownward                               = -4170
	XlDropDown                               = 2
	XlDropLines                              = 26
	XlEMDFormat                              = 10
	XlEPS                                    = 8
	XlEdgeBottom                             = 9
	XlEdgeLeft                               = 7
	XlEdgeRight                              = 10
	XlEdgeTop                                = 8
	XlEditBox                                = 3
	XlEditionDate                            = 2
	XlEmptyCellReferences                    = 7
	XlEnd                                    = 2
	XlEndSides                               = 3
	XlEntirePage                             = 1
	XlEqual                                  = 3
	XlErrDiv0                                = 2007
	XlErrNA                                  = 2042
	XlErrName                                = 2029
	XlErrNull                                = 2000
	XlErrNum                                 = 2036
	XlErrRef                                 = 2023
	XlErrValue                               = 2015
	XlErrorBarIncludeBoth                    = 1
	XlErrorBarIncludeMinusValues             = 3
	XlErrorBarIncludeNone                    = -4142
	XlErrorBarIncludePlusValues              = 2
	XlErrorBarTypeCustom                     = -4114
	XlErrorBarTypeFixedValue                 = 1
	XlErrorBarTypePercent                    = 2
	XlErrorBarTypeStDev                      = -4155
	XlErrorBarTypeStError                    = 4
	XlErrorBars                              = 9
	XlErrorHandler                           = 2
	XlErrors                                 = 16
	XlErrorsCondition                        = 16
	XlEscKey                                 = 1
	XlEvaluateToError                        = 1
	XlExcel12                                = 50
	XlExcel2                                 = 16
	XlExcel2FarEast                          = 27
	XlExcel3                                 = 29
	XlExcel4                                 = 33
	XlExcel4IntlMacroSheet                   = 4
	XlExcel4MacroSheet                       = 3
	XlExcel4Workbook                         = 35
	XlExcel5                                 = 39
	XlExcel8                                 = 56
	XlExcel9795                              = 43
	XlExcelLinks                             = 1
	XlExclusive                              = 3
	XlExponential                            = 5
	XlExpression                             = 2
	XlExternal                               = 2
	XlExtractData                            = 2
	XlFillCopy                               = 1
	XlFillDays                               = 5
	XlFillDefault                            = 0
	XlFillFormats                            = 3
	XlFillMonths                             = 7
	XlFillSeries                             = 2
	XlFillValues                             = 4
	XlFillWeekdays                           = 6
	XlFillWithAll                            = -4104
	XlFillWithContents                       = 2
	XlFillWithFormats                        = -4122
	XlFillYears                              = 8
	XlFilterAutomaticFontColor               = 13
	XlFilterCellColor                        = 8
	XlFilterCopy                             = 2
	XlFilterDynamic                          = 11
	XlFilterFontColor                        = 9
	XlFilterIcon                             = 10
	XlFilterInPlace                          = 1
	XlFilterNoFill                           = 12
	XlFilterNoIcon                           = 14
	XlFilterValues                           = 7
	XlFirstRow                               = 256
	XlFitToPage                              = 2
	XlFixedWidth                             = 2
	XlFlashFill                              = 11
	XlFloor                                  = 23
	XlFormatFromLeftOrAbove                  = 0
	XlFormatFromRightOrBelow                 = 1
	XlFormulas                               = -4123
	XlFreeFloating                           = 3
	XlFront                                  = 4
	XlFrontEnd                               = 6
	XlFrontSides                             = 5
	XlFullPage                               = 3
	XlFunction                               = 1
	XlGeneralFormat                          = 1
	XlGeneralFormatName                      = 26
	XlGreater                                = 5
	XlGreaterEqual                           = 7
	XlGroupBox                               = 4
	XlGrowth                                 = 2
	XlGrowthTrend                            = 10
	XlGuess                                  = 0
	XlHAlignCenter                           = -4108
	XlHAlignCenterAcrossSelection            = 7
	XlHAlignDistributed                      = -4117
	XlHAlignFill                             = 5
	XlHAlignGeneral                          = 1
	XlHAlignJustify                          = -4130
	XlHAlignLeft                             = -4131
	XlHAlignRight                            = -4152
	XlHGL                                    = 6
	XlHairline                               = 1
	XlHebrewFullScript                       = 0
	XlHebrewMixedAuthorizedScript            = 3
	XlHebrewMixedScript                      = 2
	XlHebrewPartialScript                    = 1
	XlHiLoLines                              = 25
	XlHidden                                 = 0
	XlHide                                   = 3
	XlHierarchy                              = 1
	XlHiragana                               = 2
	XlHorizontal                             = -4128
	XlHourCode                               = 22
	XlHtml                                   = 44
	XlHtmlCalc                               = 1
	XlHtmlChart                              = 3
	XlHtmlList                               = 2
	XlHtmlStatic                             = 0
	XlHundredMillions                        = -8
	XlHundredThousands                       = -5
	XlHundreds                               = -2
	XlIBeam                                  = 3
	XlIMEModeAlpha                           = 8
	XlIMEModeAlphaFull                       = 7
	XlIMEModeDisable                         = 3
	XlIMEModeHangul                          = 10
	XlIMEModeHangulFull                      = 9
	XlIMEModeHiragana                        = 4
	XlIMEModeKatakana                        = 5
	XlIMEModeKatakanaHalf                    = 6
	XlIMEModeNoControl                       = 0
	XlIMEModeOff                             = 2
	XlIMEModeOn                              = 1
	XlIconSets                               = 6
	XlInconsistentFormula                    = 4
	XlInconsistentListFormula                = 9
	XlIndex                                  = 9
	XlIndicatorAndButton                     = 0
	XlInfo                                   = -4129
	XlInsertDeleteCells                      = 1
	XlInsertEntireRows                       = 2
	XlInsideHorizontal                       = 12
	XlInsideVertical                         = 11
	XlInterpolated                           = 3
	XlInterrupt                              = 1
	XlIntlAddIn                              = 26
	XlIntlMacro                              = 25
	XlKatakana                               = 1
	XlKatakanaHalf                           = 0
	XlLabel                                  = 5
	XlLabelOnly                              = 1
	XlLabelPositionAbove                     = 0
	XlLabelPositionBelow                     = 1
	XlLabelPositionBestFit                   = 5
	XlLabelPositionCenter                    = -4108
	XlLabelPositionCustom                    = 7
	XlLabelPositionInsideBase                = 4
	XlLabelPositionInsideEnd                 = 3
	XlLabelPositionLeft                      = -4131
	XlLabelPositionMixed                     = 6
	XlLabelPositionOutsideEnd                = 2
	XlLabelPositionRight                     = -4152
	XlLandscape                              = 2
	XlLeaderLines                            = 29
	XlLeftBrace                              = 12
	XlLeftBracket                            = 10
	XlLegend                                 = 24
	XlLegendEntry                            = 12
	XlLegendKey                              = 13
	XlLegendPositionBottom                   = -4107
	XlLegendPositionCorner                   = 2
	XlLegendPositionCustom                   = -4161
	XlLegendPositionLeft                     = -4131
	XlLegendPositionRight                    = -4152
	XlLegendPositionTop                      = -4160
	XlLess                                   = 6
	XlLessEqual                              = 8
	XlLine                                   = 4
	XlLineMarkers                            = 65
	XlLineMarkersStacked                     = 66
	XlLineMarkersStacked100                  = 67
	XlLineStacked                            = 63
	XlLineStacked100                         = 64
	XlLineStyleNone                          = -4142
	XlLinear                                 = -4132
	XlLinearTrend                            = 9
	XlLinkInfoOLELinks                       = 2
	XlLinkInfoPublishers                     = 5
	XlLinkInfoStatus                         = 3
	XlLinkInfoSubscribers                    = 6
	XlLinkStatusCopiedValues                 = 10
	XlLinkStatusIndeterminate                = 5
	XlLinkStatusInvalidName                  = 7
	XlLinkStatusMissingFile                  = 1
	XlLinkStatusMissingSheet                 = 2
	XlLinkStatusNotStarted                   = 6
	XlLinkStatusOK                           = 0
	XlLinkStatusOld                          = 3
	XlLinkStatusSourceNotCalculated          = 4
	XlLinkStatusSourceNotOpen                = 8
	XlLinkStatusSourceOpen                   = 9
	XlLinkTypeExcelLinks                     = 1
	XlLinkTypeOLELinks                       = 2
	XlListBox                                = 6
	XlListConflictDialog                     = 0
	XlListConflictDiscardAllConflicts        = 2
	XlListConflictError                      = 3
	XlListConflictRetryAllConflicts          = 1
	XlListDataTypeCheckbox                   = 9
	XlListDataTypeChoice                     = 6
	XlListDataTypeChoiceMulti                = 7
	XlListDataTypeCounter                    = 11
	XlListDataTypeCurrency                   = 4
	XlListDataTypeDateTime                   = 5
	XlListDataTypeHyperLink                  = 10
	XlListDataTypeListLookup                 = 8
	XlListDataTypeMultiLineRichText          = 12
	XlListDataTypeMultiLineText              = 2
	XlListDataTypeNone                       = 0
	XlListDataTypeNumber                     = 3
	XlListDataTypeText                       = 1
	XlListDataValidation                     = 8
	XlListSeparator                          = 5
	XlLocalSessionChanges                    = 2
	XlLocationAsNewSheet                     = 1
	XlLocationAsObject                       = 2
	XlLocationAutomatic                      = 3
	XlLogarithmic                            = -4133
	XlLogical                                = 4
	XlLowerCaseColumnLetter                  = 9
	XlLowerCaseRowLetter                     = 8
	XlMAPI                                   = 1
	XlMDY                                    = 44
	XlMDYFormat                              = 3
	XlMSDOS                                  = 3
	XlMYDFormat                              = 6
	XlMacintosh                              = 1
	XlMajorGridlines                         = 15
	XlManualUpdate                           = 5
	XlMarkerStyleAutomatic                   = -4105
	XlMarkerStyleCircle                      = 8
	XlMarkerStyleDash                        = -4115
	XlMarkerStyleDiamond                     = 2
	XlMarkerStyleDot                         = -4118
	XlMarkerStyleNone                        = -4142
	XlMarkerStylePicture                     = -4147
	XlMarkerStylePlus                        = 9
	XlMarkerStyleSquare                      = 1
	XlMarkerStyleStar                        = 5
	XlMarkerStyleTriangle                    = 3
	XlMarkerStyleX                           = -4168
	XlMax                                    = -4136
	XlMaximized                              = -4137
	XlMeasure                                = 2
	XlMedium                                 = -4138
	XlMetric                                 = 35
	XlMicrosoftAccess                        = 4
	XlMicrosoftFoxPro                        = 5
	XlMicrosoftMail                          = 3
	XlMicrosoftPowerPoint                    = 2
	XlMicrosoftProject                       = 6
	XlMicrosoftSchedulePlus                  = 7
	XlMicrosoftWord                          = 1
	XlMillionMillions                        = -10
	XlMillions                               = -6
	XlMin                                    = -4139
	XlMinimized                              = -4140
	XlMinorGridlines                         = 16
	XlMinuteCode                             = 23
	XlMissingItemsDefault                    = -1
	XlMissingItemsMax                        = 32500
	XlMissingItemsMax2                       = 1048576
	XlMissingItemsNone                       = 0
	XlMixedLabels                            = 3
	XlMonth                                  = 3
	XlMonthCode                              = 20
	XlMonthLeadingZero                       = 41
	XlMonthNameChars                         = 30
	XlMonths                                 = 1
	XlMove                                   = 2
	XlMoveAndSize                            = 1
	XlMovingAvg                              = 6
	XlNever                                  = 2
	XlNext                                   = 1
	XlNo                                     = 2
	XlNoAdditionalCalculation                = -4143
	XlNoBlanksCondition                      = 13
	XlNoButton                               = 0
	XlNoButtonChanges                        = 1
	XlNoCap                                  = 2
	XlNoChange                               = 1
	XlNoChanges                              = 4
	XlNoConversion                           = 3
	XlNoDockingChanges                       = 3
	XlNoErrorsCondition                      = 17
	XlNoIndicator                            = 0
	XlNoKey                                  = 0
	XlNoLabels                               = -4142
	XlNoMailSystem                           = 0
	XlNoRestrictions                         = 0
	XlNoSelection                            = -4142
	XlNoShapeChanges                         = 2
	XlNonEnglishFunctions                    = 34
	XlNoncurrencyDigits                      = 29
	XlNormal                                 = -4143
	XlNormalLoad                             = 0
	XlNormalView                             = 1
	XlNorthwestArrow                         = 1
	XlNotBetween                             = 2
	XlNotEqual                               = 4
	XlNotPlotted                             = 1
	XlNotXLM                                 = 3
	XlNotYetReviewed                         = 3
	XlNotYetRouted                           = 0
	XlNothing                                = 28
	XlNumber                                 = -4145
	XlNumberAsText                           = 3
	XlNumbers                                = 1
	XlODBCQuery                              = 1
	XlOLEControl                             = 2
	XlOLEDBQuery                             = 5
	XlOLEEmbed                               = 1
	XlOLELink                                = 0
	XlOLELinks                               = 2
	XlOmittedCells                           = 5
	XlOneAfterAnother                        = 1
	XlOpenDocumentSpreadsheet                = 60
	XlOpenSource                             = 3
	XlOpenXMLAddIn                           = 55
	XlOpenXMLStrictWorkbook                  = 61
	XlOpenXMLTemplate                        = 54
	XlOpenXMLTemplateMacroEnabled            = 53
	XlOpenXMLWorkbook                        = 51
	XlOpenXMLWorkbookMacroEnabled            = 52
	XlOptionButton                           = 7
	XlOr                                     = 2
	XlOrigin                                 = 3
	XlOtherSessionChanges                    = 3
	XlOutline                                = 1
	XlOverThenDown                           = 2
	XlOverwriteCells                         = 0
	XlPCT                                    = 13
	XlPCX                                    = 10
	XlPIC                                    = 11
	XlPICT                                   = 1
	XlPLT                                    = 12
	XlPTClassic                              = 20
	XlPTNone                                 = 21
	XlPageBreakAutomatic                     = -4105
	XlPageBreakFull                          = 1
	XlPageBreakManual                        = -4135
	XlPageBreakNone                          = -4142
	XlPageBreakPartial                       = 2
	XlPageBreakPreview                       = 2
	XlPageField                              = 3
	XlPageHeader                             = 2
	XlPageItem                               = 6
	XlPageLayoutView                         = 3
	XlPaper10x14                             = 16
	XlPaper11x17                             = 17
	XlPaperA3                                = 8
	XlPaperA4                                = 9
	XlPaperA4Small                           = 10
	XlPaperA5                                = 11
	XlPaperB4                                = 12
	XlPaperB5                                = 13
	XlPaperCsheet                            = 24
	XlPaperDsheet                            = 25
	XlPaperEnvelope10                        = 20
	XlPaperEnvelope11                        = 21
	XlPaperEnvelope12                        = 22
	XlPaperEnvelope14                        = 23
	XlPaperEnvelope9                         = 19
	XlPaperEnvelopeB4                        = 33
	XlPaperEnvelopeB5                        = 34
	XlPaperEnvelopeB6                        = 35
	XlPaperEnvelopeC3                        = 29
	XlPaperEnvelopeC4                        = 30
	XlPaperEnvelopeC5                        = 28
	XlPaperEnvelopeC6                        = 31
	XlPaperEnvelopeC65                       = 32
	XlPaperEnvelopeDL                        = 27
	XlPaperEnvelopeItaly                     = 36
	XlPaperEnvelopeMonarch                   = 37
	XlPaperEnvelopePersonal                  = 38
	XlPaperEsheet                            = 26
	XlPaperExecutive                         = 7
	XlPaperFanfoldLegalGerman                = 41
	XlPaperFanfoldStdGerman                  = 40
	XlPaperFanfoldUS                         = 39
	XlPaperFolio                             = 14
	XlPaperLedger                            = 4
	XlPaperLegal                             = 5
	XlPaperLetter                            = 1
	XlPaperLetterSmall                       = 2
	XlPaperNote                              = 18
	XlPaperQuarto                            = 15
	XlPaperStatement                         = 6
	XlPaperTabloid                           = 3
	XlPaperUser                              = 256
	XlParamTypeBigInt                        = -5
	XlParamTypeBinary                        = -2
	XlParamTypeBit                           = -7
	XlParamTypeChar                          = 1
	XlParamTypeDate                          = 9
	XlParamTypeDecimal                       = 3
	XlParamTypeDouble                        = 8
	XlParamTypeFloat                         = 6
	XlParamTypeInteger                       = 4
	XlParamTypeLongVarBinary                 = -4
	XlParamTypeLongVarChar                   = -1
	XlParamTypeNumeric                       = 2
	XlParamTypeReal                          = 7
	XlParamTypeSmallInt                      = 5
	XlParamTypeTime                          = 10
	XlParamTypeTimestamp                     = 11
	XlParamTypeTinyInt                       = -6
	XlParamTypeUnknown                       = 0
	XlParamTypeVarBinary                     = -3
	XlParamTypeVarChar                       = 12
	XlParamTypeWChar                         = -8
	XlPart                                   = 2
	XlPasteAll                               = -4104
	XlPasteAllExceptBorders                  = 7
	XlPasteAllMergingConditionalFormats      = 14
	XlPasteAllUsingSourceTheme               = 13
	XlPasteColumnWidths                      = 8
	XlPasteComments                          = -4144
	XlPasteFormats                           = -4122
	XlPasteFormulas                          = -4123
	XlPasteFormulasAndNumberFormats          = 11
	XlPasteSpecialOperationAdd               = 2
	XlPasteSpecialOperationDivide            = 5
	XlPasteSpecialOperationMultiply          = 4
	XlPasteSpecialOperationNone              = -4142
	XlPasteSpecialOperationSubtract          = 3
	XlPasteValidation                        = 6
	XlPasteValues                            = -4163
	XlPasteValuesAndNumberFormats            = 12
	XlPatternAutomatic                       = -4105
	XlPatternChecker                         = 9
	XlPatternCrissCross                      = 16
	XlPatternDown                            = -4121
	XlPatternGray16                          = 17
	XlPatternGray25                          = -4124
	XlPatternGray50                          = -4125
	XlPatternGray75                          = -4126
	XlPatternGray8                           = 18
	XlPatternGrid                            = 15
	XlPatternHorizontal                      = -4128
	XlPatternLightDown                       = 13
	XlPatternLightHorizontal                 = 11
	XlPatternLightUp                         = 14
	XlPatternLightVertical                   = 12
	XlPatternLinearGradient                  = 4000
	XlPatternNone                            = -4142
	XlPatternRectangularGradient             = 4001
	XlPatternSemiGray75                      = 10
	XlPatternSolid                           = 1
	XlPatternUp                              = -4162
	XlPatternVertical                        = -4166
	XlPending                                = 2
	XlPercentDifferenceFrom                  = 4
	XlPercentOf                              = 3
	XlPercentOfColumn                        = 7
	XlPercentOfParent                        = 12
	XlPercentOfParentColumn                  = 11
	XlPercentOfParentRow                     = 10
	XlPercentOfRow                           = 6
	XlPercentOfTotal                         = 8
	XlPercentRunningTotal                    = 13
	XlPhoneticAlignCenter                    = 2
	XlPhoneticAlignDistributed               = 3
	XlPhoneticAlignLeft                      = 1
	XlPhoneticAlignNoControl                 = 0
	XlPicture                                = -4147
	XlPie                                    = 5
	XlPieExploded                            = 69
	XlPieOfPie                               = 68
	XlPinYin                                 = 1
	XlPivotCellBlankCell                     = 9
	XlPivotCellCustomSubtotal                = 7
	XlPivotCellDataField                     = 4
	XlPivotCellDataPivotField                = 8
	XlPivotCellGrandTotal                    = 3
	XlPivotCellPageFieldItem                 = 6
	XlPivotCellPivotField                    = 5
	XlPivotCellPivotItem                     = 1
	XlPivotCellSubtotal                      = 2
	XlPivotCellValue                         = 0
	XlPivotChartDropZone                     = 32
	XlPivotChartFieldButton                  = 31
	XlPivotTable                             = -4148
	XlPivotTableReport                       = 1
	XlPivotTableVersion10                    = 1
	XlPivotTableVersion11                    = 2
	XlPivotTableVersion12                    = 3
	XlPivotTableVersion14                    = 4
	XlPivotTableVersion15                    = 5
	XlPivotTableVersion2000                  = 0
	XlPivotTableVersionCurrent               = -1
	XlPlaceholders                           = 2
	XlPlotArea                               = 19
	XlPolynomial                             = 3
	XlPortrait                               = 1
	XlPower                                  = 4
	XlPowerTalk                              = 2
	XlPrevious                               = 2
	XlPrimary                                = 1
	XlPrimaryButton                          = 1
	XlPrintErrorsBlank                       = 1
	XlPrintErrorsDash                        = 2
	XlPrintErrorsDisplayed                   = 0
	XlPrintErrorsNA                          = 3
	XlPrintInPlace                           = 16
	XlPrintNoComments                        = -4142
	XlPrintSheetEnd                          = 1
	XlPrinter                                = 2
	XlPriorityHigh                           = -4127
	XlPriorityLow                            = -4134
	XlPriorityNormal                         = -4143
	XlProduct                                = -4149
	XlPrompt                                 = 0
	XlPublisher                              = 1
	XlPublishers                             = 5
	XlPyramidBarClustered                    = 109
	XlPyramidBarStacked                      = 110
	XlPyramidBarStacked100                   = 111
	XlPyramidCol                             = 112
	XlPyramidColClustered                    = 106
	XlPyramidColStacked                      = 107
	XlPyramidColStacked100                   = 108
	XlPyramidToMax                           = 2
	XlPyramidToPoint                         = 1
	XlQueryTable                             = 0
	XlR1C1                                   = -4150
	XlRTF                                    = 4
	XlRadar                                  = -4151
	XlRadarAxisLabels                        = 27
	XlRadarFilled                            = 82
	XlRadarMarkers                           = 81
	XlRange                                  = 2
	XlRangeAutoFormat3DEffects1              = 13
	XlRangeAutoFormat3DEffects2              = 14
	XlRangeAutoFormatAccounting1             = 4
	XlRangeAutoFormatAccounting2             = 5
	XlRangeAutoFormatAccounting3             = 6
	XlRangeAutoFormatAccounting4             = 17
	XlRangeAutoFormatClassic1                = 1
	XlRangeAutoFormatClassic2                = 2
	XlRangeAutoFormatClassic3                = 3
	XlRangeAutoFormatClassicPivotTable       = 31
	XlRangeAutoFormatColor1                  = 7
	XlRangeAutoFormatColor2                  = 8
	XlRangeAutoFormatColor3                  = 9
	XlRangeAutoFormatList1                   = 10
	XlRangeAutoFormatList2                   = 11
	XlRangeAutoFormatList3                   = 12
	XlRangeAutoFormatLocalFormat1            = 15
	XlRangeAutoFormatLocalFormat2            = 16
	XlRangeAutoFormatLocalFormat3            = 19
	XlRangeAutoFormatLocalFormat4            = 20
	XlRangeAutoFormatNone                    = -4142
	XlRangeAutoFormatPTNone                  = 42
	XlRangeAutoFormatReport1                 = 21
	XlRangeAutoFormatReport10                = 30
	XlRangeAutoFormatReport2                 = 22
	XlRangeAutoFormatReport3                 = 23
	XlRangeAutoFormatReport4                 = 24
	XlRangeAutoFormatReport5                 = 25
	XlRangeAutoFormatReport6                 = 26
	XlRangeAutoFormatReport7                 = 27
	XlRangeAutoFormatReport8                 = 28
	XlRangeAutoFormatReport9                 = 29
	XlRangeAutoFormatSimple                  = -4154
	XlRangeAutoFormatTable1                  = 32
	XlRangeAutoFormatTable10                 = 41
	XlRangeAutoFormatTable2                  = 33
	XlRangeAutoFormatTable3                  = 34
	XlRangeAutoFormatTable4                  = 35
	XlRangeAutoFormatTable5                  = 36
	XlRangeAutoFormatTable6                  = 37
	XlRangeAutoFormatTable7                  = 38
	XlRangeAutoFormatTable8                  = 39
	XlRangeAutoFormatTable9                  = 40
	XlRangeValueDefault                      = 10
	XlRangeValueMSPersistXML                 = 12
	XlRangeValueXMLSpreadsheet               = 11
	XlRankAscending                          = 14
	XlRankDecending                          = 15
	XlReadOnly                               = 3
	XlReadWrite                              = 2
	XlRelRowAbsColumn                        = 3
	XlRelative                               = 4
	XlRepairFile                             = 1
	XlReport1                                = 0
	XlReport10                               = 9
	XlReport2                                = 1
	XlReport3                                = 2
	XlReport4                                = 3
	XlReport5                                = 4
	XlReport6                                = 5
	XlReport7                                = 6
	XlReport8                                = 7
	XlReport9                                = 8
	XlRightBrace                             = 13
	XlRightBracket                           = 11
	XlRoutingComplete                        = 2
	XlRoutingInProgress                      = 1
	XlRowField                               = 1
	XlRowHeader                              = -4153
	XlRowItem                                = 4
	XlRowLabels                              = 1
	XlRowSeparator                           = 15
	XlRowThenColumn                          = 1
	XlRows                                   = 1
	XlRunningTotal                           = 5
	XlSYLK                                   = 2
	XlSaveChanges                            = 1
	XlScaleLinear                            = -4132
	XlScaleLogarithmic                       = -4133
	XlScenario                               = 4
	XlScreen                                 = 1
	XlScreenSize                             = 1
	XlScrollBar                              = 8
	XlSecondCode                             = 24
	XlSecondary                              = 2
	XlSecondaryButton                        = 2
	XlSeries                                 = 3
	XlSeriesAxis                             = 3
	XlSeriesLines                            = 22
	XlSet                                    = 3
	XlShape                                  = 14
	XlShared                                 = 2
	XlSheetHidden                            = 0
	XlSheetVeryHidden                        = 2
	XlSheetVisible                           = -1
	XlShiftDown                              = -4121
	XlShiftToLeft                            = -4159
	XlShiftToRight                           = -4161
	XlShiftUp                                = -4162
	XlSides                                  = 1
	XlSinceMyLastSave                        = 1
	XlSizeIsArea                             = 1
	XlSizeIsWidth                            = 2
	XlSkipColumn                             = 9
	XlSlantDashDot                           = 13
	XlSmartTagControlActiveX                 = 13
	XlSmartTagControlButton                  = 6
	XlSmartTagControlCheckbox                = 9
	XlSmartTagControlCombo                   = 12
	XlSmartTagControlHelp                    = 3
	XlSmartTagControlHelpURL                 = 4
	XlSmartTagControlImage                   = 8
	XlSmartTagControlLabel                   = 7
	XlSmartTagControlLink                    = 2
	XlSmartTagControlListbox                 = 11
	XlSmartTagControlRadioGroup              = 14
	XlSmartTagControlSeparator               = 5
	XlSmartTagControlSmartTag                = 1
	XlSmartTagControlTextbox                 = 10
	XlSortColumns                            = 1
	XlSortLabels                             = 2
	XlSortNormal                             = 0
	XlSortRows                               = 2
	XlSortTextAsNumbers                      = 1
	XlSortValues                             = 1
	XlSourceAutoFilter                       = 3
	XlSourceChart                            = 5
	XlSourcePivotTable                       = 6
	XlSourcePrintArea                        = 2
	XlSourceQuery                            = 7
	XlSourceRange                            = 4
	XlSourceSheet                            = 1
	XlSourceWorkbook                         = 0
	XlSpeakByColumns                         = 1
	XlSpeakByRows                            = 0
	XlSpecifiedTables                        = 3
	XlSpinner                                = 9
	XlSplitByCustomSplit                     = 4
	XlSplitByPercentValue                    = 3
	XlSplitByPosition                        = 1
	XlSplitByValue                           = 2
	XlSrcExternal                            = 0
	XlSrcModel                               = 4
	XlSrcQuery                               = 3
	XlSrcRange                               = 1
	XlSrcXml                                 = 2
	XlStDev                                  = -4155
	XlStDevP                                 = -4156
	XlStack                                  = 2
	XlStackScale                             = 3
	XlStandardSummary                        = 1
	XlStockHLC                               = 88
	XlStockOHLC                              = 89
	XlStockVHLC                              = 90
	XlStockVOHLC                             = 91
	XlStretch                                = 1
	XlStroke                                 = 2
	XlSubscribeToPicture                     = -4147
	XlSubscribeToText                        = -4158
	XlSubscriber                             = 2
	XlSubscribers                            = 6
	XlSum                                    = -4157
	XlSummaryAbove                           = 0
	XlSummaryBelow                           = 1
	XlSummaryOnLeft                          = -4131
	XlSummaryOnRight                         = -4152
	XlSummaryPivotTable                      = -4148
	XlSurface                                = 83
	XlSurfaceTopView                         = 85
	XlSurfaceTopViewWireframe                = 86
	XlSurfaceWireframe                       = 84
	XlSyllabary                              = 1
	XlTIF                                    = 9
	XlTabPositionFirst                       = 0
	XlTabPositionLast                        = 1
	XlTable                                  = 2
	XlTable1                                 = 10
	XlTable10                                = 19
	XlTable2                                 = 11
	XlTable3                                 = 12
	XlTable4                                 = 13
	XlTable5                                 = 14
	XlTable6                                 = 15
	XlTable7                                 = 16
	XlTable8                                 = 17
	XlTable9                                 = 18
	XlTableBody                              = 8
	XlTabular                                = 0
	XlTemplate8                              = 17
	XlTenMillions                            = -7
	XlTenThousands                           = -4
	XlText                                   = -4158
	XlTextDate                               = 2
	XlTextFormat                             = 2
	XlTextImport                             = 6
	XlTextMSDOS                              = 21
	XlTextMac                                = 19
	XlTextPrinter                            = 36
	XlTextQualifierDoubleQuote               = 1
	XlTextQualifierNone                      = -4142
	XlTextQualifierSingleQuote               = 2
	XlTextString                             = 9
	XlTextValues                             = 2
	XlTextVisualLTR                          = 1
	XlTextVisualRTL                          = 2
	XlTextWindows                            = 20
	XlThick                                  = 4
	XlThin                                   = 2
	XlThousandMillions                       = -9
	XlThousands                              = -3
	XlThousandsSeparator                     = 4
	XlTickLabelOrientationAutomatic          = -4105
	XlTickLabelOrientationDownward           = -4170
	XlTickLabelOrientationHorizontal         = -4128
	XlTickLabelOrientationUpward             = -4171
	XlTickLabelOrientationVertical           = -4166
	XlTickLabelPositionHigh                  = -4127
	XlTickLabelPositionLow                   = -4134
	XlTickLabelPositionNextToAxis            = 4
	XlTickLabelPositionNone                  = -4142
	XlTickMarkCross                          = 4
	XlTickMarkInside                         = 2
	XlTickMarkNone                           = -4142
	XlTickMarkOutside                        = 3
	XlTimeLeadingZero                        = 45
	XlTimePeriod                             = 11
	XlTimeScale                              = 3
	XlTimeSeparator                          = 18
	XlToLeft                                 = -4159
	XlToRight                                = -4161
	XlToolbarProtectionNone                  = -4143
	XlTop10                                  = 5
	XlTop10Items                             = 3
	XlTop10Percent                           = 5
	XlTotalsCalculationAverage               = 2
	XlTotalsCalculationCount                 = 3
	XlTotalsCalculationCountNums             = 4
	XlTotalsCalculationCustom                = 9
	XlTotalsCalculationMax                   = 6
	XlTotalsCalculationMin                   = 5
	XlTotalsCalculationNone                  = 0
	XlTotalsCalculationStdDev                = 7
	XlTotalsCalculationSum                   = 1
	XlTotalsCalculationVar                   = 8
	XlTrendline                              = 8
	XlUnderlineStyleDouble                   = -4119
	XlUnderlineStyleDoubleAccounting         = 5
	XlUnderlineStyleNone                     = -4142
	XlUnderlineStyleSingle                   = 2
	XlUnderlineStyleSingleAccounting         = 4
	XlUnicodeText                            = 42
	XlUniqueValues                           = 8
	XlUnknown                                = 1000
	XlUnlockedCells                          = 1
	XlUnlockedFormulaCells                   = 6
	XlUp                                     = -4162
	XlUpBars                                 = 18
	XlUpdateLinksAlways                      = 3
	XlUpdateLinksNever                       = 2
	XlUpdateLinksUserSetting                 = 1
	XlUpdateState                            = 1
	XlUpdateSubscriber                       = 2
	XlUpperCaseColumnLetter                  = 7
	XlUpperCaseRowLetter                     = 6
	XlUpward                                 = -4171
	XlUserDefined                            = 22
	XlUserResolution                         = 1
	XlVALU                                   = 8
	XlVAlignBottom                           = -4107
	XlVAlignCenter                           = -4108
	XlVAlignDistributed                      = -4117
	XlVAlignJustify                          = -4130
	XlVAlignTop                              = -4160
	XlValidAlertInformation                  = 3
	XlValidAlertStop                         = 1
	XlValidAlertWarning                      = 2
	XlValidateCustom                         = 7
	XlValidateDate                           = 4
	XlValidateDecimal                        = 2
	XlValidateInputOnly                      = 0
	XlValidateList                           = 3
	XlValidateTextLength                     = 6
	XlValidateTime                           = 5
	XlValidateWholeNumber                    = 1
	XlValue                                  = 2
	XlValues                                 = -4163
	XlVar                                    = -4164
	XlVarP                                   = -4165
	XlVerbOpen                               = 2
	XlVerbPrimary                            = 1
	XlVertical                               = -4166
	XlWBATChart                              = -4109
	XlWBATExcel4IntlMacroSheet               = 4
	XlWBATExcel4MacroSheet                   = 3
	XlWBATWorksheet                          = -4167
	XlWJ2WD1                                 = 14
	XlWJ3                                    = 40
	XlWJ3FJ3                                 = 41
	XlWK1                                    = 5
	XlWK1ALL                                 = 31
	XlWK1FMT                                 = 30
	XlWK3                                    = 15
	XlWK3FM3                                 = 32
	XlWK4                                    = 38
	XlWKS                                    = 4
	XlWMF                                    = 2
	XlWPG                                    = 3
	XlWQ1                                    = 34
	XlWait                                   = 2
	XlWalls                                  = 5
	XlWebArchive                             = 45
	XlWebFormattingAll                       = 1
	XlWebFormattingNone                      = 3
	XlWebFormattingRTF                       = 2
	XlWebQuery                               = 4
	XlWeekday                                = 2
	XlWeekdayNameChars                       = 31
	XlWhole                                  = 1
	XlWindows                                = 2
	XlWithinSheet                            = 1
	XlWithinWorkbook                         = 2
	XlWorkbook                               = 1
	XlWorkbookDefault                        = 51
	XlWorkbookNormal                         = -4143
	XlWorks2FarEast                          = 28
	XlWorksheet                              = -4167
	XlX                                      = -4168
	XlXErrorBars                             = 10
	XlXMLSpreadsheet                         = 46
	XlXYScatter                              = -4169
	XlXYScatterLines                         = 74
	XlXYScatterLinesNoMarkers                = 75
	XlXYScatterSmooth                        = 72
	XlXYScatterSmoothNoMarkers               = 73
	XlXmlExportSuccess                       = 0
	XlXmlExportValidationFailed              = 1
	XlXmlImportElementsTruncated             = 1
	XlXmlImportSuccess                       = 0
	XlXmlImportValidationFailed              = 2
	XlXmlLoadImportToList                    = 2
	XlXmlLoadMapXml                          = 3
	XlXmlLoadOpenXml                         = 1
	XlXmlLoadPromptUser                      = 0
	XlY                                      = 1
	XlYDMFormat                              = 8
	XlYErrorBars                             = 11
	XlYMDFormat                              = 5
	XlYear                                   = 4
	XlYearCode                               = 19
	XlYears                                  = 2
	XlYes                                    = 1
	XlZero                                   = 2
	_xlDialogChartSourceData                 = 541
	_xlDialogPhonetic                        = 538
)

type IDispatcher interface {
	IDispatch() *ole.IDispatch
}

type Merger interface {
	Merge(*ole.VARIANT, error) Excel
}

type Operation interface {
	IDispatcher
	Merger
}

func ToString(v *ole.VARIANT, err error) (ret string) {
	if v.Value() != nil && err == nil {
		ret = v.ToString()
	}
	return
}
func ToBool(v *ole.VARIANT, err error) (ret bool) {
	if err == nil {
		if i := v.Value(); i != nil {
			if b, ok := i.(bool); ok {
				ret = b
			}
		}
	}
	return
}
func ToTime(v *ole.VARIANT, err error) (t time.Time) {
	if err == nil {
		f := *(*float64)(unsafe.Pointer(&v.Val))
		t = time.Date(1900, time.January, 1, 0, 0, 0, 0, time.Local)
		t = t.Add(time.Hour * 24 * time.Duration(int64(f)-2))
		t = t.Add(time.Millisecond * time.Duration((f-float64(int64(f)))/(1.0/86400000.0)))
	}
	return
}

type Error []error

func (e *Error) Error() (ret string) {
	for _, it := range *e {
		if it != nil {
			ret += it.Error()
		}
	}
	return ret
}

func MultiError(e error, es ...error) error {
	var ee Error
	if len(es) <= 0 {
		ee = Error([]error{e})
	} else {
		ee = Error(append([]error{e}, es...))
	}
	return &ee
}

type Excel struct {
	Obj      *ole.IDispatch
	Err      error
	children []Excel
}

func (e *Excel) Merge(obj *ole.VARIANT, err error) Excel {
	ce := Excel{
		Obj: obj.ToIDispatch(),
		Err: err,
	}
	e.children = append(e.children, ce)
	if e.Err == nil {
		if err != nil {
			e.Err = err
		}
	} else {
		if err != nil {
			e.Err = MultiError(e.Err, err)
		}
	}
	return ce
}
func (e *Excel) IDispatch() *ole.IDispatch {
	return e.Obj
}
func (e *Excel) Error() (ret string) {
	if e.Err != nil {
		ret = e.Err.Error()
	}
	return
}
func (e *Excel) Release() {
	if e.children != nil {
		for i, _ := range e.children {
			e.children[i].Release()
		}
		e.children = nil
	}
	if e.Obj != nil {
		e.Obj.Release()
		e.Obj = nil
	}
}

type Application struct {
	Excel
}
type Axis struct {
	Excel
}
type AxisTitle struct {
	Excel
}
type Chart struct {
	Excel
}
type ChartTitle struct {
	Excel
}
type Comment struct {
	Excel
}
type Legend struct {
	Excel
}
type Name struct {
	Excel
}
type Outline struct {
	Excel
}
type Range struct {
	Excel
}
type Series struct {
	Excel
}
type Workbook struct {
	Excel
}
type Worksheet struct {
	Excel
}

func ThisApplication() *Application {
	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return nil
	}
	obj, err := unknown.QueryInterface(ole.IID_IDispatch)
	return &Application{
		Excel: Excel{Obj: obj, Err: err},
	}
}
func (a *Excel) GetColumns(a0 ...interface{}) *Range {
	return &Range{
		Excel: a.Merge(a.Obj.GetProperty("Columns", a0...)),
	}
}
func (a *Excel) GetEnd(a0 int, a1 int) *Range {
	return &Range{
		Excel: a.Merge(a.Obj.GetProperty("End", a0, a1)),
	}
}
func (a *Excel) GetUnion(a0 ...*Range) *Range {
	av := make([]interface{}, len(a0))
	for i, it := range a0 {
		av[i] = it.Obj
	}
	return &Range{
		Excel: a.Merge(a.Obj.GetProperty("Union", av...)),
	}
}
func (a *Excel) GetApplication() *Application {
	return &Application{
		Excel: a.Merge(a.Obj.GetProperty("Application")),
	}
}
func (a *Excel) GetCount() int {
	v, err := a.Obj.GetProperty("Count")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *Excel) GetCells(a0 int, a1 int) *Range {
	return &Range{
		Excel: a.Merge(a.Obj.GetProperty("Cells", a0, a1)),
	}
}
func (a *Excel) GetValue() *ole.VARIANT {
	v, err := a.Obj.GetProperty("Value")
	a.Merge(v, err)
	return v
}
func (a *Excel) SetValue(a0 ...interface{}) {
	v, err := a.Obj.PutProperty("Value", a0...)
	a.Merge(v, err)
}
func (a *Excel) GetName() string {
	v, err := a.Obj.GetProperty("Name")
	a.Merge(v, err)
	return ToString(v, err)
}
func (a *Excel) SetName(a0 string) {
	v, err := a.Obj.PutProperty("Name", a0)
	a.Merge(v, err)
}
func (a *Excel) GetRange(a0 ...*Range) *Range {
	av := make([]interface{}, len(a0))
	for i, it := range a0 {
		av[i] = it.Obj
	}
	return &Range{
		Excel: a.Merge(a.Obj.GetProperty("Range", av...)),
	}
}
func (a *Excel) GetVisible() bool {
	v, err := a.Obj.GetProperty("Visible")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Excel) SetVisible(a0 bool) {
	v, err := a.Obj.PutProperty("Visible", a0)
	a.Merge(v, err)
}
func (a *Excel) GetCreator() int {
	v, err := a.Obj.GetProperty("Creator")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *Excel) CheckSpelling(a0 ...interface{}) bool {
	v, err := a.Obj.CallMethod("CheckSpelling", a0...)
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Excel) Evaluate(a0 string) *Range {
	return &Range{
		Excel: a.Merge(a.Obj.CallMethod("Evaluate", a0)),
	}
}
func (a *Worksheet) GetOutline() *Outline {
	return &Outline{
		Excel: a.Merge(a.Obj.GetProperty("Outline")),
	}
}
func (a *Worksheet) GetComments() *Comment {
	return &Comment{
		Excel: a.Merge(a.Obj.GetProperty("Comments")),
	}
}
func (a *Worksheet) ChartObjects() *Chart {
	return &Chart{
		Excel: a.Merge(a.Obj.CallMethod("ChartObjects")),
	}
}
func (a *Worksheet) Calculate() {
	v, err := a.Obj.CallMethod("Calculate")
	a.Merge(v, err)
}
func (a *Chart) GetHasTitle() bool {
	v, err := a.Obj.GetProperty("HasTitle")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Chart) SetHasTitle(a0 bool) {
	v, err := a.Obj.PutProperty("HasTitle", a0)
	a.Merge(v, err)
}
func (a *Chart) GetChartTitle() *ChartTitle {
	return &ChartTitle{
		Excel: a.Merge(a.Obj.GetProperty("ChartTitle")),
	}
}
func (a *Chart) GetChart() *Chart {
	return &Chart{
		Excel: a.Merge(a.Obj.GetProperty("Chart")),
	}
}
func (a *Chart) SetChartType(a0 int) {
	v, err := a.Obj.PutProperty("ChartType", a0)
	a.Merge(v, err)
}
func (a *Chart) GetLegend() *Legend {
	return &Legend{
		Excel: a.Merge(a.Obj.GetProperty("Legend")),
	}
}
func (a *Chart) GetSeriesCollection(a0 ...interface{}) bool {
	v, err := a.Obj.GetProperty("SeriesCollection", a0...)
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Chart) Add(a0 int, a1 int, a2 int, a3 int) *Chart {
	return &Chart{
		Excel: a.Merge(a.Obj.CallMethod("Add", a0, a1, a2, a3)),
	}
}
func (a *Chart) SetSourceData(a0 *Range, a1 int) *Chart {
	return &Chart{
		Excel: a.Merge(a.Obj.CallMethod("SetSourceData", a0, a1)),
	}
}
func (a *Chart) Axes(a0 int, a1 int) *Axis {
	return &Axis{
		Excel: a.Merge(a.Obj.CallMethod("Axes", a0, a1)),
	}
}
func (a *Chart) Location(a0 ...interface{}) *Chart {
	return &Chart{
		Excel: a.Merge(a.Obj.CallMethod("Location", a0...)),
	}
}
func (a *Legend) GetPosition() int {
	v, err := a.Obj.GetProperty("Position")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *Legend) SetPosition(a0 int) {
	v, err := a.Obj.PutProperty("Position", a0)
	a.Merge(v, err)
}
func (a *AxisTitle) GetText() string {
	v, err := a.Obj.GetProperty("Text")
	a.Merge(v, err)
	return ToString(v, err)
}
func (a *AxisTitle) SetText(a0 string) {
	v, err := a.Obj.PutProperty("Text", a0)
	a.Merge(v, err)
}
func (a *Name) Add(a0 ...interface{}) *Name {
	return &Name{
		Excel: a.Merge(a.Obj.CallMethod("Add", a0...)),
	}
}
func (a *Name) Item(a0 ...interface{}) *Name {
	return &Name{
		Excel: a.Merge(a.Obj.CallMethod("Item", a0...)),
	}
}
func (a *Name) Delete() {
	v, err := a.Obj.CallMethod("Delete")
	a.Merge(v, err)
}
func (a *Outline) GetAutomaticStyles() bool {
	v, err := a.Obj.GetProperty("AutomaticStyles")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Outline) SetAutomaticStyles(a0 bool) {
	v, err := a.Obj.PutProperty("AutomaticStyles", a0)
	a.Merge(v, err)
}
func (a *Outline) GetSummaryColumn() int {
	v, err := a.Obj.GetProperty("SummaryColumn")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *Outline) SetSummaryColumn(a0 int) {
	v, err := a.Obj.PutProperty("SummaryColumn", a0)
	a.Merge(v, err)
}
func (a *Outline) GetSummaryRow() int {
	v, err := a.Obj.GetProperty("SummaryRow")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *Outline) SetSummaryRow(a0 int) {
	v, err := a.Obj.PutProperty("SummaryRow", a0)
	a.Merge(v, err)
}
func (a *Outline) ShowLevels(a0 ...interface{}) {
	v, err := a.Obj.CallMethod("ShowLevels", a0...)
	a.Merge(v, err)
}
func (a *Workbook) GetHasPassword() bool {
	v, err := a.Obj.GetProperty("HasPassword")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Workbook) GetPrecisionAsDisplayed() bool {
	v, err := a.Obj.GetProperty("PrecisionAsDisplayed")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Workbook) SetPrecisionAsDisplayed(a0 bool) {
	v, err := a.Obj.PutProperty("PrecisionAsDisplayed", a0)
	a.Merge(v, err)
}
func (a *Workbook) GetWorkbooks(a0 ...interface{}) *Workbook {
	return &Workbook{
		Excel: a.Merge(a.Obj.GetProperty("Workbooks", a0...)),
	}
}
func (a *Workbook) GetWorksheets(a0 ...interface{}) *Worksheet {
	return &Worksheet{
		Excel: a.Merge(a.Obj.GetProperty("Worksheets", a0...)),
	}
}
func (a *Workbook) GetNames(a0 ...interface{}) *Name {
	return &Name{
		Excel: a.Merge(a.Obj.GetProperty("Names", a0...)),
	}
}
func (a *Workbook) GetPath() string {
	v, err := a.Obj.GetProperty("Path")
	a.Merge(v, err)
	return ToString(v, err)
}
func (a *Workbook) GetFullName() string {
	v, err := a.Obj.GetProperty("FullName")
	a.Merge(v, err)
	return ToString(v, err)
}
func (a *Workbook) SetPassword(a0 string) {
	v, err := a.Obj.PutProperty("Password", a0)
	a.Merge(v, err)
}
func (a *Workbook) Protect(a0 ...interface{}) {
	v, err := a.Obj.CallMethod("Protect", a0...)
	a.Merge(v, err)
}
func (a *Workbook) ProtectSharing(a0 ...interface{}) {
	v, err := a.Obj.CallMethod("ProtectSharing", a0...)
	a.Merge(v, err)
}
func (a *Workbook) Unprotect(a0 string) {
	v, err := a.Obj.CallMethod("Unprotect", a0)
	a.Merge(v, err)
}
func (a *Workbook) UnprotectSharing(a0 string) {
	v, err := a.Obj.CallMethod("UnprotectSharing", a0)
	a.Merge(v, err)
}
func (a *Workbook) Close() {
	v, err := a.Obj.CallMethod("Close")
	a.Merge(v, err)
}
func (a *Workbook) Open(a0 string) *Workbook {
	return &Workbook{
		Excel: a.Merge(a.Obj.CallMethod("Open", a0)),
	}
}
func (a *Workbook) SaveAs(a0 string, a1 int) *Workbook {
	return &Workbook{
		Excel: a.Merge(a.Obj.CallMethod("SaveAs", a0, a1)),
	}
}
func (a *Workbook) Activate() {
	v, err := a.Obj.CallMethod("Activate")
	a.Merge(v, err)
}
func (a *Range) GetOutline() *Outline {
	return &Outline{
		Excel: a.Merge(a.Obj.GetProperty("Outline")),
	}
}
func (a *Range) GetColumn() int {
	v, err := a.Obj.GetProperty("Column")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *Range) Find(a0 ...interface{}) *Range {
	return &Range{
		Excel: a.Merge(a.Obj.CallMethod("Find", a0...)),
	}
}
func (a *Range) FindNext(a0 ...interface{}) *Range {
	return &Range{
		Excel: a.Merge(a.Obj.CallMethod("FindNext", a0...)),
	}
}
func (a *Range) FindPrevious(a0 ...interface{}) *Range {
	return &Range{
		Excel: a.Merge(a.Obj.CallMethod("FindPrevious", a0...)),
	}
}
func (a *Range) Sort(a0 ...interface{}) {
	v, err := a.Obj.CallMethod("Sort", a0...)
	a.Merge(v, err)
}
func (a *Range) Calculate() {
	v, err := a.Obj.CallMethod("Calculate")
	a.Merge(v, err)
}
func (a *Range) AddComment(a0 ...interface{}) *Comment {
	return &Comment{
		Excel: a.Merge(a.Obj.CallMethod("AddComment", a0...)),
	}
}
func (a *Range) AutoFill(a0 *Range, a1 int) {
	v, err := a.Obj.CallMethod("AutoFill", a0, a1)
	a.Merge(v, err)
}
func (a *Range) AutoComplete(a0 string) string {
	v, err := a.Obj.CallMethod("AutoComplete", a0)
	a.Merge(v, err)
	return ToString(v, err)
}
func (a *ChartTitle) GetText() string {
	v, err := a.Obj.GetProperty("Text")
	a.Merge(v, err)
	return ToString(v, err)
}
func (a *ChartTitle) SetText(a0 string) {
	v, err := a.Obj.PutProperty("Text", a0)
	a.Merge(v, err)
}
func (a *ChartTitle) GetPosition() int {
	v, err := a.Obj.GetProperty("Position")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *ChartTitle) SetPosition(a0 int) {
	v, err := a.Obj.PutProperty("Position", a0)
	a.Merge(v, err)
}
func (a *ChartTitle) GetIncludeInLayout() bool {
	v, err := a.Obj.GetProperty("IncludeInLayout")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *ChartTitle) SetIncludeInLayout(a0 bool) {
	v, err := a.Obj.PutProperty("IncludeInLayout", a0)
	a.Merge(v, err)
}
func (a *Series) GetAxisGroup() int {
	v, err := a.Obj.GetProperty("AxisGroup")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *Series) SetAxisGroup(a0 int) {
	v, err := a.Obj.PutProperty("AxisGroup", a0)
	a.Merge(v, err)
}
func (a *Series) GetXValues() *Range {
	return &Range{
		Excel: a.Merge(a.Obj.GetProperty("XValues")),
	}
}
func (a *Series) SetXValues(a0 *Range) {
	v, err := a.Obj.PutProperty("XValues", a0)
	a.Merge(v, err)
}
func (a *Axis) GetHasMajorGridlines() bool {
	v, err := a.Obj.GetProperty("HasMajorGridlines")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Axis) SetHasMajorGridlines(a0 bool) {
	v, err := a.Obj.PutProperty("HasMajorGridlines", a0)
	a.Merge(v, err)
}
func (a *Axis) GetHasMinorGridlines() bool {
	v, err := a.Obj.GetProperty("HasMinorGridlines")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Axis) SetHasMinorGridlines(a0 bool) {
	v, err := a.Obj.PutProperty("HasMinorGridlines", a0)
	a.Merge(v, err)
}
func (a *Axis) GetTickLabelPosition() int {
	v, err := a.Obj.GetProperty("TickLabelPosition")
	a.Merge(v, err)
	return (int)(v.Val)
}
func (a *Axis) SetTickLabelPosition(a0 int) {
	v, err := a.Obj.PutProperty("TickLabelPosition", a0)
	a.Merge(v, err)
}
func (a *Axis) GetHasTitle() bool {
	v, err := a.Obj.GetProperty("HasTitle")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Axis) SetHasTitle(a0 bool) {
	v, err := a.Obj.PutProperty("HasTitle", a0)
	a.Merge(v, err)
}
func (a *Axis) GetAxisTitle() *AxisTitle {
	return &AxisTitle{
		Excel: a.Merge(a.Obj.GetProperty("AxisTitle")),
	}
}
func (a *Axis) SetAxisTitle(a0 *AxisTitle) {
	v, err := a.Obj.PutProperty("AxisTitle", a0)
	a.Merge(v, err)
}
func (a *Comment) GetAuthor() string {
	v, err := a.Obj.GetProperty("Author")
	a.Merge(v, err)
	return ToString(v, err)
}
func (a *Comment) SetAuthor(a0 string) {
	v, err := a.Obj.PutProperty("Author", a0)
	a.Merge(v, err)
}
func (a *Comment) Item(a0 int) *Comment {
	return &Comment{
		Excel: a.Merge(a.Obj.CallMethod("Item", a0)),
	}
}
func (a *Comment) Delete() {
	v, err := a.Obj.CallMethod("Delete")
	a.Merge(v, err)
}
func (a *Comment) Next() *Comment {
	return &Comment{
		Excel: a.Merge(a.Obj.CallMethod("Next")),
	}
}
func (a *Comment) Previous() *Comment {
	return &Comment{
		Excel: a.Merge(a.Obj.CallMethod("Previous")),
	}
}
func (a *Comment) Text(a0 ...interface{}) string {
	v, err := a.Obj.CallMethod("Text", a0...)
	a.Merge(v, err)
	return ToString(v, err)
}
func (a *Application) GetWorkbooks(a0 ...interface{}) *Workbook {
	return &Workbook{
		Excel: a.Merge(a.Obj.GetProperty("Workbooks", a0...)),
	}
}
func (a *Application) GetDisplayAlerts() bool {
	v, err := a.Obj.GetProperty("DisplayAlerts")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Application) SetDisplayAlerts(a0 bool) {
	v, err := a.Obj.PutProperty("DisplayAlerts", a0)
	a.Merge(v, err)
}
func (a *Application) GetScreenUpdating() bool {
	v, err := a.Obj.GetProperty("ScreenUpdating")
	a.Merge(v, err)
	return ToBool(v, err)
}
func (a *Application) SetScreenUpdating(a0 bool) {
	v, err := a.Obj.PutProperty("ScreenUpdating", a0)
	a.Merge(v, err)
}
func (a *Application) Calculate() {
	v, err := a.Obj.CallMethod("Calculate")
	a.Merge(v, err)
}
func (a *Application) Quit() {
	v, err := a.Obj.CallMethod("Quit")
	a.Merge(v, err)
}
