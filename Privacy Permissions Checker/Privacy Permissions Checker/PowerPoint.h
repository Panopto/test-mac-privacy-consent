/*
 * PowerPoint.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class PowerPointBaseObject, PowerPointBaseApplication, PowerPointBaseDocument, PowerPointBasicWindow, PowerPointPrintSettings, PowerPointCommandBarControl, PowerPointCommandBarButton, PowerPointCommandBarCombobox, PowerPointCommandBarPopup, PowerPointCommandBar, PowerPointDocumentProperty, PowerPointCustomDocumentProperty, PowerPointWebPageFont, PowerPointActionSetting, PowerPointAnimationBehavior, PowerPointAnimationPoint, PowerPointAnimationSettings, PowerPointApplication, PowerPointAutocorrectEntry, PowerPointAutocorrect, PowerPointBulletFormat, PowerPointColorScheme, PowerPointColorsEffect, PowerPointCommandEffect, PowerPointCustomLayout, PowerPointDefaultWebOptions, PowerPointDesign, PowerPointDocumentWindow, PowerPointEffectInformation, PowerPointEffectParameters, PowerPointEffect, PowerPointFilterEffect, PowerPointFirstLetterException, PowerPointFont, PowerPointHeaderOrFooter, PowerPointHeadersAndFooters, PowerPointHyperlink, PowerPointMaster, PowerPointMotionEffect, PowerPointNamedSlideShow, PowerPointPageSetup, PowerPointPane, PowerPointParagraphFormat, PowerPointPlaySettings, PowerPointPreferences, PowerPointPresentation, PowerPointPresenterTool, PowerPointPresenterViewWindow, PowerPointPrintOptions, PowerPointPrintRange, PowerPointPropertyEffect, PowerPointRotatingEffect, PowerPointRulerLevel, PowerPointRuler, PowerPointSaveAsMovieSettings, PowerPointScaleEffect, PowerPointSelection, PowerPointSequence, PowerPointSetEffect, PowerPointSlideRange, PowerPointSlideShowSettings, PowerPointSlideShowTransition, PowerPointSlideShowView, PowerPointSlideShowWindow, PowerPointSlide, PowerPointSoundEffect, PowerPointTabStop, PowerPointTextStyleLevel, PowerPointTextStyle, PowerPointTimeline, PowerPointTiming, PowerPointTwoInitialCapsException, PowerPointView, PowerPointWebOptions, PowerPointAdjustment, PowerPointCalloutFormat, PowerPointConnectorFormat, PowerPointFillFormat, PowerPointGlowFormat, PowerPointGradientStop, PowerPointLineFormat, PowerPointLinkFormat, PowerPointOfficeTheme, PowerPointPictureFormat, PowerPointPlaceholderFormat, PowerPointReflectionFormat, PowerPointShadowFormat, PowerPointShapeRange, PowerPointShape, PowerPointCallout, PowerPointComment, PowerPointConnector, PowerPointLineShape, PowerPointMediaObject, PowerPointPicture, PowerPointPlaceHolder, PowerPointShapeTable, PowerPointSoftEdgeFormat, PowerPointTextBox, PowerPointTextColumn, PowerPointTextFrame, PowerPointThemeColorScheme, PowerPointThemeColor, PowerPointThemeEffectScheme, PowerPointThemeFontScheme, PowerPointThemeFont, PowerPointMajorThemeFont, PowerPointMinorThemeFont, PowerPointThreeDFormat, PowerPointWordArtFormat, PowerPointTextRange, PowerPointCharacter, PowerPointLine, PowerPointParagraph, PowerPointSentence, PowerPointTextFlow, PowerPointWord, PowerPointCell, PowerPointColumn, PowerPointRow, PowerPointTable;

typedef enum {
    PowerPointSavoYes = 'yes ' /* Save objects now */,
    PowerPointSavoNo = 'no  ' /* Do not save objects */,
    PowerPointSavoAsk = 'ask ' /* Ask the user whether to save */
} PowerPointSavo;

typedef enum {
    PowerPointKfrmIndex = 'indx' /* keyform designating indexed access */,
    PowerPointKfrmNamed = 'name' /* keyform designating named access */,
    PowerPointKfrmId = 'ID  ' /* keyform designating access by unique identifier */
} PowerPointKfrm;

typedef enum {
    PowerPointPPfdMacintoshPath = 'utxt' /* Maintosh path i.e. Foo:bar.txt */,
    PowerPointPPfdPosixPath = 'file' /* Posix path i.e. file:://foo/bar.txt */
} PowerPointPPfd;

typedef enum {
    PowerPointPPffSaveAsPresentation = '\001\000\000\000',
    PowerPointPPffSaveAsTemplate = '\000\000\000\000',
    PowerPointPPffSaveAsRTF = '\001\000\000\000',
    PowerPointPPffSaveAsShow = '\000\000\000\000',
    PowerPointPPffSaveAsDefault = '\001\000\000\000',
    PowerPointPPffSaveAsHTML = '\000\000\000\000',
    PowerPointPPffSaveAsMovie = '\001\000\000\000',
    PowerPointPPffSaveAsPackage = '\000\000\000\000',
    PowerPointPPffSaveAsPDF = '\001\000\000\000',
    PowerPointPPffSaveAsOpenXMLPresentation = '\000\000\000\000',
    PowerPointPPffSaveAsOpenXMLPresentationMacroEnabled = '\001\000\000\000',
    PowerPointPPffSaveAsOpenXMLShow = '\000\000\000\000',
    PowerPointPPffSaveAsOpenXMLShowMacroEnabled = '\001\000\000\000',
    PowerPointPPffSaveAsOpenXMLTemplate = '\000\000\000\000',
    PowerPointPPffSaveAsOpenXMLTemplateMacroEnabled = '\001\000\000\000',
    PowerPointPPffSaveAsOpenXMLTheme = '\000\000\000\000'
} PowerPointPPff;

typedef enum {
    PowerPointMlDsLineDashStyleUnset = '\000\000\000\000',
    PowerPointMlDsLineDashStyleSolid = '\001\000\000\000',
    PowerPointMlDsLineDashStyleSquareDot = '\000\000\000\000',
    PowerPointMlDsLineDashStyleRoundDot = '\001\000\000\000',
    PowerPointMlDsLineDashStyleDash = '\000\000\000\000',
    PowerPointMlDsLineDashStyleDashDot = '\001\000\000\000',
    PowerPointMlDsLineDashStyleDashDotDot = '\000\000\000\000',
    PowerPointMlDsLineDashStyleLongDash = '\001\000\000\000',
    PowerPointMlDsLineDashStyleLongDashDot = '\000\000\000\000',
    PowerPointMlDsLineDashStyleLongDashDotDot = '\001\000\000\000',
    PowerPointMlDsLineDashStyleSystemDash = '\000\000\000\000',
    PowerPointMlDsLineDashStyleSystemDot = '\001\000\000\000',
    PowerPointMlDsLineDashStyleSystemDashDot = '\000\000\000\000'
} PowerPointMlDs;

typedef enum {
    PowerPointMLnSLineStyleUnset = '\000\000\000\000',
    PowerPointMLnSSingleLine = '\001\000\000\000',
    PowerPointMLnSThinThinLine = '\000\000\000\000',
    PowerPointMLnSThinThickLine = '\001\000\000\000',
    PowerPointMLnSThickThinLine = '\000\000\000\000',
    PowerPointMLnSThickBetweenThinLine = '\001\000\000\000'
} PowerPointMLnS;

typedef enum {
    PowerPointMAhSArrowheadStyleUnset = '\001\000\000\000',
    PowerPointMAhSNoArrowhead = '\000\000\000\000',
    PowerPointMAhSTriangleArrowhead = '\001\000\000\000',
    PowerPointMAhSOpen_arrowhead = '\000\000\000\000',
    PowerPointMAhSStealthArrowhead = '\001\000\000\000',
    PowerPointMAhSDiamondArrowhead = '\000\000\000\000',
    PowerPointMAhSOvalArrowhead = '\001\000\000\000'
} PowerPointMAhS;

typedef enum {
    PowerPointMAhWArrowheadWidthUnset = '\001\000\000\000',
    PowerPointMAhWNarrowWidthArrowhead = '\000\000\000\000',
    PowerPointMAhWMediumWidthArrowhead = '\001\000\000\000',
    PowerPointMAhWWideArrowhead = '\000\000\000\000'
} PowerPointMAhW;

typedef enum {
    PowerPointMAhLArrowheadLengthUnset = '\000\000\000\000',
    PowerPointMAhLShortArrowhead = '\001\000\000\000',
    PowerPointMAhLMediumArrowhead = '\000\000\000\000',
    PowerPointMAhLLongArrowhead = '\001\000\000\000'
} PowerPointMAhL;

typedef enum {
    PowerPointMFdTFillUnset = '\001\000\000\000',
    PowerPointMFdTFillSolid = '\000\000\000\000',
    PowerPointMFdTFillPatterned = '\001\000\000\000',
    PowerPointMFdTFillGradient = '\000\000\000\000',
    PowerPointMFdTFillTextured = '\001\000\000\000',
    PowerPointMFdTFillBackground = '\000\000\000\000',
    PowerPointMFdTFillPicture = '\001\000\000\000'
} PowerPointMFdT;

typedef enum {
    PowerPointMGdSGradientUnset = '\001\000\000\000',
    PowerPointMGdSHorizontalGradient = '\000\000\000\000',
    PowerPointMGdSVerticalGradient = '\001\000\000\000',
    PowerPointMGdSDiagonalUpGradient = '\000\000\000\000',
    PowerPointMGdSDiagonalDownGradient = '\001\000\000\000',
    PowerPointMGdSFromCornerGradient = '\000\000\000\000',
    PowerPointMGdSFromTitleGradient = '\001\000\000\000',
    PowerPointMGdSFromCenterGradient = '\000\000\000\000'
} PowerPointMGdS;

typedef enum {
    PowerPointMGCtGradientTypeUnset = '\003\357\377\376',
    PowerPointMGCtSingleShadeGradientType = '\000\000\003\360',
    PowerPointMGCtTwoColorsGradientType = '\000\000\003\360',
    PowerPointMGCtPresetColorsGradientType = '\000\000\003\360',
    PowerPointMGCtMultiColorsGradientType = '\000\000\003\360'
} PowerPointMGCt;

typedef enum {
    PowerPointMxtTTextureTypeTextureTypeUnset = '\003\360\377\376',
    PowerPointMxtTTextureTypePresetTexture = '\000\000\003\361',
    PowerPointMxtTTextureTypeUserDefinedTexture = '\000\000\003\361'
} PowerPointMxtT;

typedef enum {
    PowerPointMPzTPresetTextureUnset = '\000\000\000\000',
    PowerPointMPzTTexturePapyrus = '\001\000\000\000',
    PowerPointMPzTTextureCanvas = '\000\000\000\000',
    PowerPointMPzTTextureDenim = '\001\000\000\000',
    PowerPointMPzTTextureWovenMat = '\000\000\000\000',
    PowerPointMPzTTextureWaterDroplets = '\001\000\000\000',
    PowerPointMPzTTexturePaperBag = '\000\000\000\000',
    PowerPointMPzTTextureFishFossil = '\001\000\000\000',
    PowerPointMPzTTextureSand = '\000\000\000\000',
    PowerPointMPzTTextureGreenMarble = '\001\000\000\000',
    PowerPointMPzTTextureWhiteMarble = '\000\000\000\000',
    PowerPointMPzTTextureBrownMarble = '\001\000\000\000',
    PowerPointMPzTTextureGranite = '\000\000\000\000',
    PowerPointMPzTTextureNewsprint = '\001\000\000\000',
    PowerPointMPzTTextureRecycledPaper = '\000\000\000\000',
    PowerPointMPzTTextureParchment = '\001\000\000\000',
    PowerPointMPzTTextureStationery = '\000\000\000\000',
    PowerPointMPzTTextureBlueTissuePaper = '\001\000\000\000',
    PowerPointMPzTTexturePinkTissuePaper = '\000\000\000\000',
    PowerPointMPzTTexturePurpleMesh = '\001\000\000\000',
    PowerPointMPzTTextureBouquet = '\000\000\000\000',
    PowerPointMPzTTextureCork = '\001\000\000\000',
    PowerPointMPzTTextureWalnut = '\000\000\000\000',
    PowerPointMPzTTextureOak = '\001\000\000\000',
    PowerPointMPzTTextureMediumWood = '\000\000\000\000'
} PowerPointMPzT;

typedef enum {
    PowerPointPpTyPatternUnset = '\000\000\000\000',
    PowerPointPpTyFivePercentPattern = '\001\000\000\000',
    PowerPointPpTyTenPercentPattern = '\000\000\000\000',
    PowerPointPpTyTwentyPercentPattern = '\001\000\000\000',
    PowerPointPpTyTwentyFivePercentPattern = '\000\000\000\000',
    PowerPointPpTyThirtyPercentPattern = '\001\000\000\000',
    PowerPointPpTyFortyPercentPattern = '\000\000\000\000',
    PowerPointPpTyFiftyPercentPattern = '\001\000\000\000',
    PowerPointPpTySixtyPercentPattern = '\000\000\000\000',
    PowerPointPpTySeventyPercentPattern = '\001\000\000\000',
    PowerPointPpTySeventyFivePercentPattern = '\000\000\000\000',
    PowerPointPpTyEightyPercentPattern = '\001\000\000\000',
    PowerPointPpTyNinetyPercentPattern = '\000\000\000\000',
    PowerPointPpTyDarkHorizontalPattern = '\001\000\000\000',
    PowerPointPpTyDarkVerticalPattern = '\000\000\000\000',
    PowerPointPpTyDarkDownwardDiagonalPattern = '\001\000\000\000',
    PowerPointPpTyDarkUpwardDiagonalPattern = '\000\000\000\000',
    PowerPointPpTySmallCheckerBoardPattern = '\001\000\000\000',
    PowerPointPpTyTrellisPattern = '\000\000\000\000',
    PowerPointPpTyLightHorizontalPattern = '\001\000\000\000',
    PowerPointPpTyLightVerticalPattern = '\000\000\000\000',
    PowerPointPpTyLightDownwardDiagonalPattern = '\001\000\000\000',
    PowerPointPpTyLightUpwardDiagonalPattern = '\000\000\000\000',
    PowerPointPpTySmallGridPattern = '\001\000\000\000',
    PowerPointPpTyDottedDiamondPattern = '\000\000\000\000',
    PowerPointPpTyWideDownwardDiagonal = '\001\000\000\000',
    PowerPointPpTyWideUpwardDiagonalPattern = '\000\000\000\000',
    PowerPointPpTyDashedUpwardDiagonalPattern = '\001\000\000\000',
    PowerPointPpTyDashedDownwardDiagonalPattern = '\000\000\000\000',
    PowerPointPpTyNarrowVerticalPattern = '\001\000\000\000',
    PowerPointPpTyNarrowHorizontalPattern = '\000\000\000\000',
    PowerPointPpTyDashedVerticalPattern = '\001\000\000\000',
    PowerPointPpTyDashedHorizontalPattern = '\000\000\000\000',
    PowerPointPpTyLargeConfettiPattern = '\001\000\000\000',
    PowerPointPpTyLargeGridPattern = '\000\000\000\000',
    PowerPointPpTyHorizontalBrickPattern = '\001\000\000\000',
    PowerPointPpTyLargeCheckerBoardPattern = '\000\000\000\000',
    PowerPointPpTySmallConfettiPattern = '\001\000\000\000',
    PowerPointPpTyZigZagPattern = '\000\000\000\000',
    PowerPointPpTySolidDiamondPattern = '\001\000\000\000',
    PowerPointPpTyDiagonalBrickPattern = '\000\000\000\000',
    PowerPointPpTyOutlinedDiamondPattern = '\001\000\000\000',
    PowerPointPpTyPlaidPattern = '\000\000\000\000',
    PowerPointPpTySpherePattern = '\001\000\000\000',
    PowerPointPpTyWeavePattern = '\000\000\000\000',
    PowerPointPpTyDottedGridPattern = '\001\000\000\000',
    PowerPointPpTyDivotPattern = '\000\000\000\000',
    PowerPointPpTyShinglePattern = '\001\000\000\000',
    PowerPointPpTyWavePattern = '\000\000\000\000',
    PowerPointPpTyHorizontalPattern = '\001\000\000\000',
    PowerPointPpTyVerticalPattern = '\000\000\000\000',
    PowerPointPpTyCrossPattern = '\001\000\000\000',
    PowerPointPpTyDownwardDiagonalPattern = '\000\000\000\000',
    PowerPointPpTyUpwardDiagonalPattern = '\001\000\000\000',
    PowerPointPpTyDiagonalCrossPattern = '\000\000\000\000'
} PowerPointPpTy;

typedef enum {
    PowerPointMPGbPresetGradientUnset = '\000\000\000\000',
    PowerPointMPGbGradientEarlySunset = '\001\000\000\000',
    PowerPointMPGbGradientLateSunset = '\000\000\000\000',
    PowerPointMPGbGradientNightfall = '\001\000\000\000',
    PowerPointMPGbGradientDaybreak = '\000\000\000\000',
    PowerPointMPGbGradientHorizon = '\001\000\000\000',
    PowerPointMPGbGradientDesert = '\000\000\000\000',
    PowerPointMPGbGradientOcean = '\001\000\000\000',
    PowerPointMPGbGradientCalmWater = '\000\000\000\000',
    PowerPointMPGbGradientFire = '\001\000\000\000',
    PowerPointMPGbGradientFog = '\000\000\000\000',
    PowerPointMPGbGradientMoss = '\001\000\000\000',
    PowerPointMPGbGradientPeacock = '\000\000\000\000',
    PowerPointMPGbGradientWheat = '\001\000\000\000',
    PowerPointMPGbGradientParchment = '\000\000\000\000',
    PowerPointMPGbGradientMahogany = '\001\000\000\000',
    PowerPointMPGbGradientRainbow = '\000\000\000\000',
    PowerPointMPGbGradientRainbow2 = '\001\000\000\000',
    PowerPointMPGbGradientGold = '\000\000\000\000',
    PowerPointMPGbGradientGold2 = '\001\000\000\000',
    PowerPointMPGbGradientBrass = '\000\000\000\000',
    PowerPointMPGbGradientChrome = '\001\000\000\000',
    PowerPointMPGbGradientChrome2 = '\000\000\000\000',
    PowerPointMPGbGradientSilver = '\001\000\000\000',
    PowerPointMPGbGradientSapphire = '\000\000\000\000'
} PowerPointMPGb;

typedef enum {
    PowerPointMSdTShadowUnset = '\003_\377\376',
    PowerPointMSdTShadow1 = '\000\000\003`',
    PowerPointMSdTShadow2 = '\000\000\003`',
    PowerPointMSdTShadow3 = '\000\000\003`',
    PowerPointMSdTShadow4 = '\000\000\003`',
    PowerPointMSdTShadow5 = '\000\000\003`',
    PowerPointMSdTShadow6 = '\000\000\003`',
    PowerPointMSdTShadow7 = '\000\000\003`',
    PowerPointMSdTShadow8 = '\000\000\003`',
    PowerPointMSdTShadow9 = '\000\000\003`',
    PowerPointMSdTShadow10 = '\000\000\003`',
    PowerPointMSdTShadow11 = '\000\000\003`',
    PowerPointMSdTShadow12 = '\000\000\003`',
    PowerPointMSdTShadow13 = '\000\000\003`',
    PowerPointMSdTShadow14 = '\000\000\003`',
    PowerPointMSdTShadow15 = '\000\000\003`',
    PowerPointMSdTShadow16 = '\000\000\003`',
    PowerPointMSdTShadow17 = '\000\000\003`',
    PowerPointMSdTShadow18 = '\000\000\003`',
    PowerPointMSdTShadow19 = '\000\000\003`',
    PowerPointMSdTShadow20 = '\000\000\003`'
} PowerPointMSdT;

typedef enum {
    PowerPointMPXFWordartFormatUnset = '\003\361\377\376',
    PowerPointMPXFWordartFormat1 = '\000\000\003\362',
    PowerPointMPXFWordartFormat2 = '\000\000\003\362',
    PowerPointMPXFWordartFormat3 = '\000\000\003\362',
    PowerPointMPXFWordartFormat4 = '\000\000\003\362',
    PowerPointMPXFWordartFormat5 = '\000\000\003\362',
    PowerPointMPXFWordartFormat6 = '\000\000\003\362',
    PowerPointMPXFWordartFormat7 = '\000\000\003\362',
    PowerPointMPXFWordartFormat8 = '\000\000\003\362',
    PowerPointMPXFWordartFormat9 = '\000\000\003\362',
    PowerPointMPXFWordartFormat10 = '\000\000\003\362',
    PowerPointMPXFWordartFormat11 = '\000\000\003\362',
    PowerPointMPXFWordartFormat12 = '\000\000\003\362',
    PowerPointMPXFWordartFormat13 = '\000\000\003\362',
    PowerPointMPXFWordartFormat14 = '\000\000\003\362',
    PowerPointMPXFWordartFormat15 = '\000\000\003\362',
    PowerPointMPXFWordartFormat16 = '\000\000\003\362',
    PowerPointMPXFWordartFormat17 = '\000\000\003\362',
    PowerPointMPXFWordartFormat18 = '\000\000\003\362',
    PowerPointMPXFWordartFormat19 = '\000\000\003\362',
    PowerPointMPXFWordartFormat20 = '\000\000\003\362',
    PowerPointMPXFWordartFormat21 = '\000\000\003\362',
    PowerPointMPXFWordartFormat22 = '\000\000\003\362',
    PowerPointMPXFWordartFormat23 = '\000\000\003\362',
    PowerPointMPXFWordartFormat24 = '\000\000\003\362',
    PowerPointMPXFWordartFormat25 = '\000\000\003\362',
    PowerPointMPXFWordartFormat26 = '\000\000\003\362',
    PowerPointMPXFWordartFormat27 = '\000\000\003\362',
    PowerPointMPXFWordartFormat28 = '\000\000\003\362',
    PowerPointMPXFWordartFormat29 = '\000\000\003\362',
    PowerPointMPXFWordartFormat30 = '\000\000\003\362'
} PowerPointMPXF;

typedef enum {
    PowerPointMPTsTextEffectShapeUnset = '\000\000\000\000',
    PowerPointMPTsPlainText = '\001\000\000\000',
    PowerPointMPTsStop = '\000\000\000\000',
    PowerPointMPTsTriangleUp = '\001\000\000\000',
    PowerPointMPTsTriangleDown = '\000\000\000\000',
    PowerPointMPTsChevronUp = '\001\000\000\000',
    PowerPointMPTsChevronDown = '\000\000\000\000',
    PowerPointMPTsRingInside = '\001\000\000\000',
    PowerPointMPTsRingOutside = '\000\000\000\000',
    PowerPointMPTsArchUpCurve = '\001\000\000\000',
    PowerPointMPTsArchDownCurve = '\000\000\000\000',
    PowerPointMPTsCircleCurve = '\001\000\000\000',
    PowerPointMPTsButtonCurve = '\000\000\000\000',
    PowerPointMPTsArchUpPour = '\001\000\000\000',
    PowerPointMPTsArchDownPour = '\000\000\000\000',
    PowerPointMPTsCirclePour = '\001\000\000\000',
    PowerPointMPTsButtonPour = '\000\000\000\000',
    PowerPointMPTsCurveUp = '\001\000\000\000',
    PowerPointMPTsCurveDown = '\000\000\000\000',
    PowerPointMPTsCanUp = '\001\000\000\000',
    PowerPointMPTsCanDown = '\000\000\000\000',
    PowerPointMPTsWave1 = '\001\000\000\000',
    PowerPointMPTsWave2 = '\000\000\000\000',
    PowerPointMPTsDoubleWave1 = '\001\000\000\000',
    PowerPointMPTsDoubleWave2 = '\000\000\000\000',
    PowerPointMPTsInflate = '\001\000\000\000',
    PowerPointMPTsDeflate = '\000\000\000\000',
    PowerPointMPTsInflateBottom = '\001\000\000\000',
    PowerPointMPTsDeflateBottom = '\000\000\000\000',
    PowerPointMPTsInflateTop = '\001\000\000\000',
    PowerPointMPTsDeflateTop = '\000\000\000\000',
    PowerPointMPTsDeflateInflate = '\001\000\000\000',
    PowerPointMPTsDeflateInflateDeflate = '\000\000\000\000',
    PowerPointMPTsFadeRight = '\001\000\000\000',
    PowerPointMPTsFadeLeft = '\000\000\000\000',
    PowerPointMPTsFadeUp = '\001\000\000\000',
    PowerPointMPTsFadeDown = '\000\000\000\000',
    PowerPointMPTsSlantUp = '\001\000\000\000',
    PowerPointMPTsSlantDown = '\000\000\000\000',
    PowerPointMPTsCascadeUp = '\001\000\000\000',
    PowerPointMPTsCascadeDown = '\000\000\000\000'
} PowerPointMPTs;

typedef enum {
    PowerPointMTxATextEffectAlignmentUnset = '\000\000\000\000',
    PowerPointMTxALeftTextEffectAlignment = '\001\000\000\000',
    PowerPointMTxACenteredTextEffectAlignment = '\000\000\000\000',
    PowerPointMTxARightTextEffectAlignment = '\001\000\000\000',
    PowerPointMTxAJustifyTextEffectAlignment = '\000\000\000\000',
    PowerPointMTxAWordJustifyTextEffectAlignment = '\001\000\000\000',
    PowerPointMTxAStretchJustifyTextEffectAlignment = '\000\000\000\000'
} PowerPointMTxA;

typedef enum {
    PowerPointMPLdPresetLightingDirectionUnset = '\000\000\000\000',
    PowerPointMPLdLightFromTopLeft = '\001\000\000\000',
    PowerPointMPLdLightFromTop = '\000\000\000\000',
    PowerPointMPLdLightFromTopRight = '\001\000\000\000',
    PowerPointMPLdLightFromLeft = '\000\000\000\000',
    PowerPointMPLdLightFromNone = '\001\000\000\000',
    PowerPointMPLdLightFromRight = '\000\000\000\000',
    PowerPointMPLdLightFromBottomLeft = '\001\000\000\000',
    PowerPointMPLdLightFromBottom = '\000\000\000\000',
    PowerPointMPLdLightFromBottomRight = '\001\000\000\000'
} PowerPointMPLd;

typedef enum {
    PowerPointMlSfLightingSoftnessUnset = '\001\000\000\000',
    PowerPointMlSfLightingDim = '\000\000\000\000',
    PowerPointMlSfLightingNormal = '\001\000\000\000',
    PowerPointMlSfLightingBright = '\000\000\000\000'
} PowerPointMlSf;

typedef enum {
    PowerPointMPMtPresetMaterialUnset = '\000\000\000\000',
    PowerPointMPMtMatte = '\001\000\000\000',
    PowerPointMPMtPlastic = '\000\000\000\000',
    PowerPointMPMtMetal = '\001\000\000\000',
    PowerPointMPMtWireframe = '\000\000\000\000',
    PowerPointMPMtMatte2 = '\001\000\000\000',
    PowerPointMPMtPlastic2 = '\000\000\000\000',
    PowerPointMPMtMetal2 = '\001\000\000\000',
    PowerPointMPMtWarmMatte = '\000\000\000\000',
    PowerPointMPMtTranslucentPowder = '\001\000\000\000',
    PowerPointMPMtPowder = '\000\000\000\000',
    PowerPointMPMtDarkEdge = '\001\000\000\000',
    PowerPointMPMtSoftEdge = '\000\000\000\000',
    PowerPointMPMtMaterialClear = '\001\000\000\000',
    PowerPointMPMtFlat = '\000\000\000\000',
    PowerPointMPMtSoftMetal = '\001\000\000\000'
} PowerPointMPMt;

typedef enum {
    PowerPointMExDPresetExtrusionDirectionUnset = '\001\000\000\000',
    PowerPointMExDExtrudeBottomRight = '\000\000\000\000',
    PowerPointMExDExtrudeBottom = '\001\000\000\000',
    PowerPointMExDExtrudeBottomLeft = '\000\000\000\000',
    PowerPointMExDExtrudeRight = '\001\000\000\000',
    PowerPointMExDExtrudeNone = '\000\000\000\000',
    PowerPointMExDExtrudeLeft = '\001\000\000\000',
    PowerPointMExDExtrudeTopRight = '\000\000\000\000',
    PowerPointMExDExtrudeTop = '\001\000\000\000',
    PowerPointMExDExtrudeTopLeft = '\000\000\000\000'
} PowerPointMExD;

typedef enum {
    PowerPointM3DFPresetThreeDFormatUnset = '\000\000\000\000',
    PowerPointM3DFFormat1 = '\001\000\000\000',
    PowerPointM3DFFormat2 = '\000\000\000\000',
    PowerPointM3DFFormat3 = '\001\000\000\000',
    PowerPointM3DFFormat4 = '\000\000\000\000',
    PowerPointM3DFFormat5 = '\001\000\000\000',
    PowerPointM3DFFormat6 = '\000\000\000\000',
    PowerPointM3DFFormat7 = '\001\000\000\000',
    PowerPointM3DFFormat8 = '\000\000\000\000',
    PowerPointM3DFFormat9 = '\001\000\000\000',
    PowerPointM3DFFormat10 = '\000\000\000\000',
    PowerPointM3DFFormat11 = '\001\000\000\000',
    PowerPointM3DFFormat12 = '\000\000\000\000',
    PowerPointM3DFFormat13 = '\001\000\000\000',
    PowerPointM3DFFormat14 = '\000\000\000\000',
    PowerPointM3DFFormat15 = '\001\000\000\000',
    PowerPointM3DFFormat16 = '\000\000\000\000',
    PowerPointM3DFFormat17 = '\001\000\000\000',
    PowerPointM3DFFormat18 = '\000\000\000\000',
    PowerPointM3DFFormat19 = '\001\000\000\000',
    PowerPointM3DFFormat20 = '\000\000\000\000'
} PowerPointM3DF;

typedef enum {
    PowerPointMExCExtrusionColorUnset = '\000\000\000\000',
    PowerPointMExCExtrusionColorAutomatic = '\001\000\000\000',
    PowerPointMExCExtrusionColorCustom = '\000\000\000\000'
} PowerPointMExC;

typedef enum {
    PowerPointMCtTConnectorTypeUnset = '\000\000\000\000',
    PowerPointMCtTStraight = '\001\000\000\000',
    PowerPointMCtTElbow = '\000\000\000\000',
    PowerPointMCtTCurve = '\001\000\000\000'
} PowerPointMCtT;

typedef enum {
    PowerPointMHzAHorizontalAnchorUnset = '\001\000\000\000',
    PowerPointMHzAHorizontalAnchorNone = '\000\000\000\000',
    PowerPointMHzAHorizontalAnchorCenter = '\001\000\000\000'
} PowerPointMHzA;

typedef enum {
    PowerPointMVtAVerticalAnchorUnset = '\001\000\000\000',
    PowerPointMVtAAnchorTop = '\000\000\000\000',
    PowerPointMVtAAnchorTopBaseline = '\001\000\000\000',
    PowerPointMVtAAnchorMiddle = '\000\000\000\000',
    PowerPointMVtAAnchorBottom = '\001\000\000\000',
    PowerPointMVtAAnchorBottomBaseline = '\000\000\000\000'
} PowerPointMVtA;

typedef enum {
    PowerPointMAsTAutoshapeShapeTypeUnset = '\000\000\000\000',
    PowerPointMAsTAutoshapeRectangle = '\001\000\000\000',
    PowerPointMAsTAutoshapeParallelogram = '\000\000\000\000',
    PowerPointMAsTAutoshapeTrapezoid = '\001\000\000\000',
    PowerPointMAsTAutoshapeDiamond = '\000\000\000\000',
    PowerPointMAsTAutoshapeRoundedRectangle = '\001\000\000\000',
    PowerPointMAsTAutoshapeOctagon = '\000\000\000\000',
    PowerPointMAsTAutoshapeIsoscelesTriangle = '\001\000\000\000',
    PowerPointMAsTAutoshapeRightTriangle = '\000\000\000\000',
    PowerPointMAsTAutoshapeOval = '\001\000\000\000',
    PowerPointMAsTAutoshapeHexagon = '\000\000\000\000',
    PowerPointMAsTAutoshapeCross = '\001\000\000\000',
    PowerPointMAsTAutoshapeRegularPentagon = '\000\000\000\000',
    PowerPointMAsTAutoshapeCan = '\001\000\000\000',
    PowerPointMAsTAutoshapeCube = '\000\000\000\000',
    PowerPointMAsTAutoshapeBevel = '\001\000\000\000',
    PowerPointMAsTAutoshapeFoldedCorner = '\000\000\000\000',
    PowerPointMAsTAutoshapeSmileyFace = '\001\000\000\000',
    PowerPointMAsTAutoshapeDonut = '\000\000\000\000',
    PowerPointMAsTAutoshapeNoSymbol = '\001\000\000\000',
    PowerPointMAsTAutoshapeBlockArc = '\000\000\000\000',
    PowerPointMAsTAutoshapeHeart = '\001\000\000\000',
    PowerPointMAsTAutoshapeLightningBolt = '\000\000\000\000',
    PowerPointMAsTAutoshapeSun = '\001\000\000\000',
    PowerPointMAsTAutoshapeMoon = '\000\000\000\000',
    PowerPointMAsTAutoshapeArc = '\001\000\000\000',
    PowerPointMAsTAutoshapeDoubleBracket = '\000\000\000\000',
    PowerPointMAsTAutoshapeDoubleBrace = '\001\000\000\000',
    PowerPointMAsTAutoshapePlaque = '\000\000\000\000',
    PowerPointMAsTAutoshapeLeftBracket = '\001\000\000\000',
    PowerPointMAsTAutoshapeRightBracket = '\000\000\000\000',
    PowerPointMAsTAutoshapeLeftBrace = '\001\000\000\000',
    PowerPointMAsTAutoshapeRightBrace = '\000\000\000\000',
    PowerPointMAsTAutoshapeRightArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeLeftArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeUpArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeDownArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeLeftRightArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeUpDownArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeQuadArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeLeftRightUpArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeBentArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeUTurnArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeLeftUpArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeBentUpArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeCurvedRightArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeCurvedLeftArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeCurvedUpArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeCurvedDownArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeStripedRightArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeNotchedRightArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapePentagon = '\001\000\000\000',
    PowerPointMAsTAutoshapeChevron = '\000\000\000\000',
    PowerPointMAsTAutoshapeRightArrowCallout = '\001\000\000\000',
    PowerPointMAsTAutoshapeLeftArrowCallout = '\000\000\000\000',
    PowerPointMAsTAutoshapeUpArrowCallout = '\001\000\000\000',
    PowerPointMAsTAutoshapeDownArrowCallout = '\000\000\000\000',
    PowerPointMAsTAutoshapeLeftRightArrowCallout = '\001\000\000\000',
    PowerPointMAsTAutoshapeUpDownArrowCallout = '\000\000\000\000',
    PowerPointMAsTAutoshapeQuadArrowCallout = '\001\000\000\000',
    PowerPointMAsTAutoshapeCircularArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartProcess = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartAlternateProcess = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartDecision = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartData = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartPredefinedProcess = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartInternalStorage = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartDocument = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartMultiDocument = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartTerminator = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartPreparation = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartManualInput = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartManualOperation = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartConnector = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartOffpageConnector = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartCard = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartPunchedTape = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartSummingJunction = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartOr = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartCollate = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartSort = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartExtract = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartMerge = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartStoredData = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartDelay = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartSequentialAccessStorage = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartMagneticDisk = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartDirectAccessStorage = '\001\000\000\000',
    PowerPointMAsTAutoshapeFlowchartDisplay = '\000\000\000\000',
    PowerPointMAsTAutoshapeExplosionOne = '\001\000\000\000',
    PowerPointMAsTAutoshapeExplosionTwo = '\000\000\000\000',
    PowerPointMAsTAutoshapeFourPointStar = '\001\000\000\000',
    PowerPointMAsTAutoshapeFivePointStar = '\000\000\000\000',
    PowerPointMAsTAutoshapeEightPointStar = '\001\000\000\000',
    PowerPointMAsTAutoshapeSixteenPointStar = '\000\000\000\000',
    PowerPointMAsTAutoshapeTwentyFourPointStar = '\001\000\000\000',
    PowerPointMAsTAutoshapeThirtyTwoPointStar = '\000\000\000\000',
    PowerPointMAsTAutoshapeUpRibbon = '\001\000\000\000',
    PowerPointMAsTAutoshapeDownRibbon = '\000\000\000\000',
    PowerPointMAsTAutoshapeCurvedUpRibbon = '\001\000\000\000',
    PowerPointMAsTAutoshapeCurvedDownRibbon = '\000\000\000\000',
    PowerPointMAsTAutoshapeVerticalScroll = '\001\000\000\000',
    PowerPointMAsTAutoshapeHorizontalScroll = '\000\000\000\000',
    PowerPointMAsTAutoshapeWave = '\001\000\000\000',
    PowerPointMAsTAutoshapeDoubleWave = '\000\000\000\000',
    PowerPointMAsTAutoshapeRectangularCallout = '\001\000\000\000',
    PowerPointMAsTAutoshapeRoundedRectangularCallout = '\000\000\000\000',
    PowerPointMAsTAutoshapeOvalCallout = '\001\000\000\000',
    PowerPointMAsTAutoshapeCloudCallout = '\000\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutOne = '\001\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutTwo = '\000\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutThree = '\001\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutFour = '\000\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutOneAccentBar = '\001\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutTwoAccentBar = '\000\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutThreeAccentBar = '\001\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutFourAccentBar = '\000\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutOneNoBorder = '\001\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutTwoNoBorder = '\000\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutThreeNoBorder = '\001\000\000\000',
    PowerPointMAsTAutoshapeLineCalloutFourNoBorder = '\000\000\000\000',
    PowerPointMAsTAutoshapeCalloutOneBorderAndAccentBar = '\001\000\000\000',
    PowerPointMAsTAutoshapeCalloutTwoBorderAndAccentBar = '\000\000\000\000',
    PowerPointMAsTAutoshapeCalloutThreeBorderAndAccentBar = '\001\000\000\000',
    PowerPointMAsTAutoshapeCalloutFourBorderAndAccentBar = '\000\000\000\000',
    PowerPointMAsTAutoshapeActionButtonCustom = '\001\000\000\000',
    PowerPointMAsTAutoshapeActionButtonHome = '\000\000\000\000',
    PowerPointMAsTAutoshapeActionButtonHelp = '\001\000\000\000',
    PowerPointMAsTAutoshapeActionButtonInformation = '\000\000\000\000',
    PowerPointMAsTAutoshapeActionButtonBackOrPrevious = '\001\000\000\000',
    PowerPointMAsTAutoshapeActionButtonForwardOrNext = '\000\000\000\000',
    PowerPointMAsTAutoshapeActionButtonBeginning = '\001\000\000\000',
    PowerPointMAsTAutoshapeActionButtonEnd = '\000\000\000\000',
    PowerPointMAsTAutoshapeActionButtonReturn = '\001\000\000\000',
    PowerPointMAsTAutoshapeActionButtonDocument = '\000\000\000\000',
    PowerPointMAsTAutoshapeActionButtonSound = '\001\000\000\000',
    PowerPointMAsTAutoshapeActionButtonMovie = '\000\000\000\000',
    PowerPointMAsTAutoshapeBalloon = '\001\000\000\000',
    PowerPointMAsTAutoshapeNotPrimitive = '\000\000\000\000',
    PowerPointMAsTAutoshapeFlowchartOfflineStorage = '\001\000\000\000',
    PowerPointMAsTAutoshapeLeftRightRibbon = '\000\000\000\000',
    PowerPointMAsTAutoshapeDiagonalStripe = '\001\000\000\000',
    PowerPointMAsTAutoshapePie = '\000\000\000\000',
    PowerPointMAsTAutoshapeNonIsoscelesTrapezoid = '\001\000\000\000',
    PowerPointMAsTAutoshapeDecagon = '\000\000\000\000',
    PowerPointMAsTAutoshapeHeptagon = '\001\000\000\000',
    PowerPointMAsTAutoshapeDodecagon = '\000\000\000\000',
    PowerPointMAsTAutoshapeSixPointsStar = '\001\000\000\000',
    PowerPointMAsTAutoshapeSevenPointsStar = '\000\000\000\000',
    PowerPointMAsTAutoshapeTenPointsStar = '\001\000\000\000',
    PowerPointMAsTAutoshapeTwelvePointsStar = '\000\000\000\000',
    PowerPointMAsTAutoshapeRoundOneRectangle = '\001\000\000\000',
    PowerPointMAsTAutoshapeRoundTwoSameRectangle = '\000\000\000\000',
    PowerPointMAsTAutoshapeRoundTwoDiagonalRectangle = '\001\000\000\000',
    PowerPointMAsTAutoshapeSnipRoundRectangle = '\000\000\000\000',
    PowerPointMAsTAutoshapeSnipOneRectangle = '\001\000\000\000',
    PowerPointMAsTAutoshapeSnipTwoSameRectangle = '\000\000\000\000',
    PowerPointMAsTAutoshapeSnipTwoDiagonalRectangle = '\001\000\000\000',
    PowerPointMAsTAutoshapeFrame = '\000\000\000\000',
    PowerPointMAsTAutoshapeHalfFrame = '\001\000\000\000',
    PowerPointMAsTAutoshapeTear = '\000\000\000\000',
    PowerPointMAsTAutoshapeChord = '\001\000\000\000',
    PowerPointMAsTAutoshapeCorner = '\000\000\000\000',
    PowerPointMAsTAutoshapeMathPlus = '\001\000\000\000',
    PowerPointMAsTAutoshapeMathMinus = '\000\000\000\000',
    PowerPointMAsTAutoshapeMathMultiply = '\001\000\000\000',
    PowerPointMAsTAutoshapeMathDivide = '\000\000\000\000',
    PowerPointMAsTAutoshapeMathEqual = '\001\000\000\000',
    PowerPointMAsTAutoshapeMathNotEqual = '\000\000\000\000',
    PowerPointMAsTAutoshapeCornerTabs = '\001\000\000\000',
    PowerPointMAsTAutoshapeSquareTabs = '\000\000\000\000',
    PowerPointMAsTAutoshapePlaqueTabs = '\001\000\000\000',
    PowerPointMAsTAutoshapeGearSix = '\000\000\000\000',
    PowerPointMAsTAutoshapeGearNine = '\001\000\000\000',
    PowerPointMAsTAutoshapeFunnel = '\000\000\000\000',
    PowerPointMAsTAutoshapePieWedge = '\001\000\000\000',
    PowerPointMAsTAutoshapeLeftCircularArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeLeftRightCircularArrow = '\001\000\000\000',
    PowerPointMAsTAutoshapeSwooshArrow = '\000\000\000\000',
    PowerPointMAsTAutoshapeCloud = '\001\000\000\000',
    PowerPointMAsTAutoshapeChartX = '\000\000\000\000',
    PowerPointMAsTAutoshapeChartStar = '\001\000\000\000',
    PowerPointMAsTAutoshapeChartPlus = '\000\000\000\000',
    PowerPointMAsTAutoshapeLineInverse = '\001\000\000\000'
} PowerPointMAsT;

typedef enum {
    PowerPointMShpShapeTypeUnset = '\001\000\000\000',
    PowerPointMShpShapeTypeAuto = '\000\000\000\000',
    PowerPointMShpShapeTypeCallout = '\001\000\000\000',
    PowerPointMShpShapeTypeChart = '\000\000\000\000',
    PowerPointMShpShapeTypeComment = '\001\000\000\000',
    PowerPointMShpShapeTypeFreeForm = '\000\000\000\000',
    PowerPointMShpShapeTypeGroup = '\001\000\000\000',
    PowerPointMShpShapeTypeEmbeddedOLEControl = '\000\000\000\000',
    PowerPointMShpShapeTypeFormControl = '\001\000\000\000',
    PowerPointMShpShapeTypeLine = '\000\000\000\000',
    PowerPointMShpShapeTypeLinkedOLEObject = '\001\000\000\000',
    PowerPointMShpShapeTypeLinkedPicture = '\000\000\000\000',
    PowerPointMShpShapeTypeOLEControl = '\001\000\000\000',
    PowerPointMShpShapeTypePicture = '\000\000\000\000',
    PowerPointMShpShapeTypePlaceHolder = '\001\000\000\000',
    PowerPointMShpShapeTypeWordArt = '\000\000\000\000',
    PowerPointMShpShapeTypeMedia = '\001\000\000\000',
    PowerPointMShpShapeTypeTextBox = '\000\000\000\000',
    PowerPointMShpShapeTypeTable = '\001\000\000\000',
    PowerPointMShpShapeTypeSmartartGraphic = '\000\000\000\000'
} PowerPointMShp;

typedef enum {
    PowerPointMCrTColorTypeUnset = '\000\000\000\000',
    PowerPointMCrTRGB = '\001\000\000\000',
    PowerPointMCrTScheme = '\000\000\000\000'
} PowerPointMCrT;

typedef enum {
    PowerPointMPcPictureColorTypeUnset = '\000\000\000\000',
    PowerPointMPcPictureColorAutomatic = '\001\000\000\000',
    PowerPointMPcPictureColorGrayScale = '\000\000\000\000',
    PowerPointMPcPictureColorBlackAndWhite = '\001\000\000\000',
    PowerPointMPcPictureColorWatermark = '\000\000\000\000'
} PowerPointMPc;

typedef enum {
    PowerPointMCAtAngleUnset = '\000\000\000\000',
    PowerPointMCAtAngleAutomatic = '\001\000\000\000',
    PowerPointMCAtAngle30 = '\000\000\000\000',
    PowerPointMCAtAngle45 = '\001\000\000\000',
    PowerPointMCAtAngle60 = '\000\000\000\000',
    PowerPointMCAtAngle90 = '\001\000\000\000'
} PowerPointMCAt;

typedef enum {
    PowerPointMCDtDropUnset = '\001\000\000\000',
    PowerPointMCDtDropCustom = '\000\000\000\000',
    PowerPointMCDtDropTop = '\001\000\000\000',
    PowerPointMCDtDropCenter = '\000\000\000\000',
    PowerPointMCDtDropBottom = '\001\000\000\000'
} PowerPointMCDt;

typedef enum {
    PowerPointMCotCalloutUnset = '\001\000\000\000',
    PowerPointMCotCalloutOne = '\000\000\000\000',
    PowerPointMCotCalloutTwo = '\001\000\000\000',
    PowerPointMCotCalloutThree = '\000\000\000\000',
    PowerPointMCotCalloutFour = '\001\000\000\000'
} PowerPointMCot;

typedef enum {
    PowerPointTxOrTextOrientationUnset = '\001\000\000\000',
    PowerPointTxOrHorizontal = '\000\000\000\000',
    PowerPointTxOrUpward = '\001\000\000\000',
    PowerPointTxOrDownward = '\000\000\000\000',
    PowerPointTxOrVerticalEastAsian = '\001\000\000\000',
    PowerPointTxOrVertical = '\000\000\000\000',
    PowerPointTxOrHorizontalRotatedEastAsian = '\001\000\000\000'
} PowerPointTxOr;

typedef enum {
    PowerPointMSFrScaleFromTopLeft = '\001\000\000\000',
    PowerPointMSFrScaleFromMiddle = '\000\000\000\000',
    PowerPointMSFrScaleFromBottomRight = '\001\000\000\000'
} PowerPointMSFr;

typedef enum {
    PowerPointMPzCPresetCameraUnset = '\001\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromTopLeft = '\000\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromTop = '\001\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromTopright = '\000\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromLeft = '\001\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromFront = '\000\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromRight = '\001\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromBottomLeft = '\000\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromBottom = '\001\000\000\000',
    PowerPointMPzCCameraLegacyObliqueFromBottomRight = '\000\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromTopLeft = '\001\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromTop = '\000\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromTopRight = '\001\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromLeft = '\000\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromFront = '\001\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromRight = '\000\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromBottomLeft = '\001\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromBottom = '\000\000\000\000',
    PowerPointMPzCCameraLegacyPerspectiveFromBottomRight = '\001\000\000\000',
    PowerPointMPzCCameraOrthographic = '\000\000\000\000',
    PowerPointMPzCCameraIsometricFromTopUp = '\001\000\000\000',
    PowerPointMPzCCameraIsometricFromTopDown = '\000\000\000\000',
    PowerPointMPzCCameraIsometricFromBottomUp = '\001\000\000\000',
    PowerPointMPzCCameraIsometricFromBottomDown = '\000\000\000\000',
    PowerPointMPzCCameraIsometricFromLeftUp = '\001\000\000\000',
    PowerPointMPzCCameraIsometricFromLeftDown = '\000\000\000\000',
    PowerPointMPzCCameraIsometricFromRightUp = '\001\000\000\000',
    PowerPointMPzCCameraIsometricFromRightDown = '\000\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis1FromLeft = '\001\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis1FromRight = '\000\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis1FromTop = '\001\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis2FromLeft = '\000\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis2FromRight = '\001\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis2FromTop = '\000\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis3FromLeft = '\001\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis3FromRight = '\000\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis3FromBottom = '\001\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis4FromLeft = '\000\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis4FromRight = '\001\000\000\000',
    PowerPointMPzCCameraIsometricOffAxis4FromBottom = '\000\000\000\000',
    PowerPointMPzCCameraObliqueFromTopLeft = '\001\000\000\000',
    PowerPointMPzCCameraObliqueFromTop = '\000\000\000\000',
    PowerPointMPzCCameraObliqueFromTopRight = '\001\000\000\000',
    PowerPointMPzCCameraObliqueFromLeft = '\000\000\000\000',
    PowerPointMPzCCameraObliqueFromRight = '\001\000\000\000',
    PowerPointMPzCCameraObliqueFromBottomLeft = '\000\000\000\000',
    PowerPointMPzCCameraObliqueFromBottom = '\001\000\000\000',
    PowerPointMPzCCameraObliqueFromBottomRight = '\000\000\000\000',
    PowerPointMPzCCameraPerspectiveFromFront = '\001\000\000\000',
    PowerPointMPzCCameraPerspectiveFromLeft = '\000\000\000\000',
    PowerPointMPzCCameraPerspectiveFromRight = '\001\000\000\000',
    PowerPointMPzCCameraPerspectiveFromAbove = '\000\000\000\000',
    PowerPointMPzCCameraPerspectiveFromBelow = '\001\000\000\000',
    PowerPointMPzCCameraPerspectiveFromAboveFacingLeft = '\000\000\000\000',
    PowerPointMPzCCameraPerspectiveFromAboveFacingRight = '\001\000\000\000',
    PowerPointMPzCCameraPerspectiveContrastingFacingLeft = '\000\000\000\000',
    PowerPointMPzCCameraPerspectiveContrastingFacingRight = '\001\000\000\000',
    PowerPointMPzCCameraPerspectiveHeroicFacingLeft = '\000\000\000\000',
    PowerPointMPzCCameraPerspectiveHeroicFacingRight = '\001\000\000\000',
    PowerPointMPzCCameraPerspectiveHeroicExtremeFacingLeft = '\000\000\000\000',
    PowerPointMPzCCameraPerspectiveHeroicExtremeFacingRight = '\001\000\000\000',
    PowerPointMPzCCameraPerspectiveRelaxed = '\000\000\000\000',
    PowerPointMPzCCameraPerspectiveRelaxedModerately = '\001\000\000\000'
} PowerPointMPzC;

typedef enum {
    PowerPointMLtTLightRigUnset = '\001\000\000\000',
    PowerPointMLtTLightRigFlat1 = '\000\000\000\000',
    PowerPointMLtTLightRigFlat2 = '\001\000\000\000',
    PowerPointMLtTLightRigFlat3 = '\000\000\000\000',
    PowerPointMLtTLightRigFlat4 = '\001\000\000\000',
    PowerPointMLtTLightRigNormal1 = '\000\000\000\000',
    PowerPointMLtTLightRigNormal2 = '\001\000\000\000',
    PowerPointMLtTLightRigNormal3 = '\000\000\000\000',
    PowerPointMLtTLightRigNormal4 = '\001\000\000\000',
    PowerPointMLtTLightRigHarsh1 = '\000\000\000\000',
    PowerPointMLtTLightRigHarsh2 = '\001\000\000\000',
    PowerPointMLtTLightRigHarsh3 = '\000\000\000\000',
    PowerPointMLtTLightRigHarsh4 = '\001\000\000\000',
    PowerPointMLtTLightRigThreePoint = '\000\000\000\000',
    PowerPointMLtTLightRigBalanced = '\001\000\000\000',
    PowerPointMLtTLightRigSoft = '\000\000\000\000',
    PowerPointMLtTLightRigHarsh = '\001\000\000\000',
    PowerPointMLtTLightRigFlood = '\000\000\000\000',
    PowerPointMLtTLightRigContrasting = '\001\000\000\000',
    PowerPointMLtTLightRigMorning = '\000\000\000\000',
    PowerPointMLtTLightRigSunrise = '\001\000\000\000',
    PowerPointMLtTLightRigSunset = '\000\000\000\000',
    PowerPointMLtTLightRigChilly = '\001\000\000\000',
    PowerPointMLtTLightRigFreezing = '\000\000\000\000',
    PowerPointMLtTLightRigFlat = '\001\000\000\000',
    PowerPointMLtTLightRigTwoPoint = '\000\000\000\000',
    PowerPointMLtTLightRigGlow = '\001\000\000\000',
    PowerPointMLtTLightRigBrightRoom = '\000\000\000\000'
} PowerPointMLtT;

typedef enum {
    PowerPointMBlTBevelTypeUnset = '\000\000\000\000',
    PowerPointMBlTBevelNone = '\001\000\000\000',
    PowerPointMBlTBevelRelaxedInset = '\000\000\000\000',
    PowerPointMBlTBevelCircle = '\001\000\000\000',
    PowerPointMBlTBevelSlope = '\000\000\000\000',
    PowerPointMBlTBevelCross = '\001\000\000\000',
    PowerPointMBlTBevelAngle = '\000\000\000\000',
    PowerPointMBlTBevelSoftRound = '\001\000\000\000',
    PowerPointMBlTBevelConvex = '\000\000\000\000',
    PowerPointMBlTBevelCoolSlant = '\001\000\000\000',
    PowerPointMBlTBevelDivot = '\000\000\000\000',
    PowerPointMBlTBevelRiblet = '\001\000\000\000',
    PowerPointMBlTBevelHardEdge = '\000\000\000\000',
    PowerPointMBlTBevelArtDeco = '\001\000\000\000'
} PowerPointMBlT;

typedef enum {
    PowerPointMSStShadowStyleUnset = '\001\000\000\000',
    PowerPointMSStShadowStyleInner = '\000\000\000\000',
    PowerPointMSStShadowStyleOuter = '\001\000\000\000'
} PowerPointMSSt;

typedef enum {
    PowerPointPpgAParagraphAlignmentUnset = '\001\000\000\000',
    PowerPointPpgAParagraphAlignLeft = '\000\000\000\000',
    PowerPointPpgAParagraphAlignCenter = '\001\000\000\000',
    PowerPointPpgAParagraphAlignRight = '\000\000\000\000',
    PowerPointPpgAParagraphAlignJustify = '\001\000\000\000',
    PowerPointPpgAParagraphAlignDistribute = '\000\000\000\000',
    PowerPointPpgAParagraphAlignThai = '\001\000\000\000',
    PowerPointPpgAParagraphAlignJustifyLow = '\000\000\000\000'
} PowerPointPpgA;

typedef enum {
    PowerPointMTStStrikeUnset = '\000\000\000\000',
    PowerPointMTStNoStrike = '\001\000\000\000',
    PowerPointMTStSingleStrike = '\000\000\000\000',
    PowerPointMTStDoubleStrike = '\001\000\000\000'
} PowerPointMTSt;

typedef enum {
    PowerPointMTCaCapsUnset = '\001\000\000\000',
    PowerPointMTCaNoCaps = '\000\000\000\000',
    PowerPointMTCaSmallCaps = '\001\000\000\000',
    PowerPointMTCaAllCaps = '\000\000\000\000'
} PowerPointMTCa;

typedef enum {
    PowerPointMTUlUnderlineUnset = '\003\356\377\376',
    PowerPointMTUlNoUnderline = '\000\000\003\357',
    PowerPointMTUlUnderlineWordsOnly = '\000\000\003\357',
    PowerPointMTUlUnderlineSingleLine = '\000\000\003\357',
    PowerPointMTUlUnderlineDoubleLine = '\000\000\003\357',
    PowerPointMTUlUnderlineHeavyLine = '\000\000\003\357',
    PowerPointMTUlUnderlineDottedLine = '\000\000\003\357',
    PowerPointMTUlUnderlineHeavyDottedLine = '\000\000\003\357',
    PowerPointMTUlUnderlineDashLine = '\000\000\003\357',
    PowerPointMTUlUnderlineHeavyDashLine = '\000\000\003\357',
    PowerPointMTUlUnderlineLongDashLine = '\000\000\003\357',
    PowerPointMTUlUnderlineHeavyLongDashLine = '\000\000\003\357',
    PowerPointMTUlUnderlineDotDashLine = '\000\000\003\357',
    PowerPointMTUlUnderlineHeavyDotDashLine = '\000\000\003\357',
    PowerPointMTUlUnderlineDotDotDashLine = '\000\000\003\357',
    PowerPointMTUlUnderlineHeavyDotDotDashLine = '\000\000\003\357',
    PowerPointMTUlUnderlineWavyLine = '\000\000\003\357',
    PowerPointMTUlUnderlineHeavyWavyLine = '\000\000\003\357',
    PowerPointMTUlUnderlineWavyDoubleLine = '\000\000\003\357'
} PowerPointMTUl;

typedef enum {
    PowerPointMTTATabUnset = '\000\000\000\000',
    PowerPointMTTALeftTab = '\001\000\000\000',
    PowerPointMTTACenterTab = '\000\000\000\000',
    PowerPointMTTARightTab = '\001\000\000\000',
    PowerPointMTTADecimalTab = '\000\000\000\000'
} PowerPointMTTA;

typedef enum {
    PowerPointMTCWCharacterWrapUnset = '\000\000\000\000',
    PowerPointMTCWNoCharacterWrap = '\001\000\000\000',
    PowerPointMTCWStandardCharacterWrap = '\000\000\000\000',
    PowerPointMTCWStrictCharacterWrap = '\001\000\000\000',
    PowerPointMTCWCustomCharacterWrap = '\000\000\000\000'
} PowerPointMTCW;

typedef enum {
    PowerPointMTFAFontAlignmentUnset = '\000\000\000\000',
    PowerPointMTFAAutomaticAlignment = '\001\000\000\000',
    PowerPointMTFATopAlignment = '\000\000\000\000',
    PowerPointMTFACenterAlignment = '\001\000\000\000',
    PowerPointMTFABaselineAlignment = '\000\000\000\000',
    PowerPointMTFABottomAlignment = '\001\000\000\000'
} PowerPointMTFA;

typedef enum {
    PowerPointPAtSAutoSizeUnset = '\001\000\000\000',
    PowerPointPAtSAutoSizeNone = '\000\000\000\000',
    PowerPointPAtSShapeToFitText = '\001\000\000\000',
    PowerPointPAtSTextToFitShape = '\000\000\000\000'
} PowerPointPAtS;

typedef enum {
    PowerPointMPFoPathTypeUnset = '\000\000\000\000',
    PowerPointMPFoNoPathType = '\001\000\000\000',
    PowerPointMPFoPathType1 = '\000\000\000\000',
    PowerPointMPFoPathType2 = '\001\000\000\000',
    PowerPointMPFoPathType3 = '\000\000\000\000',
    PowerPointMPFoPathType4 = '\001\000\000\000'
} PowerPointMPFo;

typedef enum {
    PowerPointMWFoWarpFormatUnset = '\001\000\000\000',
    PowerPointMWFoWarpFormat1 = '\000\000\000\000',
    PowerPointMWFoWarpFormat2 = '\001\000\000\000',
    PowerPointMWFoWarpFormat3 = '\000\000\000\000',
    PowerPointMWFoWarpFormat4 = '\001\000\000\000',
    PowerPointMWFoWarpFormat5 = '\000\000\000\000',
    PowerPointMWFoWarpFormat6 = '\001\000\000\000',
    PowerPointMWFoWarpFormat7 = '\000\000\000\000',
    PowerPointMWFoWarpFormat8 = '\001\000\000\000',
    PowerPointMWFoWarpFormat9 = '\000\000\000\000',
    PowerPointMWFoWarpFormat10 = '\001\000\000\000',
    PowerPointMWFoWarpFormat11 = '\000\000\000\000',
    PowerPointMWFoWarpFormat12 = '\001\000\000\000',
    PowerPointMWFoWarpFormat13 = '\000\000\000\000',
    PowerPointMWFoWarpFormat14 = '\001\000\000\000',
    PowerPointMWFoWarpFormat15 = '\000\000\000\000',
    PowerPointMWFoWarpFormat16 = '\001\000\000\000',
    PowerPointMWFoWarpFormat17 = '\000\000\000\000',
    PowerPointMWFoWarpFormat18 = '\001\000\000\000',
    PowerPointMWFoWarpFormat19 = '\000\000\000\000',
    PowerPointMWFoWarpFormat20 = '\001\000\000\000',
    PowerPointMWFoWarpFormat21 = '\000\000\000\000',
    PowerPointMWFoWarpFormat22 = '\001\000\000\000',
    PowerPointMWFoWarpFormat23 = '\000\000\000\000',
    PowerPointMWFoWarpFormat24 = '\001\000\000\000',
    PowerPointMWFoWarpFormat25 = '\000\000\000\000',
    PowerPointMWFoWarpFormat26 = '\001\000\000\000',
    PowerPointMWFoWarpFormat27 = '\000\000\000\000',
    PowerPointMWFoWarpFormat28 = '\001\000\000\000',
    PowerPointMWFoWarpFormat29 = '\000\000\000\000',
    PowerPointMWFoWarpFormat30 = '\001\000\000\000',
    PowerPointMWFoWarpFormat31 = '\000\000\000\000',
    PowerPointMWFoWarpFormat32 = '\001\000\000\000',
    PowerPointMWFoWarpFormat33 = '\000\000\000\000',
    PowerPointMWFoWarpFormat34 = '\001\000\000\000',
    PowerPointMWFoWarpFormat35 = '\000\000\000\000',
    PowerPointMWFoWarpFormat36 = '\001\000\000\000'
} PowerPointMWFo;

typedef enum {
    PowerPointPCgCCaseSentence = '\001\000\000\000',
    PowerPointPCgCCaseLower = '\000\000\000\000',
    PowerPointPCgCCaseUpper = '\001\000\000\000',
    PowerPointPCgCCaseTitle = '\000\000\000\000',
    PowerPointPCgCCaseToggle = '\001\000\000\000'
} PowerPointPCgC;

typedef enum {
    PowerPointMDTFDateTimeFormatUnset = '\001\000\000\000',
    PowerPointMDTFDateTimeFormatMdyy = '\000\000\000\000',
    PowerPointMDTFDateTimeFormatDdddMMMMddyyyy = '\001\000\000\000',
    PowerPointMDTFDateTimeFormatDMMMMyyyy = '\000\000\000\000',
    PowerPointMDTFDateTimeFormatMMMMdyyyy = '\001\000\000\000',
    PowerPointMDTFDateTimeFormatDMMMyy = '\000\000\000\000',
    PowerPointMDTFDateTimeFormatMMMMyy = '\001\000\000\000',
    PowerPointMDTFDateTimeFormatMMyy = '\000\000\000\000',
    PowerPointMDTFDateTimeFormatMMddyyHmm = '\001\000\000\000',
    PowerPointMDTFDateTimeFormatMMddyyhmmAMPM = '\000\000\000\000',
    PowerPointMDTFDateTimeFormatHmm = '\001\000\000\000',
    PowerPointMDTFDateTimeFormatHmmss = '\000\000\000\000',
    PowerPointMDTFDateTimeFormatHmmAMPM = '\001\000\000\000',
    PowerPointMDTFDateTimeFormatHmmssAMPM = '\000\000\000\000',
    PowerPointMDTFDateTimeFormatFigureOut = '\001\000\000\000'
} PowerPointMDTF;

typedef enum {
    PowerPointMSETSoftEdgeUnset = '\001\000\000\000',
    PowerPointMSETNoSoftEdge = '\000\000\000\000',
    PowerPointMSETSoftEdgeType1 = '\001\000\000\000',
    PowerPointMSETSoftEdgeType2 = '\000\000\000\000',
    PowerPointMSETSoftEdgeType3 = '\001\000\000\000',
    PowerPointMSETSoftEdgeType4 = '\000\000\000\000',
    PowerPointMSETSoftEdgeType5 = '\001\000\000\000',
    PowerPointMSETSoftEdgeType6 = '\000\000\000\000'
} PowerPointMSET;

typedef enum {
    PowerPointMCSIFirstDarkSchemeColor = '\000\000\000\000',
    PowerPointMCSIFirstLightSchemeColor = '\001\000\000\000',
    PowerPointMCSISecondDarkSchemeColor = '\000\000\000\000',
    PowerPointMCSISecondLightSchemeColor = '\001\000\000\000',
    PowerPointMCSIFirstAccentSchemeColor = '\000\000\000\000',
    PowerPointMCSISecondAccentSchemeColor = '\001\000\000\000',
    PowerPointMCSIThirdAccentSchemeColor = '\000\000\000\000',
    PowerPointMCSIFourthAccentSchemeColor = '\001\000\000\000',
    PowerPointMCSIFifthAccentSchemeColor = '\000\000\000\000',
    PowerPointMCSISixthAccentSchemeColor = '\001\000\000\000',
    PowerPointMCSIHyperlinkSchemeColor = '\000\000\000\000',
    PowerPointMCSIFollowedHyperlinkSchemeColor = '\001\000\000\000'
} PowerPointMCSI;

typedef enum {
    PowerPointMCoIThemeColorUnset = '\001\000\000\000',
    PowerPointMCoINoThemeColor = '\000\000\000\000',
    PowerPointMCoIFirstDarkThemeColor = '\001\000\000\000',
    PowerPointMCoIFirstLightThemeColor = '\000\000\000\000',
    PowerPointMCoISecondDarkThemeColor = '\001\000\000\000',
    PowerPointMCoISecondLightThemeColor = '\000\000\000\000',
    PowerPointMCoIFirstAccentThemeColor = '\001\000\000\000',
    PowerPointMCoISecondAccentThemeColor = '\000\000\000\000',
    PowerPointMCoIThirdAccentThemeColor = '\001\000\000\000',
    PowerPointMCoIFourthAccentThemeColor = '\000\000\000\000',
    PowerPointMCoIFifthAccentThemeColor = '\001\000\000\000',
    PowerPointMCoISixthAccentThemeColor = '\000\000\000\000',
    PowerPointMCoIHyperlinkThemeColor = '\001\000\000\000',
    PowerPointMCoIFollowedHyperlinkThemeColor = '\000\000\000\000',
    PowerPointMCoIFirstTextThemeColor = '\001\000\000\000',
    PowerPointMCoIFirstBackgroundThemeColor = '\000\000\000\000',
    PowerPointMCoISecondTextThemeColor = '\001\000\000\000',
    PowerPointMCoISecondBackgroundThemeColor = '\000\000\000\000'
} PowerPointMCoI;

typedef enum {
    PowerPointMFLIThemeFontLatin = '\000\000\000\000',
    PowerPointMFLIThemeFontComplexScript = '\001\000\000\000',
    PowerPointMFLIThemeFontHighAnsi = '\000\000\000\000',
    PowerPointMFLIThemeFontEastAsian = '\001\000\000\000'
} PowerPointMFLI;

typedef enum {
    PowerPointMSSIShapeStyleUnset = '\001\000\000\000',
    PowerPointMSSIShapeNotAPreset = '\000\000\000\000',
    PowerPointMSSIShapePreset1 = '\001\000\000\000',
    PowerPointMSSIShapePreset2 = '\000\000\000\000',
    PowerPointMSSIShapePreset3 = '\001\000\000\000',
    PowerPointMSSIShapePreset4 = '\000\000\000\000',
    PowerPointMSSIShapePreset5 = '\001\000\000\000',
    PowerPointMSSIShapePreset6 = '\000\000\000\000',
    PowerPointMSSIShapePreset7 = '\001\000\000\000',
    PowerPointMSSIShapePreset8 = '\000\000\000\000',
    PowerPointMSSIShapePreset9 = '\001\000\000\000',
    PowerPointMSSIShapePreset10 = '\000\000\000\000',
    PowerPointMSSIShapePreset11 = '\001\000\000\000',
    PowerPointMSSIShapePreset12 = '\000\000\000\000',
    PowerPointMSSIShapePreset13 = '\001\000\000\000',
    PowerPointMSSIShapePreset14 = '\000\000\000\000',
    PowerPointMSSIShapePreset15 = '\001\000\000\000',
    PowerPointMSSIShapePreset16 = '\000\000\000\000',
    PowerPointMSSIShapePreset17 = '\001\000\000\000',
    PowerPointMSSIShapePreset18 = '\000\000\000\000',
    PowerPointMSSIShapePreset19 = '\001\000\000\000',
    PowerPointMSSIShapePreset20 = '\000\000\000\000',
    PowerPointMSSIShapePreset21 = '\001\000\000\000',
    PowerPointMSSIShapePreset22 = '\000\000\000\000',
    PowerPointMSSIShapePreset23 = '\001\000\000\000',
    PowerPointMSSIShapePreset24 = '\000\000\000\000',
    PowerPointMSSIShapePreset25 = '\001\000\000\000',
    PowerPointMSSIShapePreset26 = '\000\000\000\000',
    PowerPointMSSIShapePreset27 = '\001\000\000\000',
    PowerPointMSSIShapePreset28 = '\000\000\000\000',
    PowerPointMSSIShapePreset29 = '\001\000\000\000',
    PowerPointMSSIShapePreset30 = '\000\000\000\000',
    PowerPointMSSIShapePreset31 = '\001\000\000\000',
    PowerPointMSSIShapePreset32 = '\000\000\000\000',
    PowerPointMSSIShapePreset33 = '\001\000\000\000',
    PowerPointMSSIShapePreset34 = '\000\000\000\000',
    PowerPointMSSIShapePreset35 = '\001\000\000\000',
    PowerPointMSSIShapePreset36 = '\000\000\000\000',
    PowerPointMSSIShapePreset37 = '\001\000\000\000',
    PowerPointMSSIShapePreset38 = '\000\000\000\000',
    PowerPointMSSIShapePreset39 = '\001\000\000\000',
    PowerPointMSSIShapePreset40 = '\000\000\000\000',
    PowerPointMSSIShapePreset41 = '\001\000\000\000',
    PowerPointMSSIShapePreset42 = '\000\000\000\000',
    PowerPointMSSILinePreset1 = '\001\000\000\000',
    PowerPointMSSILinePreset2 = '\000\000\000\000',
    PowerPointMSSILinePreset3 = '\001\000\000\000',
    PowerPointMSSILinePreset4 = '\000\000\000\000',
    PowerPointMSSILinePreset5 = '\001\000\000\000',
    PowerPointMSSILinePreset6 = '\000\000\000\000',
    PowerPointMSSILinePreset7 = '\001\000\000\000',
    PowerPointMSSILinePreset8 = '\000\000\000\000',
    PowerPointMSSILinePreset9 = '\001\000\000\000',
    PowerPointMSSILinePreset10 = '\000\000\000\000',
    PowerPointMSSILinePreset11 = '\001\000\000\000',
    PowerPointMSSILinePreset12 = '\000\000\000\000',
    PowerPointMSSILinePreset13 = '\001\000\000\000',
    PowerPointMSSILinePreset14 = '\000\000\000\000',
    PowerPointMSSILinePreset15 = '\001\000\000\000',
    PowerPointMSSILinePreset16 = '\000\000\000\000',
    PowerPointMSSILinePreset17 = '\001\000\000\000',
    PowerPointMSSILinePreset18 = '\000\000\000\000',
    PowerPointMSSILinePreset19 = '\001\000\000\000',
    PowerPointMSSILinePreset20 = '\000\000\000\000',
    PowerPointMSSILinePreset21 = '\001\000\000\000'
} PowerPointMSSI;

typedef enum {
    PowerPointMBSIBackgroundUnset = '\001\000\000\000',
    PowerPointMBSIBackgroundNotAPreset = '\000\000\000\000',
    PowerPointMBSIBackgroundPreset1 = '\001\000\000\000',
    PowerPointMBSIBackgroundPreset2 = '\000\000\000\000',
    PowerPointMBSIBackgroundPreset3 = '\001\000\000\000',
    PowerPointMBSIBackgroundPreset4 = '\000\000\000\000',
    PowerPointMBSIBackgroundPreset5 = '\001\000\000\000',
    PowerPointMBSIBackgroundPreset6 = '\000\000\000\000',
    PowerPointMBSIBackgroundPreset7 = '\001\000\000\000',
    PowerPointMBSIBackgroundPreset8 = '\000\000\000\000',
    PowerPointMBSIBackgroundPreset9 = '\001\000\000\000',
    PowerPointMBSIBackgroundPreset10 = '\000\000\000\000',
    PowerPointMBSIBackgroundPreset11 = '\001\000\000\000',
    PowerPointMBSIBackgroundPreset12 = '\000\000\000\000'
} PowerPointMBSI;

typedef enum {
    PowerPointPDrTTextDirectionUnset = '\000\000\000\000',
    PowerPointPDrTLeftToRight = '\001\000\000\000',
    PowerPointPDrTRightToLeft = '\000\000\000\000'
} PowerPointPDrT;

typedef enum {
    PowerPointPBtTBulletTypeUnset = '\000\000\000\000',
    PowerPointPBtTBulletTypeNone = '\001\000\000\000',
    PowerPointPBtTBulletTypeUnnumbered = '\000\000\000\000',
    PowerPointPBtTBulletTypeNumbered = '\001\000\000\000',
    PowerPointPBtTPictureBulletType = '\000\000\000\000'
} PowerPointPBtT;

typedef enum {
    PowerPointPBtSNumberedBulletStyleUnset = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleAlphaLowercasePeriod = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleAlphaUppercasePeriod = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleArabicRightParen = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleArabicPeriod = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleRomanLowercaseParenBoth = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleRomanLowercaseParenRight = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleRomanLowercasePeriod = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleRomanUppercasePeriod = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleAlphaLowercaseParenBoth = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleAlphaLowercaseParenRight = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleAlphaUppercaseParenBoth = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleAlphaUppercaseParenRight = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleArabicParenBoth = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleArabicPlain = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleRomanUppercaseParenBoth = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleRomanUppercaseParenRight = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleSimplifiedChinesePlain = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleSimplifiedChinesePeriod = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleCircleNumberPlain = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleCircleNumberWhitePlain = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleCircleNumberBlackPlain = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleTraditionalChinesePlain = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleTraditionalChinesePeriod = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleArabicAlphaDash = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleArabicAbjadDash = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleHebrewAlphaDash = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleKanjiKoreanPlain = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleKanjiKoreanPeriod = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleArabicDBPlain = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleArabicDBPeriod = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleThaiAlphaPeriod = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleThaiAlphaParenRight = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleThaiAlphaParenBoth = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleThaiNumberPeriod = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleThaiNumberParenRight = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleThaiParenBoth = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleHindiAlphaPeriod = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleHindiNumberPeriod = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\001\000\000\000',
    PowerPointPBtSNumberedBulletStyleHindiNumberParenRight = '\000\000\000\000',
    PowerPointPBtSNumberedBulletStyleHindiAlpha1Period = '\001\000\000\000'
} PowerPointPBtS;

typedef enum {
    PowerPointPTSpTabstopUnset = '\001\000\000\000',
    PowerPointPTSpTabstopLeft = '\000\000\000\000',
    PowerPointPTSpTabstopCenter = '\001\000\000\000',
    PowerPointPTSpTabstopRight = '\000\000\000\000',
    PowerPointPTSpTabstopDecimal = '\001\000\000\000'
} PowerPointPTSp;

typedef enum {
    PowerPointMRfTReflectionUnset = '\003\351\377\376',
    PowerPointMRfTReflectionTypeNone = '\000\000\003\352',
    PowerPointMRfTReflectionType1 = '\000\000\003\352',
    PowerPointMRfTReflectionType2 = '\000\000\003\352',
    PowerPointMRfTReflectionType3 = '\000\000\003\352',
    PowerPointMRfTReflectionType4 = '\000\000\003\352',
    PowerPointMRfTReflectionType5 = '\000\000\003\352',
    PowerPointMRfTReflectionType6 = '\000\000\003\352',
    PowerPointMRfTReflectionType7 = '\000\000\003\352',
    PowerPointMRfTReflectionType8 = '\000\000\003\352',
    PowerPointMRfTReflectionType9 = '\000\000\003\352'
} PowerPointMRfT;

typedef enum {
    PowerPointMTtATextureUnset = '\003\352\377\376',
    PowerPointMTtATextureTopLeft = '\000\000\003\353',
    PowerPointMTtATextureTop = '\000\000\003\353',
    PowerPointMTtATextureTopRight = '\000\000\003\353',
    PowerPointMTtATextureLeft = '\000\000\003\353',
    PowerPointMTtATextureCenter = '\000\000\003\353',
    PowerPointMTtATextureRight = '\000\000\003\353',
    PowerPointMTtATextureBottomLeft = '\000\000\003\353',
    PowerPointMTtATextureBotton = '\000\000\003\353',
    PowerPointMTtATextureBottomRight = '\000\000\003\353'
} PowerPointMTtA;

typedef enum {
    PowerPointPBlABaselineAlignmentUnset = '\003\353\377\376',
    PowerPointPBlABaselineAlignBaseline = '\000\000\003\354',
    PowerPointPBlABaselineAlignTop = '\000\000\003\354',
    PowerPointPBlABaselineAlignCenter = '\000\000\003\354',
    PowerPointPBlABaselineAlignEastAsian50 = '\000\000\003\354',
    PowerPointPBlABaselineAlignAutomatic = '\000\000\003\354'
} PowerPointPBlA;

typedef enum {
    PowerPointMCbFClipboardFormatUnset = '\003\354\377\376',
    PowerPointMCbFNativeClipboardFormat = '\000\000\003\355',
    PowerPointMCbFHTMlClipboardFormat = '\000\000\003\355',
    PowerPointMCbFRTFClipboardFormat = '\000\000\003\355',
    PowerPointMCbFPlainTextClipboardFormat = '\000\000\003\355'
} PowerPointMCbF;

typedef enum {
    PowerPointMTiPInsertBefore = '\000\000\003\356',
    PowerPointMTiPInsertAfter = '\000\000\003\356'
} PowerPointMTiP;

typedef enum {
    PowerPointMPiTSaveAsDefault = '\003\362\377\376',
    PowerPointMPiTSaveAsPNGFile = '\000\000\003\363',
    PowerPointMPiTSaveAsBMPFile = '\000\000\003\363',
    PowerPointMPiTSaveAsGIFFile = '\000\000\003\363',
    PowerPointMPiTSaveAsJPGFile = '\000\000\003\363',
    PowerPointMPiTSaveAsPDFFile = '\000\000\003\363'
} PowerPointMPiT;

typedef enum {
    PowerPointMAlCAlignLefts = '\001\000\000\000',
    PowerPointMAlCAlignCenters = '\000\000\000\000',
    PowerPointMAlCAlignRights = '\001\000\000\000',
    PowerPointMAlCAlignTops = '\000\000\000\000',
    PowerPointMAlCAlignMiddles = '\001\000\000\000',
    PowerPointMAlCAlignBottoms = '\000\000\000\000'
} PowerPointMAlC;

typedef enum {
    PowerPointMDsCDistributeHorizontally = '\000\000\000\000',
    PowerPointMDsCDistributeVertically = '\001\000\000\000'
} PowerPointMDsC;

typedef enum {
    PowerPointMOrTOrientationUnset = '\001\000\000\000',
    PowerPointMOrTHorizontalOrientation = '\000\000\000\000',
    PowerPointMOrTVerticalOrientation = '\001\000\000\000'
} PowerPointMOrT;

typedef enum {
    PowerPointMZoCBringShapeToFront = '\001\000\000\000',
    PowerPointMZoCSendShapeToBack = '\000\000\000\000',
    PowerPointMZoCBringShapeForward = '\001\000\000\000',
    PowerPointMZoCSendShapeBackward = '\000\000\000\000',
    PowerPointMZoCBringShapeInFrontOfText = '\001\000\000\000',
    PowerPointMZoCSendShapeBehindText = '\000\000\000\000'
} PowerPointMZoC;

typedef enum {
    PowerPointMSgTLine = '\000\000\000\000',
    PowerPointMSgTCurve = '\001\000\000\000'
} PowerPointMSgT;

typedef enum {
    PowerPointMEdTAuto = '\001\000\000\000',
    PowerPointMEdTCorner = '\000\000\000\000',
    PowerPointMEdTSmooth = '\001\000\000\000',
    PowerPointMEdTSymmetric = '\000\000\000\000'
} PowerPointMEdT;

typedef enum {
    PowerPointMFlCFlipHorizontal = '\000\000\000\000',
    PowerPointMFlCFlipVertical = '\001\000\000\000'
} PowerPointMFlC;

typedef enum {
    PowerPointMTriTrue = '\001\000\000\000',
    PowerPointMTriFalse = '\000\000\000\000',
    PowerPointMTriCTrue = '\001\000\000\000',
    PowerPointMTriToggle = '\000\000\000\000',
    PowerPointMTriTriStateUnset = '\001\000\000\000'
} PowerPointMTri;

typedef enum {
    PowerPointMBWBlackAndWhiteUnset = '\001\000\000\000',
    PowerPointMBWBlackAndWhiteModeAutomatic = '\000\000\000\000',
    PowerPointMBWBlackAndWhiteModeGrayScale = '\001\000\000\000',
    PowerPointMBWBlackAndWhiteModeLightGrayScale = '\000\000\000\000',
    PowerPointMBWBlackAndWhiteModeInverseGrayScale = '\001\000\000\000',
    PowerPointMBWBlackAndWhiteModeGrayOutline = '\000\000\000\000',
    PowerPointMBWBlackAndWhiteModeBlackTextAndLine = '\001\000\000\000',
    PowerPointMBWBlackAndWhiteModeHighContrast = '\000\000\000\000',
    PowerPointMBWBlackAndWhiteModeBlack = '\001\000\000\000',
    PowerPointMBWBlackAndWhiteModeWhite = '\000\000\000\000',
    PowerPointMBWBlackAndWhiteModeDontShow = '\001\000\000\000'
} PowerPointMBW;

typedef enum {
    PowerPointMBPSBarLeft = '\001\000\000\000',
    PowerPointMBPSBarTop = '\000\000\000\000',
    PowerPointMBPSBarRight = '\001\000\000\000',
    PowerPointMBPSBarBottom = '\000\000\000\000',
    PowerPointMBPSBarFloating = '\001\000\000\000',
    PowerPointMBPSBarPopUp = '\000\000\000\000',
    PowerPointMBPSBarMenu = '\001\000\000\000'
} PowerPointMBPS;

typedef enum {
    PowerPointMBPtNoProtection = '\001\000\000\000',
    PowerPointMBPtNoCustomize = '\000\000\000\000',
    PowerPointMBPtNoResize = '\001\000\000\000',
    PowerPointMBPtNoMove = '\000\000\000\000',
    PowerPointMBPtNoChangeVisible = '\001\000\000\000',
    PowerPointMBPtNoChangeDock = '\000\000\000\000',
    PowerPointMBPtNoVerticalDock = '\001\000\000\000',
    PowerPointMBPtNoHorizontalDock = '\000\000\000\000'
} PowerPointMBPt;

typedef enum {
    PowerPointMBTYNormalCommandBar = '\000\000\000\000',
    PowerPointMBTYMenubarCommandBar = '\001\000\000\000',
    PowerPointMBTYPopupCommandBar = '\000\000\000\000'
} PowerPointMBTY;

typedef enum {
    PowerPointMCLTControlCustom = '\000\000\000\000',
    PowerPointMCLTControlButton = '\001\000\000\000',
    PowerPointMCLTControlEdit = '\000\000\000\000',
    PowerPointMCLTControlDropDown = '\001\000\000\000',
    PowerPointMCLTControlCombobox = '\000\000\000\000',
    PowerPointMCLTButtonDropDown = '\001\000\000\000',
    PowerPointMCLTSplitDropDown = '\000\000\000\000',
    PowerPointMCLTOCXDropDown = '\001\000\000\000',
    PowerPointMCLTGenericDropDown = '\000\000\000\000',
    PowerPointMCLTGraphicDropDown = '\001\000\000\000',
    PowerPointMCLTControlPopup = '\000\000\000\000',
    PowerPointMCLTGraphicPopup = '\001\000\000\000',
    PowerPointMCLTButtonPopup = '\000\000\000\000',
    PowerPointMCLTSplitButtonPopup = '\001\000\000\000',
    PowerPointMCLTSplitButtonMRUPopup = '\000\000\000\000',
    PowerPointMCLTControlLabel = '\001\000\000\000',
    PowerPointMCLTExpandingGrid = '\000\000\000\000',
    PowerPointMCLTSplitExpandingGrid = '\001\000\000\000',
    PowerPointMCLTControlGrid = '\000\000\000\000',
    PowerPointMCLTControlGauge = '\001\000\000\000',
    PowerPointMCLTGraphicCombobox = '\000\000\000\000',
    PowerPointMCLTControlPane = '\001\000\000\000',
    PowerPointMCLTActiveX = '\000\000\000\000',
    PowerPointMCLTControlGroup = '\001\000\000\000',
    PowerPointMCLTControlTab = '\000\000\000\000',
    PowerPointMCLTControlSpinner = '\001\000\000\000'
} PowerPointMCLT;

typedef enum {
    PowerPointMBnsButtonStateUp = '\001\000\000\000',
    PowerPointMBnsButtonStateDown = '\000\000\000\000',
    PowerPointMBnsButtonStateUnset = '\001\000\000\000'
} PowerPointMBns;

typedef enum {
    PowerPointMcOuNeither = '\001\000\000\000',
    PowerPointMcOuServer = '\000\000\000\000',
    PowerPointMcOuClient = '\001\000\000\000',
    PowerPointMcOuBoth = '\000\000\000\000'
} PowerPointMcOu;

typedef enum {
    PowerPointMBTsButtonAutomatic = '\000\000\000\000',
    PowerPointMBTsButtonIcon = '\001\000\000\000',
    PowerPointMBTsButtonCaption = '\000\000\000\000',
    PowerPointMBTsButtonIconAndCaption = '\001\000\000\000'
} PowerPointMBTs;

typedef enum {
    PowerPointMXcbComboboxStyleNormal = '\001\000\000\000',
    PowerPointMXcbComboboxStyleLabel = '\000\000\000\000'
} PowerPointMXcb;

typedef enum {
    PowerPointMMNANone = '\000\000\000\000',
    PowerPointMMNARandom = '\001\000\000\000',
    PowerPointMMNAUnfold = '\000\000\000\000',
    PowerPointMMNASlide = '\001\000\000\000'
} PowerPointMMNA;

typedef enum {
    PowerPointMHlTHyperlinkTypeTextRange = '\001\000\000\000',
    PowerPointMHlTHyperlinkTypeShape = '\000\000\000\000',
    PowerPointMHlTHyperlinkTypeInlineShape = '\001\000\000\000'
} PowerPointMHlT;

typedef enum {
    PowerPointMXiMAppendString = '\001\000\000\000',
    PowerPointMXiMPostString = '\000\000\000\000'
} PowerPointMXiM;

typedef enum {
    PowerPointMANTIdle = '\000\000\000\000',
    PowerPointMANTGreeting = '\001\000\000\000',
    PowerPointMANTGoodbye = '\000\000\000\000',
    PowerPointMANTBeginSpeaking = '\001\000\000\000',
    PowerPointMANTCharacterSuccessMajor = '\000\000\000\000',
    PowerPointMANTGetAttentionMajor = '\001\000\000\000',
    PowerPointMANTGetAttentionMinor = '\000\000\000\000',
    PowerPointMANTSearching = '\001\000\000\000',
    PowerPointMANTPrinting = '\000\000\000\000',
    PowerPointMANTGestureRight = '\001\000\000\000',
    PowerPointMANTWritingNotingSomething = '\000\000\000\000',
    PowerPointMANTWorkingAtSomething = '\001\000\000\000',
    PowerPointMANTThinking = '\000\000\000\000',
    PowerPointMANTSendingMail = '\001\000\000\000',
    PowerPointMANTListensToComputer = '\000\000\000\000',
    PowerPointMANTDisappear = '\001\000\000\000',
    PowerPointMANTAppear = '\000\000\000\000',
    PowerPointMANTGetArtsy = '\001\000\000\000',
    PowerPointMANTGetTechy = '\000\000\000\000',
    PowerPointMANTGetWizardy = '\001\000\000\000',
    PowerPointMANTCheckingSomething = '\000\000\000\000',
    PowerPointMANTLookDown = '\001\000\000\000',
    PowerPointMANTLookDownLeft = '\000\000\000\000',
    PowerPointMANTLookDownRight = '\001\000\000\000',
    PowerPointMANTLookLeft = '\000\000\000\000',
    PowerPointMANTLookRight = '\001\000\000\000',
    PowerPointMANTLookUp = '\000\000\000\000',
    PowerPointMANTLookUpLeft = '\001\000\000\000',
    PowerPointMANTLookUpRight = '\000\000\000\000',
    PowerPointMANTSaving = '\001\000\000\000',
    PowerPointMANTGestureDown = '\000\000\000\000',
    PowerPointMANTGestureLeft = '\001\000\000\000',
    PowerPointMANTGestureUp = '\000\000\000\000',
    PowerPointMANTEmptyTrash = '\001\000\000\000'
} PowerPointMANT;

typedef enum {
    PowerPointMBStButtonNone = '\001\000\000\000',
    PowerPointMBStButtonOk = '\000\000\000\000',
    PowerPointMBStButtonCancel = '\001\000\000\000',
    PowerPointMBStButtonsOkCancel = '\000\000\000\000',
    PowerPointMBStButtonsYesNo = '\001\000\000\000',
    PowerPointMBStButtonsYesNoCancel = '\000\000\000\000',
    PowerPointMBStButtonsBackClose = '\001\000\000\000',
    PowerPointMBStButtonsNextClose = '\000\000\000\000',
    PowerPointMBStButtonsBackNextClose = '\001\000\000\000',
    PowerPointMBStButtonsRetryCancel = '\000\000\000\000',
    PowerPointMBStButtonsAbortRetryIgnore = '\001\000\000\000',
    PowerPointMBStButtonsSearchClose = '\000\000\000\000',
    PowerPointMBStButtonsBackNextSnooze = '\001\000\000\000',
    PowerPointMBStButtonsTipsOptionsClose = '\000\000\000\000',
    PowerPointMBStButtonsYesAllNoCancel = '\001\000\000\000'
} PowerPointMBSt;

typedef enum {
    PowerPointMIctIconNone = '\001\000\000\000',
    PowerPointMIctIconApplication = '\000\000\000\000',
    PowerPointMIctIconAlert = '\001\000\000\000',
    PowerPointMIctIconTip = '\000\000\000\000',
    PowerPointMIctIconAlertCritical = '\001\000\000\000',
    PowerPointMIctIconAlertWarning = '\000\000\000\000',
    PowerPointMIctIconAlertInfo = '\001\000\000\000'
} PowerPointMIct;

typedef enum {
    PowerPointMWAtInactive = '\001\000\000\000',
    PowerPointMWAtActive = '\000\000\000\000',
    PowerPointMWAtSuspend = '\001\000\000\000',
    PowerPointMWAtResume = '\000\000\000\000'
} PowerPointMWAt;

typedef enum {
    PowerPointMeDPPropertyTypeNumber = '\000\000\000\000',
    PowerPointMeDPPropertyTypeBoolean = '\001\000\000\000',
    PowerPointMeDPPropertyTypeDate = '\000\000\000\000',
    PowerPointMeDPPropertyTypeString = '\001\000\000\000',
    PowerPointMeDPPropertyTypeFloat = '\000\000\000\000'
} PowerPointMeDP;

typedef enum {
    PowerPointMAScMsoAutomationSecurityLow = '\000\000\000\000',
    PowerPointMAScMsoAutomationSecurityByUI = '\001\000\000\000',
    PowerPointMAScMsoAutomationSecurityForceDisable = '\000\000\000\000'
} PowerPointMASc;

typedef enum {
    PowerPointMSszResolution544x376 = '\000\000\000\000',
    PowerPointMSszResolution640x480 = '\001\000\000\000',
    PowerPointMSszResolution720x512 = '\000\000\000\000',
    PowerPointMSszResolution800x600 = '\001\000\000\000',
    PowerPointMSszResolution1024x768 = '\000\000\000\000',
    PowerPointMSszResolution1152x882 = '\001\000\000\000',
    PowerPointMSszResolution1152x900 = '\000\000\000\000',
    PowerPointMSszResolution1280x1024 = '\001\000\000\000',
    PowerPointMSszResolution1600x1200 = '\000\000\000\000',
    PowerPointMSszResolution1800x1440 = '\001\000\000\000',
    PowerPointMSszResolution1920x1200 = '\000\000\000\000'
} PowerPointMSsz;

typedef enum {
    PowerPointMChSArabicCharacterSet = '\000\000\000\000',
    PowerPointMChSCyrillicCharacterSet = '\001\000\000\000',
    PowerPointMChSEnglishCharacterSet = '\000\000\000\000',
    PowerPointMChSGreekCharacterSet = '\001\000\000\000',
    PowerPointMChSHebrewCharacterSet = '\000\000\000\000',
    PowerPointMChSJapaneseCharacterSet = '\001\000\000\000',
    PowerPointMChSKoreanCharacterSet = '\000\000\000\000',
    PowerPointMChSMultilingualUnicodeCharacterSet = '\001\000\000\000',
    PowerPointMChSSimplifiedChineseCharacterSet = '\000\000\000\000',
    PowerPointMChSThaiCharacterSet = '\001\000\000\000',
    PowerPointMChSTraditionalChineseCharacterSet = '\000\000\000\000',
    PowerPointMChSVietnameseCharacterSet = '\001\000\000\000'
} PowerPointMChS;

typedef enum {
    PowerPointMtEnEncodingThai = '\001\000\000\000',
    PowerPointMtEnEncodingJapaneseShiftJIS = '\000\000\000\000',
    PowerPointMtEnEncodingSimplifiedChinese = '\001\000\000\000',
    PowerPointMtEnEncodingKorean = '\000\000\000\000',
    PowerPointMtEnEncodingBig5TraditionalChinese = '\001\000\000\000',
    PowerPointMtEnEncodingLittleEndian = '\000\000\000\000',
    PowerPointMtEnEncodingBigEndian = '\001\000\000\000',
    PowerPointMtEnEncodingCentralEuropean = '\000\000\000\000',
    PowerPointMtEnEncodingCyrillic = '\001\000\000\000',
    PowerPointMtEnEncodingWestern = '\000\000\000\000',
    PowerPointMtEnEncodingGreek = '\001\000\000\000',
    PowerPointMtEnEncodingTurkish = '\000\000\000\000',
    PowerPointMtEnEncodingHebrew = '\001\000\000\000',
    PowerPointMtEnEncodingArabic = '\000\000\000\000',
    PowerPointMtEnEncodingBaltic = '\001\000\000\000',
    PowerPointMtEnEncodingVietnamese = '\000\000\000\000',
    PowerPointMtEnEncodingISO88591Latin1 = '\001\000\000\000',
    PowerPointMtEnEncodingISO88592CentralEurope = '\000\000\000\000',
    PowerPointMtEnEncodingISO88593Latin3 = '\001\000\000\000',
    PowerPointMtEnEncodingISO88594Baltic = '\000\000\000\000',
    PowerPointMtEnEncodingISO88595Cyrillic = '\001\000\000\000',
    PowerPointMtEnEncodingISO88596Arabic = '\000\000\000\000',
    PowerPointMtEnEncodingISO88597Greek = '\001\000\000\000',
    PowerPointMtEnEncodingISO88598Hebrew = '\000\000\000\000',
    PowerPointMtEnEncodingISO88599Turkish = '\001\000\000\000',
    PowerPointMtEnEncodingISO885915Latin9 = '\000\000\000\000',
    PowerPointMtEnEncodingISO2022JapaneseNoHalfWidthKatakana = '\001\000\000\000',
    PowerPointMtEnEncodingISO2022JapaneseJISX02021984 = '\000\000\000\000',
    PowerPointMtEnEncodingISO2022JapaneseJISX02011989 = '\001\000\000\000',
    PowerPointMtEnEncodingISO2022KR = '\000\000\000\000',
    PowerPointMtEnEncodingISO2022CNTraditionalChinese = '\001\000\000\000',
    PowerPointMtEnEncodingISO2022CNSimplifiedChinese = '\000\000\000\000',
    PowerPointMtEnEncodingMacRoman = '\001\000\000\000',
    PowerPointMtEnEncodingMacJapanese = '\000\000\000\000',
    PowerPointMtEnEncodingMacTraditionalChinese = '\001\000\000\000',
    PowerPointMtEnEncodingMacKorean = '\000\000\000\000',
    PowerPointMtEnEncodingMacArabic = '\001\000\000\000',
    PowerPointMtEnEncodingMacHebrew = '\000\000\000\000',
    PowerPointMtEnEncodingMacGreek1 = '\001\000\000\000',
    PowerPointMtEnEncodingMacCyrillic = '\000\000\000\000',
    PowerPointMtEnEncodingMacSimplifiedChineseGB2312 = '\001\000\000\000',
    PowerPointMtEnEncodingMacRomania = '\000\000\000\000',
    PowerPointMtEnEncodingMacUkraine = '\001\000\000\000',
    PowerPointMtEnEncodingMacLatin2 = '\000\000\000\000',
    PowerPointMtEnEncodingMacIcelandic = '\001\000\000\000',
    PowerPointMtEnEncodingMacTurkish = '\000\000\000\000',
    PowerPointMtEnEncodingMacCroatia = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICUSCanada = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICInternational = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICMultilingualROECELatin2 = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICGreekModern = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICTurkishLatin5 = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICGermany = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICDenmarkNorway = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICFinlandSweden = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICItaly = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICLatinAmericaSpain = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICUnitedKingdom = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICJapaneseKatakanaExtended = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICFrance = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICArabic = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICGreek = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICHebrew = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICKoreanExtended = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICThai = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICIcelandic = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICTurkish = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICRussian = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICSerbianBulgarian = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICUSCanadaAndJapanese = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICExtendedAndKorean = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\001\000\000\000',
    PowerPointMtEnEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\000\000\000',
    PowerPointMtEnEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\001\000\000\000',
    PowerPointMtEnEncodingOEMUnitedStates = '\000\000\000\000',
    PowerPointMtEnEncodingOEMGreek = '\001\000\000\000',
    PowerPointMtEnEncodingOEMBaltic = '\000\000\000\000',
    PowerPointMtEnEncodingOEMMultilingualLatinI = '\001\000\000\000',
    PowerPointMtEnEncodingOEMMultilingualLatinII = '\000\000\000\000',
    PowerPointMtEnEncodingOEMCyrillic = '\001\000\000\000',
    PowerPointMtEnEncodingOEMTurkish = '\000\000\000\000',
    PowerPointMtEnEncodingOEMPortuguese = '\001\000\000\000',
    PowerPointMtEnEncodingOEMIcelandic = '\000\000\000\000',
    PowerPointMtEnEncodingOEMHebrew = '\001\000\000\000',
    PowerPointMtEnEncodingOEMCanadianFrench = '\000\000\000\000',
    PowerPointMtEnEncodingOEMArabic = '\001\000\000\000',
    PowerPointMtEnEncodingOEMNordic = '\000\000\000\000',
    PowerPointMtEnEncodingOEMCyrillicII = '\001\000\000\000',
    PowerPointMtEnEncodingOEMModernGreek = '\000\000\000\000',
    PowerPointMtEnEncodingEUCJapanese = '\001\000\000\000',
    PowerPointMtEnEncodingEUCChineseSimplifiedChinese = '\000\000\000\000',
    PowerPointMtEnEncodingEUCKorean = '\001\000\000\000',
    PowerPointMtEnEncodingEUCTaiwaneseTraditionalChinese = '\000\000\000\000',
    PowerPointMtEnEncodingDevanagari = '\001\000\000\000',
    PowerPointMtEnEncodingBengali = '\000\000\000\000',
    PowerPointMtEnEncodingTamil = '\001\000\000\000',
    PowerPointMtEnEncodingTelugu = '\000\000\000\000',
    PowerPointMtEnEncodingAssamese = '\001\000\000\000',
    PowerPointMtEnEncodingOriya = '\000\000\000\000',
    PowerPointMtEnEncodingKannada = '\001\000\000\000',
    PowerPointMtEnEncodingMalayalam = '\000\000\000\000',
    PowerPointMtEnEncodingGujarati = '\001\000\000\000',
    PowerPointMtEnEncodingPunjabi = '\000\000\000\000',
    PowerPointMtEnEncodingArabicASMO = '\001\000\000\000',
    PowerPointMtEnEncodingArabicTransparentASMO = '\000\000\000\000',
    PowerPointMtEnEncodingKoreanJohab = '\001\000\000\000',
    PowerPointMtEnEncodingTaiwanCNS = '\000\000\000\000',
    PowerPointMtEnEncodingTaiwanTCA = '\001\000\000\000',
    PowerPointMtEnEncodingTaiwanEten = '\000\000\000\000',
    PowerPointMtEnEncodingTaiwanIBM5550 = '\001\000\000\000',
    PowerPointMtEnEncodingTaiwanTeletext = '\000\000\000\000',
    PowerPointMtEnEncodingTaiwanWang = '\001\000\000\000',
    PowerPointMtEnEncodingIA5IRV = '\000\000\000\000',
    PowerPointMtEnEncodingIA5German = '\001\000\000\000',
    PowerPointMtEnEncodingIA5Swedish = '\000\000\000\000',
    PowerPointMtEnEncodingIA5Norwegian = '\001\000\000\000',
    PowerPointMtEnEncodingUSASCII = '\000\000\000\000',
    PowerPointMtEnEncodingT61 = '\001\000\000\000',
    PowerPointMtEnEncodingISO6937NonspacingAccent = '\000\000\000\000',
    PowerPointMtEnEncodingKOI8R = '\001\000\000\000',
    PowerPointMtEnEncodingExtAlphaLowercase = '\000\000\000\000',
    PowerPointMtEnEncodingKOI8U = '\001\000\000\000',
    PowerPointMtEnEncodingEuropa3 = '\000\000\000\000',
    PowerPointMtEnEncodingHZGBSimplifiedChinese = '\001\000\000\000',
    PowerPointMtEnEncodingUTF7 = '\000\000\000\000',
    PowerPointMtEnEncodingUTF8 = '\001\000\000\000'
} PowerPointMtEn;

typedef enum {
    PowerPoint4000CommandBar = 'msCB',
    PowerPoint4000CommandBarControl = 'mCBC'
} PowerPoint4000;

typedef enum {
    PowerPointMHyTHyperlinkRange = '\000\000\000\000',
    PowerPointMHyTHyperlinkShape = '\001\000\000\000',
    PowerPointMHyTHyperlinkInlineShape = '\000\000\000\000'
} PowerPointMHyT;

typedef enum {
    PowerPointPWnSWindowNormal = '\000\000\000\000',
    PowerPointPWnSWindowMinimized = '\001\000\000\000'
} PowerPointPWnS;

typedef enum {
    PowerPointPArSArrangeTiled = '\001\000\000\000',
    PowerPointPArSArrangeCascade = '\000\000\000\000'
} PowerPointPArS;

typedef enum {
    PowerPointPVTySlideView = '\000\000\000\000',
    PowerPointPVTyMasterView = '\001\000\000\000',
    PowerPointPVTyPageView = '\000\000\000\000',
    PowerPointPVTyHandoutMasterView = '\001\000\000\000',
    PowerPointPVTyNotesMasterView = '\000\000\000\000',
    PowerPointPVTyOutlineView = '\001\000\000\000',
    PowerPointPVTySlideSorterView = '\000\000\000\000',
    PowerPointPVTyTitleMasterView = '\001\000\000\000',
    PowerPointPVTyNormalView = '\000\000\000\000',
    PowerPointPVTyThumbnailView = '\001\000\000\000',
    PowerPointPVTyThumbnailMasterView = '\000\000\000\000'
} PowerPointPVTy;

typedef enum {
    PowerPointPCSiSchemeColorUnset = '\000\000\000\000',
    PowerPointPCSiNotASchemeColor = '\001\000\000\000',
    PowerPointPCSiBackgroundScheme = '\000\000\000\000',
    PowerPointPCSiForegroundScheme = '\001\000\000\000',
    PowerPointPCSiShadowScheme = '\000\000\000\000',
    PowerPointPCSiTitleScheme = '\001\000\000\000',
    PowerPointPCSiFillScheme = '\000\000\000\000',
    PowerPointPCSiAccent1Scheme = '\001\000\000\000',
    PowerPointPCSiAccent2Scheme = '\000\000\000\000',
    PowerPointPCSiAccent3Scheme = '\001\000\000\000'
} PowerPointPCSi;

typedef enum {
    PowerPointSSzTSlideSizeOnScreen = '\001\000\000\000',
    PowerPointSSzTSlideSizeLetterPaper = '\000\000\000\000',
    PowerPointSSzTSlideSizeA4Paper = '\001\000\000\000',
    PowerPointSSzTSlideSize35MM = '\000\000\000\000',
    PowerPointSSzTSlideSizeOverhead = '\001\000\000\000',
    PowerPointSSzTSlideSizeBanner = '\000\000\000\000',
    PowerPointSSzTSlideSizeCustom = '\001\000\000\000'
} PowerPointSSzT;

typedef enum {
    PowerPointPSATSaveAsPresentation = '\001\000\000\000',
    PowerPointPSATSaveAsTemplate = '\000\000\000\000',
    PowerPointPSATSaveAsRTF = '\001\000\000\000',
    PowerPointPSATSaveAsShow = '\000\000\000\000',
    PowerPointPSATSaveAsDefault = '\001\000\000\000',
    PowerPointPSATSaveAsHTML = '\000\000\000\000',
    PowerPointPSATSaveAsMovie = '\001\000\000\000',
    PowerPointPSATSaveAsPackage = '\000\000\000\000',
    PowerPointPSATSaveAsPDF = '\001\000\000\000',
    PowerPointPSATSaveAsOpenXMLPresentation = '\000\000\000\000',
    PowerPointPSATSaveAsOpenXMLPresentationMacroEnabled = '\001\000\000\000',
    PowerPointPSATSaveAsOpenXMLShow = '\000\000\000\000',
    PowerPointPSATSaveAsOpenXMLShowMacroEnabled = '\001\000\000\000',
    PowerPointPSATSaveAsOpenXMLTemplate = '\000\000\000\000',
    PowerPointPSATSaveAsOpenXMLTemplateMacroEnabled = '\001\000\000\000',
    PowerPointPSATSaveAsOpenXMLTheme = '\000\000\000\000'
} PowerPointPSAT;

typedef enum {
    PowerPointPTstTextStyleDefault = '\000\000\001*',
    PowerPointPTstTextStyleTitle = '\000\000\001*',
    PowerPointPTstTextStyleBody = '\000\000\001*'
} PowerPointPTst;

typedef enum {
    PowerPointPSLoSlideLayoutUnset = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTitleSlide = '\001\000\000\000',
    PowerPointPSLoSlideLayoutTextSlide = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTwoColumnText = '\001\000\000\000',
    PowerPointPSLoSlideLayoutTable = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTextAndChart = '\001\000\000\000',
    PowerPointPSLoSlideLayoutChartAndText = '\000\000\000\000',
    PowerPointPSLoSlideLayoutOrgchart = '\001\000\000\000',
    PowerPointPSLoSlideLayoutChart = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTextAndClipart = '\001\000\000\000',
    PowerPointPSLoSlideLayoutClipartAndText = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTitleOnly = '\001\000\000\000',
    PowerPointPSLoSlideLayoutBlank = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTextAndObject = '\001\000\000\000',
    PowerPointPSLoSlideLayoutObjectAndText = '\000\000\000\000',
    PowerPointPSLoSlideLayoutLargeObject = '\001\000\000\000',
    PowerPointPSLoSlideLayoutObject = '\000\000\000\000',
    PowerPointPSLoSlideLayoutMediaClip = '\001\000\000\000',
    PowerPointPSLoSlideLayoutMediaClipAndText = '\000\000\000\000',
    PowerPointPSLoSlideLayoutObjectOverText = '\001\000\000\000',
    PowerPointPSLoSlideLayoutTextOverObject = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTextAndTwoObjects = '\001\000\000\000',
    PowerPointPSLoSlideLayoutTwoObjectsAndText = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTwoObjectsOverText = '\001\000\000\000',
    PowerPointPSLoSlideLayoutFourObjects = '\000\000\000\000',
    PowerPointPSLoSlideLayoutVerticalText = '\001\000\000\000',
    PowerPointPSLoSlideLayoutClipartAndVerticalText = '\000\000\000\000',
    PowerPointPSLoSlideLayoutVerticalTitleAndText = '\001\000\000\000',
    PowerPointPSLoSlideLayoutVerticalTitleAndTextOverChart = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTwoObjects = '\001\000\000\000',
    PowerPointPSLoSlideLayoutObjectAndTwoObjects = '\000\000\000\000',
    PowerPointPSLoSlideLayoutTwoObjectsAndObject = '\001\000\000\000',
    PowerPointPSLoSlideLayoutCustom = '\000\000\000\000',
    PowerPointPSLoSlideLayoutSectionHeader = '\001\000\000\000',
    PowerPointPSLoSlideLayoutComparison = '\000\000\000\000',
    PowerPointPSLoSlideLayoutContentWithCaption = '\001\000\000\000',
    PowerPointPSLoSlideLayoutPictureWithCaption = '\000\000\000\000'
} PowerPointPSLo;

typedef enum {
    PowerPointPEeFEntryEffectUnset = '\000\000\000\000',
    PowerPointPEeFEntryEffectNone = '\001\000\000\000',
    PowerPointPEeFEntryEffectCut = '\000\000\000\000',
    PowerPointPEeFEntryEffectCutBlack = '\001\000\000\000',
    PowerPointPEeFEntryEffectRandom = '\000\000\000\000',
    PowerPointPEeFEntryEffectBlindsHorizontal = '\001\000\000\000',
    PowerPointPEeFEntryEffectBlindsVertical = '\000\000\000\000',
    PowerPointPEeFEntryEffectCheckerboardAcross = '\001\000\000\000',
    PowerPointPEeFEntryEffectCheckerboardDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectCoverLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectCoverUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectCoverRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectCoverDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectCoverLeftUp = '\001\000\000\000',
    PowerPointPEeFEntryEffectCoverRightUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectCoverLeftDown = '\001\000\000\000',
    PowerPointPEeFEntryEffectCoverRightDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectDissolve = '\001\000\000\000',
    PowerPointPEeFEntryEffectFadeBlack = '\000\000\000\000',
    PowerPointPEeFEntryEffectUncoverLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectUncoverUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectUncoverRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectUncoverDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectUncoverLeftUp = '\001\000\000\000',
    PowerPointPEeFEntryEffectUncoverRightUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectUncoverLeftDown = '\001\000\000\000',
    PowerPointPEeFEntryEffectUncoverRightDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectRandomBarsHorizontal = '\001\000\000\000',
    PowerPointPEeFEntryEffectRandomBarsVertical = '\000\000\000\000',
    PowerPointPEeFEntryEffectStripsLeftUp = '\001\000\000\000',
    PowerPointPEeFEntryEffectStripsRightUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectStripsLeftDown = '\001\000\000\000',
    PowerPointPEeFEntryEffectStripsRightDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectWipeLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectWipeUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectWipeRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectWipeDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectBoxOut = '\001\000\000\000',
    PowerPointPEeFEntryEffectBoxIn = '\000\000\000\000',
    PowerPointPEeFEntryEffectFlyFromLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectFlyFromTop = '\000\000\000\000',
    PowerPointPEeFEntryEffectFlyFromRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectFlyFromBottom = '\000\000\000\000',
    PowerPointPEeFEntryEffectFlyFromTopLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectFlyFromTopRight = '\000\000\000\000',
    PowerPointPEeFEntryEffectFlyFromBottomLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectFlyFromBottomRight = '\000\000\000\000',
    PowerPointPEeFEntryEffectPeekFromLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectPeekFromDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectPeekFromRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectPeekFromUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectCrawlFromLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectCrawlFromUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectCrawlFromRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectCrawlFromDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectZoomIn = '\001\000\000\000',
    PowerPointPEeFEntryEffectZoomInSlightly = '\000\000\000\000',
    PowerPointPEeFEntryEffectZoomOut = '\001\000\000\000',
    PowerPointPEeFEntryEffectZoomOutSlightly = '\000\000\000\000',
    PowerPointPEeFEntryEffectZoomCenter = '\001\000\000\000',
    PowerPointPEeFEntryEffectZoomBottom = '\000\000\000\000',
    PowerPointPEeFEntryEffectStretchAcross = '\001\000\000\000',
    PowerPointPEeFEntryEffectCollapseAcross = '\000\000\000\000',
    PowerPointPEeFEntryEffectStretchLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectCollapseLeft = '\000\000\000\000',
    PowerPointPEeFEntryEffectStretchUp = '\001\000\000\000',
    PowerPointPEeFEntryEffectCollapseUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectStretchRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectCollapseRight = '\000\000\000\000',
    PowerPointPEeFEntryEffectStretchDown = '\001\000\000\000',
    PowerPointPEeFEntryEffectCollapseBottom = '\000\000\000\000',
    PowerPointPEeFEntryEffectSwivel = '\001\000\000\000',
    PowerPointPEeFEntryEffectSpiral = '\000\000\000\000',
    PowerPointPEeFEntryEffectFadeFlyFromLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectFadeFlyFromTop = '\000\000\000\000',
    PowerPointPEeFEntryEffectFadeFlyFromRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectFadeFlyFromBottom = '\000\000\000\000',
    PowerPointPEeFEntryEffectFadeFlyFromTopLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectFadeFlyFromTopRight = '\000\000\000\000',
    PowerPointPEeFEntryEffectFadeFlyFromBottomLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectFadeFlyFromBottomRight = '\000\000\000\000',
    PowerPointPEeFEntryEffectSplitHorizontalOut = '\001\000\000\000',
    PowerPointPEeFEntryEffectSplitHorizontalIn = '\000\000\000\000',
    PowerPointPEeFEntryEffectSplitVerticalOut = '\001\000\000\000',
    PowerPointPEeFEntryEffectSplitVerticalIn = '\000\000\000\000',
    PowerPointPEeFEntryEffectFlashOnceFast = '\001\000\000\000',
    PowerPointPEeFEntryEffectFlashOnceMedium = '\000\000\000\000',
    PowerPointPEeFEntryEffectFlashOnceSlow = '\001\000\000\000',
    PowerPointPEeFEntryEffectAppear = '\000\000\000\000',
    PowerPointPEeFEntryEffectCircle = '\001\000\000\000',
    PowerPointPEeFEntryEffectDiamond = '\000\000\000\000',
    PowerPointPEeFEntryEffectCombHorizontal = '\001\000\000\000',
    PowerPointPEeFEntryEffectCombVertical = '\000\000\000\000',
    PowerPointPEeFEntryEffectFade = '\001\000\000\000',
    PowerPointPEeFEntryEffectFadeSmoothly = '\000\000\000\000',
    PowerPointPEeFEntryEffectNewsFlash = '\001\000\000\000',
    PowerPointPEeFEntryEffectSpinner = '\000\000\000\000',
    PowerPointPEeFEntryEffectPlus = '\001\000\000\000',
    PowerPointPEeFEntryEffectPushDown = '\000\000\000\000',
    PowerPointPEeFEntryEffectPushLeft = '\001\000\000\000',
    PowerPointPEeFEntryEffectPushRight = '\000\000\000\000',
    PowerPointPEeFEntryEffectPushUp = '\001\000\000\000',
    PowerPointPEeFEntryEffectWedge = '\000\000\000\000',
    PowerPointPEeFEntryEffectWheel1Spoke = '\001\000\000\000',
    PowerPointPEeFEntryEffectWheel2Spokes = '\000\000\000\000',
    PowerPointPEeFEntryEffectWheel3Spokes = '\001\000\000\000',
    PowerPointPEeFEntryEffectWheel4Spokes = '\000\000\000\000',
    PowerPointPEeFEntryEffectWheel8Spokes = '\001\000\000\000',
    PowerPointPEeFEntryEffectFlipLeft = '\000\000\000\000',
    PowerPointPEeFEntryEffectFlipRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectFlipUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectFlipDown = '\001\000\000\000',
    PowerPointPEeFEntryEffectCubeLeft = '\000\000\000\000',
    PowerPointPEeFEntryEffectCubeRight = '\001\000\000\000',
    PowerPointPEeFEntryEffectCubeUp = '\000\000\000\000',
    PowerPointPEeFEntryEffectCubeDown = '\001\000\000\000'
} PowerPointPEeF;

typedef enum {
    PowerPointPTlEAnimationLevelUnset = '\001\000\000\000',
    PowerPointPTlEAnimateLevelNone = '\000\000\000\000',
    PowerPointPTlEAnimateLevelFirstLevel = '\001\000\000\000',
    PowerPointPTlEAnimateLevelSecondLevel = '\000\000\000\000',
    PowerPointPTlEAnimateLevelThirdLevel = '\001\000\000\000',
    PowerPointPTlEAnimateLevelFourthLevel = '\000\000\000\000',
    PowerPointPTlEAnimateLevelFifthLevel = '\001\000\000\000',
    PowerPointPTlEAnimateLevelAllLevels = '\000\000\000\000'
} PowerPointPTlE;

typedef enum {
    PowerPointPTuEAnimationUnitUnset = '\000\000\000\000',
    PowerPointPTuETextUnitEffectByParagraph = '\001\000\000\000',
    PowerPointPTuETextUnitEffectByWord = '\000\000\000\000',
    PowerPointPTuETextUnitEffectByCharacter = '\001\000\000\000'
} PowerPointPTuE;

typedef enum {
    PowerPointPCuEAnimationChartUnset = '\001\000\000\000',
    PowerPointPCuEChartUnitEffectBySeries = '\000\000\000\000',
    PowerPointPCuEChartUnitEffectByCategory = '\001\000\000\000',
    PowerPointPCuEChartUnitEffectBySeriesElement = '\000\000\000\000'
} PowerPointPCuE;

typedef enum {
    PowerPointAsAeAfterEffectUnset = '\000\000\000\000',
    PowerPointAsAeAfterEffectNone = '\001\000\000\000',
    PowerPointAsAeAfterEffectHide = '\000\000\000\000',
    PowerPointAsAeAfterEffectDim = '\001\000\000\000'
} PowerPointAsAe;

typedef enum {
    PowerPointAdMdAdvanceModeUnset = '\001\000\000\000',
    PowerPointAdMdAdvanceModeOnClick = '\000\000\000\000'
} PowerPointAdMd;

typedef enum {
    PowerPointPSnXSoundEffectUnset = '\000\000\000\000',
    PowerPointPSnXSoundEffectNone = '\001\000\000\000',
    PowerPointPSnXSoundEffectStopPrevious = '\000\000\000\000',
    PowerPointPSnXSoundEffectFile = '\001\000\000\000'
} PowerPointPSnX;

typedef enum {
    PowerPointPUdOUpdateOptionUnset = '\001\000\000\000',
    PowerPointPUdOUpdateOptionManual = '\000\000\000\000'
} PowerPointPUdO;

typedef enum {
    PowerPointPDgMDialogModeUnset = '\001\000\000\000',
    PowerPointPDgMDialogModeModless = '\000\000\000\000',
    PowerPointPDgMDialogModeModal = '\001\000\000\000'
} PowerPointPDgM;

typedef enum {
    PowerPointPDgSDialogStyleUnset = '\001\000\000\000',
    PowerPointPDgSDialogStyleStandard = '\000\000\000\000'
} PowerPointPDgS;

typedef enum {
    PowerPointPSsPSlideShowPointerNone = '\000\000\000\000',
    PowerPointPSsPSlideShowPointerArrow = '\001\000\000\000',
    PowerPointPSsPSlideShowPointerPen = '\000\000\000\000',
    PowerPointPSsPSlideShowPointerAlwaysHidden = '\001\000\000\000'
} PowerPointPSsP;

typedef enum {
    PowerPointPShSSlideShowStateRunning = '\001\000\000\000',
    PowerPointPShSSlideShowStatePaused = '\000\000\000\000',
    PowerPointPShSSlideShowStateBlackScreen = '\001\000\000\000',
    PowerPointPShSSlideShowStateWhiteScreen = '\000\000\000\000'
} PowerPointPShS;

typedef enum {
    PowerPointPSaMSlideShowAdvanceManualAdvance = '\000\000\000\000',
    PowerPointPSaMSlideShowAdvanceUseSlideTimings = '\001\000\000\000'
} PowerPointPSaM;

typedef enum {
    PowerPointPtOtPrintSlides = '\001\000\000\000',
    PowerPointPtOtPrintTwoSlideHandouts = '\000\000\000\000',
    PowerPointPtOtPrintThreeSlideHandouts = '\001\000\000\000',
    PowerPointPtOtPrintSixSlideHandouts = '\000\000\000\000',
    PowerPointPtOtPrintNotesPages = '\001\000\000\000',
    PowerPointPtOtPrintOutline = '\000\000\000\000',
    PowerPointPtOtPrintFourSlideHandouts = '\001\000\000\000',
    PowerPointPtOtPrintNineSlideHandouts = '\000\000\000\000'
} PowerPointPtOt;

typedef enum {
    PowerPointPrCtPrintColor = '\000\000\000\000',
    PowerPointPrCtPrintBlackAndWhite = '\001\000\000\000'
} PowerPointPrCt;

typedef enum {
    PowerPointPSELSelectionTypeNone = '\001\000\000\000',
    PowerPointPSELSelectionTypeSlides = '\000\000\000\000',
    PowerPointPSELSelectionTypeShapes = '\001\000\000\000',
    PowerPointPSELSelectionTypeText = '\000\000\000\000'
} PowerPointPSEL;

typedef enum {
    PowerPointPDtFUnset = '\001\000\000\000',
    PowerPointPDtFMdyy = '\000\000\000\000',
    PowerPointPDtFDdddMMMMddyyyy = '\001\000\000\000',
    PowerPointPDtFMMMMyyyy = '\000\000\000\000',
    PowerPointPDtFMMMMdyyyy = '\001\000\000\000',
    PowerPointPDtFMMMyy = '\000\000\000\000',
    PowerPointPDtFMMMMyy = '\001\000\000\000',
    PowerPointPDtFMMyy = '\000\000\000\000',
    PowerPointPDtFMMddyyHmm = '\001\000\000\000',
    PowerPointPDtFMddyyhmmAMPM = '\000\000\000\000',
    PowerPointPDtFHmm = '\001\000\000\000',
    PowerPointPDtFHmmss = '\000\000\000\000',
    PowerPointPDtFHmmAMPM = '\001\000\000\000',
    PowerPointPDtFHmmssAMPM = '\000\000\000\000'
} PowerPointPDtF;

typedef enum {
    PowerPointPTnSTransitionSpeedUnset = '\000\000\000\000',
    PowerPointPTnSTransistionSpeedSlow = '\001\000\000\000',
    PowerPointPTnSTransistionSpeedMedium = '\000\000\000\000'
} PowerPointPTnS;

typedef enum {
    PowerPointPMtvMouseActivationMouseClick = '\000\000\000\000',
    PowerPointPMtvMouseActivationMouseOver = '\001\000\000\000'
} PowerPointPMtv;

typedef enum {
    PowerPointPAxTActionTypeUnset = '\001\000\000\000',
    PowerPointPAxTActionTypeNone = '\000\000\000\000',
    PowerPointPAxTActionTypeNextSlide = '\001\000\000\000',
    PowerPointPAxTActionTypePreviousSlide = '\000\000\000\000',
    PowerPointPAxTActionTypeFirstSlide = '\001\000\000\000',
    PowerPointPAxTActionTypeLastSlide = '\000\000\000\000',
    PowerPointPAxTActionTypeLastSlideViewed = '\001\000\000\000',
    PowerPointPAxTActionTypeEndShow = '\000\000\000\000',
    PowerPointPAxTActionTypeHyperlinkAction = '\001\000\000\000',
    PowerPointPAxTActionTypeRunMacro = '\000\000\000\000',
    PowerPointPAxTActionTypeRunProgram = '\001\000\000\000',
    PowerPointPAxTActionTypeNamedSlideshowAction = '\000\000\000\000',
    PowerPointPAxTActionTypeOLEVerb = '\001\000\000\000'
} PowerPointPAxT;

typedef enum {
    PowerPointPPhdPlaceholderTypeUnset = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeTitlePlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypeBodyPlaceholder = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeCenterTitlePlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypeSubtitlePlaceholder = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeVerticalTitlePlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypeVerticalBodyPlaceholder = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeObjectPlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypeChartPlaceholder = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeBitmapPlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypeMediaClipPlaceholder = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeOrgChartPlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypeTablePlaceholder = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeSlideNumberPlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypeHeaderPlaceholder = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeFooterPlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypeDatePlaceholder = '\001\000\000\000',
    PowerPointPPhdPlaceholderTypeVerticalObjectPlaceholder = '\000\000\000\000',
    PowerPointPPhdPlaceholderTypePicturePlaceholder = '\001\000\000\000'
} PowerPointPPhd;

typedef enum {
    PowerPointPSStSlideShowTypeSpeaker = '\001\000\000\000',
    PowerPointPSStSlideShowTypeWindow = '\000\000\000\000',
    PowerPointPSStSlideShowTypePresenter = '\001\000\000\000',
    PowerPointPSStSlideShowTypeKiosk = '\000\000\000\000'
} PowerPointPSSt;

typedef enum {
    PowerPointRgTyPrintRangeAll = '\000\000\000\000',
    PowerPointRgTyPrintRangeSelection = '\001\000\000\000',
    PowerPointRgTyPrintRangeCurrent = '\000\000\000\000',
    PowerPointRgTyPrintRangeSlideRange = '\001\000\000\000'
} PowerPointRgTy;

typedef enum {
    PowerPointMedTMediaTypeUnset = '\000\000\000\000',
    PowerPointMedTMediaTypeOther = '\001\000\000\000',
    PowerPointMedTMediaTypeSound = '\000\000\000\000',
    PowerPointMedTMediaTypeMovie = '\001\000\000\000'
} PowerPointMedT;

typedef enum {
    PowerPointPSFySoundFormatUnset = '\001\000\000\000',
    PowerPointPSFySoundFormatNone = '\000\000\000\000',
    PowerPointPSFySoundFormatWAV = '\001\000\000\000',
    PowerPointPSFySoundFormatMIDI = '\000\000\000\000'
} PowerPointPSFy;

typedef enum {
    PowerPointPeBlEastAsianLineBreakNormal = '\000\000\000\000',
    PowerPointPeBlEastAsianLineBreakStrict = '\001\000\000\000',
    PowerPointPeBlEastAsianLineBreakCustom = '\000\000\000\000'
} PowerPointPeBl;

typedef enum {
    PowerPointSRgTSlideShowRangeShowAll = '\000\000\000\000',
    PowerPointSRgTSlideShowRange = '\001\000\000\000',
    PowerPointSRgTSlideShowRangeNamedSlideshow = '\000\000\000\000'
} PowerPointSRgT;

typedef enum {
    PowerPointFClrFrameColorsBrowserColors = '\000\000\000\000',
    PowerPointFClrFrameColorsPresentationSchemeTextColor = '\001\000\000\000',
    PowerPointFClrFrameColorsPresentationSchemeAccentColor = '\000\000\000\000',
    PowerPointFClrFrameColorsWhiteTextOnBlack = '\001\000\000\000',
    PowerPointFClrFrameColorsBlackTextOnWhite = '\000\000\000\000'
} PowerPointFClr;

typedef enum {
    PowerPointPMOtMovieOptimizationNormal = '\000\000\000\000',
    PowerPointPMOtMovieOptimizationSize = '\001\000\000\000',
    PowerPointPMOtMovieOptimizationSpeed = '\000\000\000\000',
    PowerPointPMOtMovieOptimizationQuality = '\001\000\000\000'
} PowerPointPMOt;

typedef enum {
    PowerPointPShFShapeFormatGIF = '\001\000\000\000',
    PowerPointPShFShapeFormatJPEG = '\000\000\000\000',
    PowerPointPShFShapeFormatPNG = '\001\000\000\000',
    PowerPointPShFShapeFormatPICT = '\000\000\000\000'
} PowerPointPShF;

typedef enum {
    PowerPointPBrTTopBorder = '\000\000\001\032',
    PowerPointPBrTLeftBorder = '\000\000\001\032',
    PowerPointPBrTBottomBorder = '\000\000\001\032',
    PowerPointPBrTRightBorder = '\000\000\001\032',
    PowerPointPBrTDiagonalDownBorder = '\000\000\001\032',
    PowerPointPBrTDiagonalUpBorder = '\000\000\001\032'
} PowerPointPBrT;

typedef enum {
    PowerPointPALOPageLayoutNormal = '\000\000\000\000',
    PowerPointPALOPageLayoutFullScreen = '\001\000\000\000'
} PowerPointPALO;

typedef enum {
    PowerPointPBuTRegular = '\001\000\000\000',
    PowerPointPBuTFancy = '\000\000\000\000',
    PowerPointPBuTTextOnly = '\001\000\000\000'
} PowerPointPBuT;

typedef enum {
    PowerPointPNBpBarPlacementTop = '\001\000\000\000',
    PowerPointPNBpBarPlacementBottom = '\000\000\000\000'
} PowerPointPNBp;

typedef enum {
    PowerPointAnFXAnimationTypeCustom = '\000\000\001\002',
    PowerPointAnFXAnimationTypeAppear = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFly = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBlinds = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBox = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCheckerboard = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCircle = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCrawl = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDiamond = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDissolve = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFade = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFlashOnce = '\000\000\001\002',
    PowerPointAnFXAnimationTypePeek = '\000\000\001\002',
    PowerPointAnFXAnimationTypePlus = '\000\000\001\002',
    PowerPointAnFXAnimationTypeRandomBars = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSpiral = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSplit = '\000\000\001\002',
    PowerPointAnFXAnimationTypeStretch = '\000\000\001\002',
    PowerPointAnFXAnimationTypeStrips = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSwivel = '\000\000\001\002',
    PowerPointAnFXAnimationTypeWedge = '\000\000\001\002',
    PowerPointAnFXAnimationTypeWheel = '\000\000\001\002',
    PowerPointAnFXAnimationTypeWipe = '\000\000\001\002',
    PowerPointAnFXAnimationTypeZoom = '\000\000\001\002',
    PowerPointAnFXAnimationTypeRandomEffect = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBoomerang = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBounce = '\000\000\001\002',
    PowerPointAnFXAnimationTypeColorReveal = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCredits = '\000\000\001\002',
    PowerPointAnFXAnimationTypeEaseIn = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFloat = '\000\000\001\002',
    PowerPointAnFXAnimationTypeGrowAndTurn = '\000\000\001\002',
    PowerPointAnFXAnimationTypeLightSpeed = '\000\000\001\002',
    PowerPointAnFXAnimationTypePinwheel = '\000\000\001\002',
    PowerPointAnFXAnimationTypeRiseUp = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSwish = '\000\000\001\002',
    PowerPointAnFXAnimationTypeThinLine = '\000\000\001\002',
    PowerPointAnFXAnimationTypeUnfold = '\000\000\001\002',
    PowerPointAnFXAnimationTypeWhip = '\000\000\001\002',
    PowerPointAnFXAnimationTypeAscend = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCenterRevolve = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFadedSwivel = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDescend = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSling = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSpinner = '\000\000\001\002',
    PowerPointAnFXAnimationTypeStretchy = '\000\000\001\002',
    PowerPointAnFXAnimationTypeZip = '\000\000\001\002',
    PowerPointAnFXAnimationTypeArcUp = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFadeZoom = '\000\000\001\002',
    PowerPointAnFXAnimationTypeGlide = '\000\000\001\002',
    PowerPointAnFXAnimationTypeExpand = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFlip = '\000\000\001\002',
    PowerPointAnFXAnimationTypeShimmer = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFold = '\000\000\001\002',
    PowerPointAnFXAnimationTypeChangeFillColor = '\000\000\001\002',
    PowerPointAnFXAnimationTypeChangeFont = '\000\000\001\002',
    PowerPointAnFXAnimationTypeChangeFontColor = '\000\000\001\002',
    PowerPointAnFXAnimationTypeChangeFontSize = '\000\000\001\002',
    PowerPointAnFXAnimationTypeChangeFontStyle = '\000\000\001\002',
    PowerPointAnFXAnimationTypeGrowShrink = '\000\000\001\002',
    PowerPointAnFXAnimationTypeChangeLineColor = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSpin = '\000\000\001\002',
    PowerPointAnFXAnimationTypeTransparency = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBoldFlash = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBlast = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBoldReveal = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBrushOnColor = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBrushOnUnderline = '\000\000\001\002',
    PowerPointAnFXAnimationTypeColorBlend = '\000\000\001\002',
    PowerPointAnFXAnimationTypeColorWave = '\000\000\001\002',
    PowerPointAnFXAnimationTypeComplementaryColor = '\000\000\001\002',
    PowerPointAnFXAnimationTypeComplementaryColor2 = '\000\000\001\002',
    PowerPointAnFXAnimationTypeConstrastingColor = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDarken = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDesaturate = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFlashBulb = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFlicker = '\000\000\001\002',
    PowerPointAnFXAnimationTypeGrowWithColor = '\000\000\001\002',
    PowerPointAnFXAnimationTypeLighten = '\000\000\001\002',
    PowerPointAnFXAnimationTypeStyleEmphasis = '\000\000\001\002',
    PowerPointAnFXAnimationTypeTeeter = '\000\000\001\002',
    PowerPointAnFXAnimationTypeVerticalGrow = '\000\000\001\002',
    PowerPointAnFXAnimationTypeWave = '\000\000\001\002',
    PowerPointAnFXAnimationTypeMediaPlay = '\000\000\001\002',
    PowerPointAnFXAnimationTypeMediaPause = '\000\000\001\002',
    PowerPointAnFXAnimationTypeMediaStop = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCirclePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeRightTrianglePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDiamondPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeHexagonPath = '\000\000\001\002',
    PowerPointAnFXAnimationType5PointStarPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCrescentMoonPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSquarePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeTrapezoidPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeHeartPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeOctagonPath = '\000\000\001\002',
    PowerPointAnFXAnimationType6PointStarPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFootballPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeEqualTrianglePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeParallelogramPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypePentagonPath = '\000\000\001\002',
    PowerPointAnFXAnimationType4PointStarPath = '\000\000\001\002',
    PowerPointAnFXAnimationType8PointStarPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeTeardropPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypePointyStarPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCurvedSquarePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCurvedXPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeVerticalFigure8Path = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCurvyStarPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeLoopDeLoopPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBuzzsawPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeHorizontalFigure8Path = '\000\000\001\002',
    PowerPointAnFXAnimationTypePeanutPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFigure8FourPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeNeutronPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSwooshPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBeanPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypePlusPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeInvertedTrianglePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeInvertedSquarePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeLeftPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeTurnRightPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeArcDownPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeZigzagPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSCurve2Path = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSineWavePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBounceLeftPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDownPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeTurnUpPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeArcUpPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeHeartbeatPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSpiralRightPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeWavePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCurvyLeftPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDiagonalDownRightPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeTurnDownPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeArcLeftPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeFunnelPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSpringPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeBounceRightPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSpiralLeftPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDiagonalUpRightPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeTurnUpRightPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeArcRightPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeSCurve1Path = '\000\000\001\002',
    PowerPointAnFXAnimationTypeDecayingWavePath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeCurvyRightPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeStairsDownPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeUpPath = '\000\000\001\002',
    PowerPointAnFXAnimationTypeRightPath = '\000\000\001\002'
} PowerPointAnFX;

typedef enum {
    PowerPointAnLvTextByNoLevels = '\000\000\001\001',
    PowerPointAnLvTextByAllLevels = '\000\000\001\001',
    PowerPointAnLvTextByFirstLevel = '\000\000\001\001',
    PowerPointAnLvTextBySecondLevel = '\000\000\001\001',
    PowerPointAnLvTextByThirdLevel = '\000\000\001\001',
    PowerPointAnLvTextByFourthLevel = '\000\000\001\001',
    PowerPointAnLvTextByFifthLevel = '\000\000\001\001',
    PowerPointAnLvChartAllAtOnce = '\000\000\001\001',
    PowerPointAnLvChartByCategory = '\000\000\001\001',
    PowerPointAnLvChartByCtageoryElements = '\000\000\001\001',
    PowerPointAnLvChartBySeries = '\000\000\001\001',
    PowerPointAnLvChartBySeriesElements = '\000\000\001\001'
} PowerPointAnLv;

typedef enum {
    PowerPointAnTrNoTrigger = '\000\000\000\001',
    PowerPointAnTrOnPageClick = '\000\000\000\001',
    PowerPointAnTrWithPrevious = '\000\000\000\001',
    PowerPointAnTrAfterPrevious = '\000\000\000\001',
    PowerPointAnTrOnShapeClick = '\000\000\000\001'
} PowerPointAnTr;

typedef enum {
    PowerPointAnAENoAfterEffect = '\000\000\000\000',
    PowerPointAnAEDim = '\001\000\000\000',
    PowerPointAnAEHide = '\000\000\000\000',
    PowerPointAnAEHideOnNextClick = '\001\000\000\000'
} PowerPointAnAE;

typedef enum {
    PowerPointAnTUByParagraph = '\001\000\000\000',
    PowerPointAnTUByCharacter = '\000\000\000\000',
    PowerPointAnTUByWord = '\001\000\000\000'
} PowerPointAnTU;

typedef enum {
    PowerPointAnRsRestartAlways = '\001\000\000\000',
    PowerPointAnRsRestartWhenOff = '\000\000\000\000',
    PowerPointAnRsNeverRestart = '\001\000\000\000'
} PowerPointAnRs;

typedef enum {
    PowerPointAnEAAfterFreeze = '\001\000\000\000',
    PowerPointAnEAAfterRemove = '\000\000\000\000',
    PowerPointAnEAAfterHold = '\001\000\000\000',
    PowerPointAnEAAfterTransition = '\000\000\000\000'
} PowerPointAnEA;

typedef enum {
    PowerPointAnDiNoDirection = '\000\000\000\000',
    PowerPointAnDiUp = '\001\000\000\000',
    PowerPointAnDiRight = '\000\000\000\000',
    PowerPointAnDiDown = '\001\000\000\000',
    PowerPointAnDiLeft = '\000\000\000\000',
    PowerPointAnDiOrdinalMask = '\001\000\000\000',
    PowerPointAnDiUpLeft = '\000\000\000\000',
    PowerPointAnDiUpRight = '\001\000\000\000',
    PowerPointAnDiDownRight = '\000\000\000\000',
    PowerPointAnDiDownLeft = '\001\000\000\000',
    PowerPointAnDiTop = '\000\000\000\000',
    PowerPointAnDiBottom = '\001\000\000\000',
    PowerPointAnDiTopLeft = '\000\000\000\000',
    PowerPointAnDiTopRight = '\001\000\000\000',
    PowerPointAnDiBottomRight = '\000\000\000\000',
    PowerPointAnDiBottomLeft = '\001\000\000\000',
    PowerPointAnDiHorizontal = '\000\000\000\000',
    PowerPointAnDiVertical = '\001\000\000\000',
    PowerPointAnDiAcross = '\000\000\000\000',
    PowerPointAnDiInward = '\001\000\000\000',
    PowerPointAnDiOut = '\000\000\000\000',
    PowerPointAnDiClockwise = '\001\000\000\000',
    PowerPointAnDiCounterclockwise = '\000\000\000\000',
    PowerPointAnDiHorizontalIn = '\001\000\000\000',
    PowerPointAnDiHorizontalOut = '\000\000\000\000',
    PowerPointAnDiVerticalIn = '\001\000\000\000',
    PowerPointAnDiVerticalOut = '\000\000\000\000',
    PowerPointAnDiSlightly = '\001\000\000\000',
    PowerPointAnDiCenter = '\000\000\000\000',
    PowerPointAnDiInSlightly = '\001\000\000\000',
    PowerPointAnDiInCenter = '\000\000\000\000',
    PowerPointAnDiInBottom = '\001\000\000\000',
    PowerPointAnDiOutSlightly = '\000\000\000\000',
    PowerPointAnDiOutCenter = '\001\000\000\000',
    PowerPointAnDiOutBottom = '\000\000\000\000',
    PowerPointAnDiFontBold = '\001\000\000\000',
    PowerPointAnDiFontItalic = '\000\000\000\000',
    PowerPointAnDiFontUnderline = '\001\000\000\000',
    PowerPointAnDiFontStrikethrough = '\000\000\000\000',
    PowerPointAnDiFontShadow = '\001\000\000\000',
    PowerPointAnDiFontAllCaps = '\000\000\000\000',
    PowerPointAnDiInstant = '\001\000\000\000',
    PowerPointAnDiGradual = '\000\000\000\000',
    PowerPointAnDiCycleClockwise = '\001\000\000\000',
    PowerPointAnDiCycleCounterclockwise = '\000\000\000\000'
} PowerPointAnDi;

typedef enum {
    PowerPointAnTyAnimationTypeNone = '\000\000\001\003',
    PowerPointAnTyAnimationTypeMotion = '\000\000\001\003',
    PowerPointAnTyAnimationTypeColor = '\000\000\001\003',
    PowerPointAnTyAnimationTypeScale = '\000\000\001\003',
    PowerPointAnTyAnimationTypeRotation = '\000\000\001\003',
    PowerPointAnTyAnimationTypeProperty = '\000\000\001\003',
    PowerPointAnTyAnimationTypeCommand = '\000\000\001\003',
    PowerPointAnTyAnimationTypeFilter = '\000\000\001\003',
    PowerPointAnTyAnimationTypeSet = '\000\000\001\003'
} PowerPointAnTy;

typedef enum {
    PowerPointAnppNoAdditive = '\000\000\001\007',
    PowerPointAnppMotion = '\000\000\001\007'
} PowerPointAnpp;

typedef enum {
    PowerPointAnSmNoAccumulate = '\000\000\001\004',
    PowerPointAnSmAlways = '\000\000\001\004'
} PowerPointAnSm;

typedef enum {
    PowerPointAnPrNoProperty = '\000\000\001\005',
    PowerPointAnPrX = '\000\000\001\005',
    PowerPointAnPrY = '\000\000\001\005',
    PowerPointAnPrWidth = '\000\000\001\005',
    PowerPointAnPrHeight = '\000\000\001\005',
    PowerPointAnPrOpacity = '\000\000\001\005',
    PowerPointAnPrRotation = '\000\000\001\005',
    PowerPointAnPrColors = '\000\000\001\005',
    PowerPointAnPrVisibility = '\000\000\001\005',
    PowerPointAnPrTextFontBold = '\000\000\001\005',
    PowerPointAnPrTextFontColor = '\000\000\001\005',
    PowerPointAnPrTextFontEmboss = '\000\000\001\005',
    PowerPointAnPrTextFontItalic = '\000\000\001\005',
    PowerPointAnPrTextFontName = '\000\000\001\005',
    PowerPointAnPrTextFontShadow = '\000\000\001\005',
    PowerPointAnPrTextFontSize = '\000\000\001\005',
    PowerPointAnPrTextFontSubscript = '\000\000\001\005',
    PowerPointAnPrTextFontSuperscript = '\000\000\001\005',
    PowerPointAnPrTextFontUnderline = '\000\000\001\005',
    PowerPointAnPrTextFontStrikethrough = '\000\000\001\005',
    PowerPointAnPrTextBulletCharacter = '\000\000\001\005',
    PowerPointAnPrTextBulletFontName = '\000\000\001\005',
    PowerPointAnPrTextBulletNumber = '\000\000\001\005',
    PowerPointAnPrTextBulletColor = '\000\000\001\005',
    PowerPointAnPrTextBulletRelativeSize = '\000\000\001\005',
    PowerPointAnPrTextBulletStyle = '\000\000\001\005',
    PowerPointAnPrTextBulletType = '\000\000\001\005',
    PowerPointAnPrShapePictureContrast = '\001\005\003\350',
    PowerPointAnPrShapePictureBrightness = '\001\005\003\351',
    PowerPointAnPrShapePictureGamma = '\001\005\003\352',
    PowerPointAnPrShapePictureGrayscale = '\001\005\003\353',
    PowerPointAnPrShapeFillOn = '\001\005\003\354',
    PowerPointAnPrShapeFillColor = '\001\005\003\355',
    PowerPointAnPrShapeFillOpacity = '\001\005\003\356',
    PowerPointAnPrShapeFillBackColor = '\001\005\003\357',
    PowerPointAnPrShapeLineOn = '\001\005\003\360',
    PowerPointAnPrShapeLineColor = '\001\005\003\361',
    PowerPointAnPrShapeShadowOn = '\001\005\003\362',
    PowerPointAnPrShapeShadowType = '\001\005\003\363',
    PowerPointAnPrShapeShadowColor = '\001\005\003\364',
    PowerPointAnPrShapeShadowOpacity = '\001\005\003\365',
    PowerPointAnPrShapeShadowOffsetX = '\001\005\003\366',
    PowerPointAnPrShapeShadowOffsetY = '\001\005\003\367'
} PowerPointAnPr;

typedef enum {
    PowerPointAnCTEvent = '\000\000\001\006',
    PowerPointAnCTCall = '\000\000\001\006',
    PowerPointAnCTVerb = '\000\000\001\006'
} PowerPointAnCT;

typedef enum {
    PowerPointAfetNoFilterEffectType = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeBarn = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeBlinds = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeBox = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeCheckerboard = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeCircle = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeDiamond = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeDissolve = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeFade = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeImage = '\000\000\001\010',
    PowerPointAfetFilterEffectTypePixelate = '\000\000\001\010',
    PowerPointAfetFilterEffectTypePlus = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeRandomBar = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeSlide = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeStretch = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeStrips = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeWedge = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeWheel = '\000\000\001\010',
    PowerPointAfetFilterEffectTypeWipe = '\000\000\001\010'
} PowerPointAfet;

typedef enum {
    PowerPointAfesNoEffectSubtype = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeInVertical = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeOutVertical = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeInHorizontal = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeOutHorizontal = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeHorizontal = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeVertical = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeInward = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeOut = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeAcross = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeFromLeft = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeFromRight = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeFromTop = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeFromBottom = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeDownLeft = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeUpLeft = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeDownRight = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeUpRight = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeSpoke1 = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeSpokes2 = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeSpokes3 = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeSpokes4 = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeSpokes8 = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeLeft = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeRight = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeDown = '\000\000\001 ',
    PowerPointAfesFilterEffectSubtypeUp = '\000\000\001 '
} PowerPointAfes;

typedef enum {
    PowerPoint4002View = 'pVEW',
    PowerPoint4002Presentation = 'pptP'
} PowerPoint4002;

typedef enum {
    PowerPoint4001Slide = 'pSLD',
    PowerPoint4001Master = 'pMtr',
    PowerPoint4001Presentation = 'pptP'
} PowerPoint4001;

typedef enum {
    PowerPoint4010Shape = 'pShp',
    PowerPoint4010FillFormat = 'pFFm'
} PowerPoint4010;

typedef enum {
    PowerPoint4015Shape = 'pShp',
    PowerPoint4015FillFormat = 'pFFm'
} PowerPoint4015;

typedef enum {
    PowerPoint4023Callout = 'cD00',
    PowerPoint4023CalloutFormat = 'cCoF'
} PowerPoint4023;

typedef enum {
    PowerPoint4018Connector = 'cD01',
    PowerPoint4018ConnectorFormat = 'pCxF'
} PowerPoint4018;

typedef enum {
    PowerPoint4019Connector = 'cD01',
    PowerPoint4019ConnectorFormat = 'pCxF'
} PowerPoint4019;

typedef enum {
    PowerPoint4025Callout = 'cD00',
    PowerPoint4025CalloutFormat = 'cCoF'
} PowerPoint4025;

typedef enum {
    PowerPoint4024Callout = 'cD00',
    PowerPoint4024CalloutFormat = 'cCoF'
} PowerPoint4024;

typedef enum {
    PowerPoint4020Connector = 'cD01',
    PowerPoint4020ConnectorFormat = 'pCxF'
} PowerPoint4020;

typedef enum {
    PowerPoint4021Connector = 'cD01',
    PowerPoint4021ConnectorFormat = 'pCxF'
} PowerPoint4021;

typedef enum {
    PowerPoint4011Shape = 'pShp',
    PowerPoint4011FillFormat = 'pFFm'
} PowerPoint4011;

typedef enum {
    PowerPoint4009Shape = 'pShp',
    PowerPoint4009ShapeRange = 'ShpR'
} PowerPoint4009;

typedef enum {
    PowerPoint4014Shape = 'pShp',
    PowerPoint4014FillFormat = 'pFFm'
} PowerPoint4014;

typedef enum {
    PowerPoint4022Shape = 'pShp',
    PowerPoint4022ThreeDFormat = 'D3Df'
} PowerPoint4022;

typedef enum {
    PowerPoint4016Shape = 'pShp',
    PowerPoint4016FillFormat = 'pFFm'
} PowerPoint4016;

typedef enum {
    PowerPoint4017Shape = 'pShp',
    PowerPoint4017FillFormat = 'pFFm'
} PowerPoint4017;

typedef enum {
    PowerPoint4008Shape = 'pShp',
    PowerPoint4008ShapeRange = 'ShpR'
} PowerPoint4008;

typedef enum {
    PowerPoint4013Shape = 'pShp',
    PowerPoint4013FillFormat = 'pFFm'
} PowerPoint4013;

typedef enum {
    PowerPoint4012Shape = 'pShp',
    PowerPoint4012FillFormat = 'pFFm'
} PowerPoint4012;

typedef enum {
    PowerPoint4003Shape = 'pShp',
    PowerPoint4003ShapeRange = 'ShpR'
} PowerPoint4003;

typedef enum {
    PowerPoint4007Shape = 'pShp',
    PowerPoint4007ShapeRange = 'ShpR'
} PowerPoint4007;

typedef enum {
    PowerPoint4004Shape = 'pShp',
    PowerPoint4004ShapeRange = 'ShpR'
} PowerPoint4004;

typedef enum {
    PowerPoint4005Shape = 'pShp',
    PowerPoint4005ShapeRange = 'ShpR'
} PowerPoint4005;

typedef enum {
    PowerPoint4006Shape = 'pShp',
    PowerPoint4006ShapeRange = 'ShpR'
} PowerPoint4006;



/*
 * Standard Suite
 */

// A scriptable object
@interface PowerPointBaseObject : SBObject

@property (copy) NSDictionary *properties;  // All of the object's properties

- (void) closeSaving:(PowerPointSavo)saving savingIn:(PowerPointPPfd)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) open;  // Open the specified object(s)
- (void) saveIn:(PowerPointPPfd)in_ as:(PowerPointPPff)as;  // Save an object
- (void) select;  // Make a selection
- (void) quit;
- (void) setPrinterSettingsPrinterSettings:(NSInteger)printerSettings;

@end

// A basic application program
@interface PowerPointBaseApplication : PowerPointBaseObject

@property (readonly) BOOL frontmost;  // Is this the frontmost application?
@property (copy, readonly) NSString *name;  // the name
@property (copy, readonly) NSString *version;  // the version of the application


@end

// Every document
@interface PowerPointBaseDocument : PowerPointBaseObject

@property (readonly) BOOL modified;  // Has the document been modified since the last save?
@property (copy, readonly) NSString *name;  // the name


@end

// Every basic window
@interface PowerPointBasicWindow : PowerPointBaseObject

@property NSRect bounds;  // the boundary rectangle for the window
@property (readonly) BOOL closeable;  // Does the window have a close box?
@property (readonly) BOOL titled;  // Does the window have a title bar?
@property (readonly) NSInteger entryIndex;  // the number of the window
@property (readonly) BOOL floating;  // Does the window float?
@property (readonly) BOOL modal;  // Is the window modal?
@property NSPoint position;  // upper left coordinates of the window
@property (readonly) BOOL resizable;  // Is the window resizable?
@property (readonly) BOOL zoomable;  // Is the window zoomable?
@property BOOL zoomed;  // Is the window zoomed?
@property (copy, readonly) NSString *name;  // the title of the window
@property (readonly) BOOL visible;  // Is the window visible?
@property (readonly) BOOL collapsable;  // Is the window collapasable?
@property BOOL collapsed;  // Is the window collapsed?
@property (readonly) BOOL sheet;  // Is this window a sheet window?


@end

@interface PowerPointPrintSettings : SBObject

@property (readonly) NSInteger copies;  // the number of copies of a document to be printed 
@property (readonly) BOOL collating;  // Should printed copies be collated?
@property (readonly) NSInteger startingPage;  // the first page of the document to be printed
@property (readonly) NSInteger endingPage;  // the last page of the document to be printed
@property (readonly) NSInteger pagesAcross;  // number of logical pages laid across a physical page
@property (readonly) NSInteger pagesDown;  // number of logical pages laid out down a physical page
@property (copy, readonly) NSDate *requestedPrintTime;  // the time at which the desktop printer should print the document...
@property (copy, readonly) id errorHandling;  // how errors are handled
@property (copy, readonly) NSString *faxNumber;  // for fax number
@property (copy, readonly) NSString *targetPrinter;  // the queue name of the target printer

- (void) closeSaving:(PowerPointSavo)saving savingIn:(PowerPointPPfd)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) open;  // Open the specified object(s)
- (void) saveIn:(PowerPointPPfd)in_ as:(PowerPointPPff)as;  // Save an object
- (void) select;  // Make a selection
- (void) quit;
- (void) setPrinterSettingsPrinterSettings:(NSInteger)printerSettings;

@end



/*
 * Microsoft Office Suite
 */

// Every command bar control
@interface PowerPointCommandBarControl : PowerPointBaseObject

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) PowerPointMCLT controlType;  // Returns the type of the command bar control.
@property (copy) NSString *descriptionText;  // Returns or sets the description for a command bar control.  The description is not displayed to the user, but it can be useful for documenting the behavior of a control.
@property BOOL enabled;  // Returns or sets if the command bar control is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number for this command bar control.
@property NSInteger height;  // Returns or sets the height of a command bar control.
@property NSInteger helpContextID;  // Returns or sets the help context ID number for the Help topic attached to the command bar control.
@property (copy) NSString *helpFile;  // Returns or sets the file name for the help topic attached to the command bar.  To use this property, you must also set the help context ID property.
- (NSInteger) id;  // Returns the id for a built-in command bar control.
@property (readonly) NSInteger leftPosition;  // Returns the left position of the command bar control.
@property (copy) NSString *name;  // Returns or sets the caption text for a command bar control.
@property (copy) NSString *parameter;  // Returns or sets a string that is used to execute a command.
@property NSInteger priority;  // Returns or sets the priority of a command bar control. A controls priority determines whether the control can be dropped from a docked command bar if the command bar controls can not fit in a single row.  Valid priority number are 0 through 7.
@property (copy) NSString *tag;  // Returns or sets information about the command bar control, such as data that can be used as an argument in procedures, or information that identifies the control.
@property (copy) NSString *tooltipText;  // Returns or sets the text displayed in a command bar controls tooltip.
@property (readonly) NSInteger top;  // Returns the top position of a command bar control.
@property BOOL visible;  // Returns or sets if the command bar control is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar control.

- (void) execute;  // Runs the procedure or built-in command assigned to the specified command bar control.

@end

// Every command bar button
@interface PowerPointCommandBarButton : PowerPointCommandBarControl

@property (readonly) BOOL buttonFaceIsDefault;  // Returns if the face of a command bar button control is the original built-in face.
@property PowerPointMBns buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property PowerPointMBTs buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// Every command bar combobox
@interface PowerPointCommandBarCombobox : PowerPointCommandBarControl

@property PowerPointMXcb comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
@property (copy) NSString *comboboxText;  // Returns or sets the text in the display or edit portion of the command bar combobox control.
@property NSInteger dropDownLines;  // Returns or sets the number of lines in a command bar control combobox control.  The combobox control must be a custom control.
@property NSInteger dropDownWidth;  // Returns or sets the width in pixels of the list for the specified command bar combobox control.  An error occurs if you attempt to set this property for a built-in combobox control.
@property NSInteger listIndex;

- (void) addItemToComboboxComboboxItem:(NSString *)comboboxItem entry_index:(NSInteger)entry_index;  // Add a new string to a custom combobox control.
- (void) clearCombobox;  // Clear all of the strings form a custom combobox.
- (NSString *) getComboboxItemEntry_index:(NSInteger)entry_index;  // Return the string at the given index within a combobox.
- (NSInteger) getCountOfComboboxItems;  // Return the number of strings within a combobox.
- (void) removeAnItemFromComboboxEntry_index:(NSInteger)entry_index;  // Remove a string from a custom combobox.
- (void) setComboboxItemEntry_index:(NSInteger)entry_index comboboxItem:(NSString *)comboboxItem;  // Set the string an a given index for a custom combobox.

@end

// Every command bar popup
@interface PowerPointCommandBarPopup : PowerPointCommandBarControl

- (SBElementArray *) commandBarControls;


@end

// Every command bar
@interface PowerPointCommandBar : PowerPointBaseObject

- (SBElementArray *) commandBarControls;

@property PowerPointMBPS barPosition;  // Returns or sets the position of the command bar.
@property (readonly) PowerPointMBTY barType;  // Returns the type of this command bar.
@property (readonly) BOOL builtIn;  // True if the command bar is built-in.
@property (copy, readonly) NSString *context;  // Returns or sets a string that determines where a command bar will be saved.
@property (readonly) BOOL embeddable;  // Returns if the command bar can be displayed inside the document window.
@property BOOL embedded;  // Returns or sets if the command bar will be displayed inside the document window.
@property BOOL enabled;  // Returns or set if the command bar is enabled.
@property (readonly) NSInteger entry_index;  // The index of the command bar.
@property NSInteger height;  // Returns or sets the height of the command bar.
@property NSInteger leftPosition;  // Returns or sets the left position of the command bar.
@property (copy) NSString *localName;  // Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar.
@property (copy) NSString *name;  // Returns or sets the name of the command bar.
@property (copy) NSArray *protection;  // Returns or sets the way a command bar is protected from user customization.  It accepts a list of the following items: no protection, no customize, no resize, no move, no change visible, no change dock, no vertical dock, no horizontal dock.
@property NSInteger rowIndex;  // Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero.
@property NSInteger top;  // Returns or sets the top position of a command bar.
@property BOOL visible;  // Returns or sets if the command bar is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar.


@end

// Every document property
@interface PowerPointDocumentProperty : PowerPointBaseObject

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

// Every custom document property
@interface PowerPointCustomDocumentProperty : PowerPointDocumentProperty


@end

@interface PowerPointWebPageFont : PowerPointBaseObject

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft PowerPoint Suite
 */

@interface PowerPointActionSetting : PowerPointBaseObject

@property PowerPointPAxT action;
@property (copy) NSString *actionSettingToRun;
@property (copy, readonly) PowerPointSoundEffect *actionSoundEffect;
@property (copy) NSString *actionVerb;
@property BOOL animateAction;
@property (copy, readonly) PowerPointHyperlink *hyperlink;
@property (copy) NSString *slideShowName;


@end

// Every animation behavior
@interface PowerPointAnimationBehavior : PowerPointBaseObject

@property PowerPointAnSm accumulate;
@property PowerPointAnpp additive;
@property PowerPointAnTy animationBehaviorType;
@property (copy, readonly) PowerPointColorsEffect *colorsEffect;
@property (copy, readonly) PowerPointCommandEffect *commandEffect;
@property (copy, readonly) PowerPointFilterEffect *filterEffect;
@property (copy, readonly) PowerPointMotionEffect *motionEffect;
@property (copy, readonly) PowerPointPropertyEffect *propertyEffect;
@property (copy, readonly) PowerPointRotatingEffect *rotatingEffect;
@property (copy, readonly) PowerPointScaleEffect *scaleEffect;
@property (copy, readonly) PowerPointSetEffect *setEffect;
@property (copy, readonly) PowerPointTiming *timing;


@end

// Every animation point
@interface PowerPointAnimationPoint : PowerPointBaseObject

@property (copy) NSString *formula;
@property double time;
@property (copy) SBObject *value;


@end

@interface PowerPointAnimationSettings : PowerPointBaseObject

@property double advanceTime;
@property PowerPointAsAe afterEffect;
@property BOOL animate;
@property BOOL animateBackground;
@property BOOL animateTextInReverse;
@property NSInteger animationOrder;
@property (copy, readonly) PowerPointPlaySettings *animationPlaySettings;
@property (copy, readonly) PowerPointSoundEffect *animationSoundEffect;
@property PowerPointPCuE chartUnitEffect;
@property (copy) NSColor *dimColor;
@property PowerPointMCoI dimColorThemeIndex;
@property PowerPointPEeF entryEffect;
@property PowerPointPTlE textLevelEffect;
@property PowerPointPTuE textUnitEffect;


@end

// An AppleScript object representing the Microsoft POWERPOINT application.
@interface PowerPointApplication : SBApplication

- (SBElementArray *) presentations;
- (SBElementArray *) documentWindows;
- (SBElementArray *) slideShowWindows;
- (SBElementArray *) commandBars;

@property (copy, readonly) NSString *Version;
@property (copy, readonly) PowerPointPresentation *activePresentation;
@property (copy, readonly) NSString *activePrinter;
@property (copy, readonly) PowerPointDocumentWindow *activeWindow;
@property (copy, readonly) PowerPointAutocorrect *autocorrectObject;  // Returns the autocorrect object
@property (copy, readonly) NSString *build;
@property (copy, readonly) NSString *caption;
@property PowerPointPSAT defaultSaveFormat;
@property (copy, readonly) PowerPointDefaultWebOptions *defaultWebOptionsObject;
@property NSInteger keyboardScript;  // Returns the current keyboard script
@property (copy, readonly) NSString *name;
@property (copy, readonly) NSString *operatingSystem;
@property (copy, readonly) NSString *path;
@property (copy, readonly) PowerPointPreferences *preferenceObject;
@property (copy, readonly) PowerPointSaveAsMovieSettings *saveAsMovieSettingsObject;
@property BOOL startUpDialog;

- (void) print:(NSArray *)x withProperties:(PowerPointPrintSettings *)withProperties;  // Print the specified object(s)
- (void) quitSaving:(PowerPointSavo)saving;  // Quit an application program
- (void) select;  // Make a selection
- (void) reset:(PowerPoint4000)x;  // Resets a built-in command bar or command bar control to its default configuration.
- (void) applyTheme:(PowerPoint4001)x fileName:(NSString *)fileName;  // Applies a theme or design template to the specified slide, master or presentation
- (void) arrangeWindows:(PowerPointPArS)x;  // Arrange Document Windows
- (void) insertTheText:(NSString *)theText at:(SBObject *)at;
- (void) pasteObject:(PowerPoint4002)x;
- (void) apply:(PowerPoint4003)x;
- (void) automaticLength:(PowerPoint4023)x;
- (void) beginConnect:(PowerPoint4018)x connectedShape:(PowerPointShape *)connectedShape connectionSite:(NSInteger)connectionSite;
- (void) beginDisconnect:(PowerPoint4019)x;
- (void) customDrop:(PowerPoint4024)x dropAmount:(double)dropAmount;
- (void) customLength:(PowerPoint4025)x length:(double)length;
- (void) endConnect:(PowerPoint4020)x connectedShape:(PowerPointShape *)connectedShape connectionSite:(NSInteger)connectionSite;
- (void) endDisconnect:(PowerPoint4021)x;
- (void) flip:(PowerPoint4007)x direction:(PowerPointMFlC)direction;
- (PowerPointActionSetting *) getActionSettingFor:(PowerPoint4009)x event:(PowerPointPMtv)event;
- (void) oneColorGradient:(PowerPoint4010)x style:(PowerPointMGdS)style variant:(NSInteger)variant degree:(double)degree;
- (void) patterned:(PowerPoint4011)x pattern:(PowerPointPpTy)pattern;
- (void) pickUp:(PowerPoint4004)x;
- (void) presetGradient:(PowerPoint4012)x style:(PowerPointMGdS)style variant:(NSInteger)variant gradientType:(PowerPointMPGb)gradientType;
- (void) presetTextured:(PowerPoint4013)x texture:(PowerPointMPzT)texture;
- (void) rerouteConnections:(PowerPoint4005)x;
- (void) resetRotation:(PowerPoint4022)x;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
- (void) setShapesDefaultProperties:(PowerPoint4006)x;
- (void) solid:(PowerPoint4014)x;
- (void) twoColorGradient:(PowerPoint4015)x style:(PowerPointMGdS)style variant:(NSInteger)variant;
- (void) userPicture:(PowerPoint4016)x pictureFile:(NSString *)pictureFile;
- (void) userTextured:(PowerPoint4017)x textureFile:(NSString *)textureFile;
- (void) zOrder:(PowerPoint4008)x zOrderPosition:(PowerPointMZoC)zOrderPosition;

@end

// Every autocorrect entry
@interface PowerPointAutocorrectEntry : PowerPointBaseObject

@property (copy, readonly) NSString *autocorrectValue;  // Returns the value of the auto correct entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the auto correct entry.


@end

// Represents the autocorrect functionality in PowerPoint.
@interface PowerPointAutocorrect : PowerPointBaseObject

- (SBElementArray *) autocorrectEntries;
- (SBElementArray *) firstLetterExceptions;
- (SBElementArray *) twoInitialCapsExceptions;

@property BOOL correctDays;  // Returns if PowerPoint automatically capitalizes the first letter of days of the week.
@property BOOL correctInitialCaps;  // Returns if PowerPoint automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, POwerPoint is corrected to PowerPoint.
@property BOOL correctSentenceCaps;  // Returns if PowerPoint automatically capitalizes the first letter in each sentence.
@property BOOL replaceText;  // Returns if Microsoft PowerPoint automatically replaces specified text with entries from the autocorrect list.


@end

@interface PowerPointBulletFormat : PowerPointBaseObject

@property (copy) NSString *bulletCharacter;  // Returns or sets the Unicode character value that is used for bullets in the specified text.
@property (copy, readonly) PowerPointFont *bulletFont;  // Returns a font object that represents character formatting for a bullet format object.
@property (readonly) NSInteger bulletNumber;  // Returns the bullet number of a paragraph.
@property NSInteger bulletStartValue;
@property PowerPointPBtS bulletStyle;  // Returns or sets a constant that represents the style of a bullet.
@property PowerPointPBtT bulletType;  // Returns or sets a constant that represents the type of bullet.
@property double relativeSize;  // Returns or sets the bullet size relative to the size of the first text character in the paragraph.
@property BOOL useTextColor;  // Determines whether the specified bullets are set to the color of the first text character in the paragraph.
@property BOOL useTextFont;  // Determines whether the specified bullets are set to the font of the first text character in the paragraph.
@property BOOL visible;  // Returns or sets a value that specifies whether the bullet is visible.

- (void) setBulletPicturePictureFile:(NSString *)pictureFile;  // Sets the graphics file to be used for bullets in a bulleted list.

@end

// Every color scheme
@interface PowerPointColorScheme : PowerPointBaseObject

- (NSColor *) getColorFromAt:(PowerPointPCSi)at;
- (void) setColorForAt:(PowerPointPCSi)at toColor:(NSColor *)toColor;

@end

@interface PowerPointColorsEffect : PowerPointBaseObject

@property (copy) NSColor *by_color;
@property PowerPointMCoI by_colorThemeIndex;  // Returns an object that represents a change to the color of the object by the specified number, expressed in RGB format.
@property (copy) NSColor *from_color;
@property PowerPointMCoI from_colorThemeIndex;  // Returns or sets an object that represents the starting RGB color value of an animation behavior.
@property (copy) NSColor *to_color;
@property PowerPointMCoI to_colorThemeIndex;  // Returns or sets an object that represents the RGB color value of an animation behavior.


@end

@interface PowerPointCommandEffect : PowerPointBaseObject

@property (copy) NSString *command;
@property PowerPointAnCT type;


@end

// Every custom layout
@interface PowerPointCustomLayout : PowerPointBaseObject

- (SBElementArray *) shapes;

@property (copy, readonly) PowerPointShape *background;
@property (copy, readonly) PowerPointDesign *design;
@property BOOL displayMasterShapes;
@property BOOL followMasterBackground;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property (readonly) double height;
@property (copy, readonly) PowerPointThemeColorScheme *themeColorScheme;  // Returns the color scheme of a custom layout.
@property (copy, readonly) PowerPointTimeline *timeline;
@property (readonly) double width;


@end

@interface PowerPointDefaultWebOptions : PowerPointBaseObject

@property BOOL allowPNG;
@property BOOL alwaysSaveInDefaultEncoding;
@property PowerPointPBuT buttonsType;
@property BOOL checkIfOfficeIsHTMLEditor;
@property PowerPointMtEn encoding;
@property PowerPointFClr frameColors;
@property BOOL includeBinaryFile;
@property PowerPointPNBp navBarPlacement;
@property BOOL supportIE4;
@property BOOL supportNN4;
@property BOOL supportOlderBrowsers;
@property BOOL updateLinksOnSave;
@property (copy) NSString *webPageKeywords;
@property (copy) NSString *webPageTitle;


@end

// Every design
@interface PowerPointDesign : PowerPointBaseObject

@property (copy, readonly) PowerPointMaster *slideMaster;


@end

// Every document window
@interface PowerPointDocumentWindow : PowerPointBasicWindow

- (SBElementArray *) panes;

@property (readonly) BOOL active;
@property (copy, readonly) PowerPointPane *activePane;
@property BOOL blackAndWhite;
@property (copy, readonly) NSString *caption;
@property (readonly) NSInteger entry_index;
@property double height;
@property double leftPosition;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointSelection *selection;  // Represents the selection in the specified document window.
@property NSInteger splitHorizontal;
@property NSInteger splitVertical;
@property double top;
@property (copy, readonly) PowerPointView *view;
@property PowerPointPVTy viewType;
@property double width;

- (void) launchSpellerOn;

@end

@interface PowerPointEffectInformation : PowerPointBaseObject

@property (readonly) PowerPointAnAE afterEffectInformation;
@property (readonly) BOOL animateBackgroundInformation;
@property (readonly) BOOL animateTextInReverseInformation;
@property (readonly) PowerPointAnLv buildByLevel;
@property (copy, readonly) NSColor *dim;
@property (copy, readonly) PowerPointPlaySettings *playSettingsInformation;
@property (copy, readonly) PowerPointSoundEffect *soundEffectInformation;
@property (readonly) PowerPointAnTU textUnitEffectInformation;


@end

@interface PowerPointEffectParameters : PowerPointBaseObject

@property double amount;
@property (copy, readonly) NSColor *color2;
@property (readonly) PowerPointMCoI color2ColorThemeIndex;  // Returns an object that represents the color on which to end a color-cycle animation.
@property PowerPointAnDi direction;
@property (copy) NSString *font2;
@property BOOL relative;
@property double size2;


@end

// Every effect
@interface PowerPointEffect : PowerPointBaseObject

- (SBElementArray *) animationBehaviors;

@property PowerPointAnFX animationEffectType;
@property (copy, readonly) PowerPointEffectInformation *effectInformation;
@property (copy, readonly) PowerPointEffectParameters *effectParameters;
@property BOOL exitAnimation;
@property (copy, readonly) NSString *name;
@property (copy) PowerPointShape *shape;
@property NSInteger targetParagraph;
@property (readonly) NSInteger textRangeLength;
@property (readonly) NSInteger textRangeStart;
@property (copy, readonly) PowerPointTiming *timing;

- (PowerPointAnimationBehavior *) addBehaviorType:(PowerPointAnTy)type;  // add an animation behavior

@end

@interface PowerPointFilterEffect : PowerPointBaseObject

@property PowerPointAfet filterType;
@property BOOL reveal;
@property PowerPointAfes subtype;


@end

// Represents an abbreviation excluded from automatic correction.
@interface PowerPointFirstLetterException : PowerPointBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the FirstLetterException.


@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface PowerPointFont : PowerPointBaseObject

@property (copy) NSString *ASCIIName;  // Returns or sets the font used for Latin text; characters with character codes from 0 through 127.
@property BOOL autoRotateNumbers;  // Returns or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.
@property double baseLineOffset;  // Returns or sets a value specifying the horizontaol offset of the selected font.
@property BOOL bold;  // Returns or sets a value specifying whether the font should be bold.
@property PowerPointMTCa capsType;  // Returns or sets a value specifying how the text should be capitalized.
@property (copy) NSString *eastAsianName;  // Returns or sets the font name used for Asian text.
@property (readonly) BOOL embedable;  // Returns a value indicating whether the font can be embedded in a page.
@property (readonly) BOOL embedded;  // Returns a value specifying whether the font is embedded in a page.
@property BOOL emboss;
@property BOOL equalizeCharacterHeight;  // Returns or sets a value specifying whether the text should have the same horizontal height.
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns a value specifying the fill formatting for text.
@property (copy) NSColor *fontColor;
@property PowerPointMCoI fontColorThemeIndex;  // Returns or sets the color for the specified font.
@property (copy) NSString *fontName;  // Returns or sets a value specifying the font to use for a selection.
@property (copy) NSString *fontNameOther;  // Returns or sets the font used for characters whose character set numbers are greater than 127.
@property double fontSize;
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns a value specifying the glow formatting of the text.
@property (copy) NSColor *highlightColor;  // Returns or sets the text highlight color for object.
@property PowerPointMCoI highlightColorThemeIndex;  // Returns or sets the specified text highlight color for object.
@property BOOL italic;
@property double kerning;  // Returns or sets a value specifying the amount of spacing between text characters.
@property (copy, readonly) PowerPointLineFormat *lineFormat;  // Returns a value specifiying the line formatting of the text.
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns a value specifying the reflection formatting of the text.
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;  // Returns the value specifying the type of shadow effect for the selection of text.
@property PowerPointMSET softEdgeFormat;  // Returns or sets a value specifying the soft edge fromatting of the text.
@property double spacing;  // Returns or sets a value specifying the spacing between characters in a selection of text.
@property PowerPointMTSt strikeType;  // Returns or sets a value specifying the strike format used for a selection of text.
@property BOOL subscript;  // Returns or sets a value specifying that the selected text should be displayed a subscript.
@property BOOL superscript;  // Returns or sets a value specifying that the selected text should be displayed a superscript.
@property double transparency;
@property BOOL underline;
@property (copy) NSColor *underlineColor;  // Returns or sets the 24-bit color of the underline for the specified font object.
@property PowerPointMCoI underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.
@property PowerPointMTUl underlineStyle;  // Returns or sets a value specifying the underline style for the selected text.
@property PowerPointMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.


@end

@interface PowerPointHeaderOrFooter : PowerPointBaseObject

@property PowerPointPDtF dateFormat;
@property (copy) NSString *headerFooterText;
@property BOOL useDateFormat;
@property BOOL visible;


@end

@interface PowerPointHeadersAndFooters : PowerPointBaseObject

@property (copy, readonly) PowerPointHeaderOrFooter *dateAndTime;
@property BOOL displayHeadersFootersOnTitleSlide;
@property (copy, readonly) PowerPointHeaderOrFooter *footer;
@property (copy, readonly) PowerPointHeaderOrFooter *header;
@property (copy, readonly) PowerPointHeaderOrFooter *slideNumber;


@end

// Every hyperlink
@interface PowerPointHyperlink : PowerPointBaseObject

@property (copy) NSString *hyperlinkAddress;
@property (copy) NSString *hyperlinkSubAddress;
@property (readonly) PowerPointMHyT hyperlinkType;


@end

@interface PowerPointMaster : PowerPointBaseObject

- (SBElementArray *) shapes;
- (SBElementArray *) hyperlinks;
- (SBElementArray *) customLayouts;

@property (copy, readonly) PowerPointShape *background;
@property (copy, readonly) PowerPointColorScheme *colorScheme;
@property (copy, readonly) PowerPointDesign *design;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property (readonly) double height;
@property (copy, readonly) PowerPointOfficeTheme *theme;
@property (copy, readonly) PowerPointTimeline *timeline;
@property (readonly) double width;

- (PowerPointTextStyle *) getTextStyleFromAt:(PowerPointPTst)at;

@end

@interface PowerPointMotionEffect : PowerPointBaseObject

@property double byX;
@property double byY;
@property double fromX;
@property double fromY;
@property (copy) NSString *path;
@property double toX;
@property double toY;


@end

// Every named slide show
@interface PowerPointNamedSlideShow : PowerPointBaseObject

@property (copy, readonly) NSString *name;
@property (readonly) NSInteger numberOfSlides;
@property (copy, readonly) NSArray *slideIDs;


@end

@interface PowerPointPageSetup : PowerPointBaseObject

@property NSInteger firstSlideNumber;
@property PowerPointMOrT notesOrientation;
@property PowerPointMOrT slideOrientation;
@property PowerPointSSzT slideSize;
@property double slideWidth;


@end

// Every pane
@interface PowerPointPane : PowerPointBaseObject

@property (readonly) BOOL active;
@property (readonly) PowerPointPVTy paneViewType;


@end

@interface PowerPointParagraphFormat : PowerPointBaseObject

- (SBElementArray *) tabStops;

@property PowerPointPpgA alignment;
@property PowerPointPBlA baselineAlignment;  // Returns or sets a constant that represents the vertical position of fonts in a paragraph.
@property (copy, readonly) PowerPointBulletFormat *bulletFormat;
@property BOOL eastAsianLineBreakControl;
@property double firstLineIndent;  // Returns or sets the value, in points, for a first line or hanging indent.
@property BOOL hangingPunctuation;  // Determines whether hanging punctuation is enabled for the specified paragraphs.
@property NSInteger indentLevel;  // Returns or sets a value representing the indent level assigned to text in the selected paragraph.
@property double leftIndent;  // Returns or sets a value that represents the left indent value, in points, for the specified paragraphs.
@property BOOL lineRuleAfter;  // Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleBefore;  // Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleWithin;  // Determines whether line spacing between base lines is set to a specific number of points or lines.
@property double rightIndent;  // Returns or sets the right indent, in points, for the specified paragraphs.
@property double spaceAfter;  // Returns or sets the spacing, in points, after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing, in points, before the specified paragraphs.
@property double spaceWithin;  // Returns or sets the amount of space between base lines in the specified paragraph, in points or lines.
@property PowerPointPDrT textDirection;  // Returns or sets the text direction for the specified paragraph.
@property BOOL wordWrap;  // Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.


@end

@interface PowerPointPlaySettings : PowerPointBaseObject

@property BOOL hideWhileNotPlaying;
@property BOOL loopUntilStopped;
@property BOOL pauseAnimation;
@property BOOL playOnEntry;
@property BOOL rewindMove;
@property NSInteger stopAfterSlides;


@end

@interface PowerPointPreferences : PowerPointBaseObject

@property NSInteger alwaysSuggestCorrections;
@property NSInteger appendDOSExtentions;
@property NSInteger autoFit;
@property NSInteger autoRecoveryFrequency;
@property NSInteger backgroundSpelling;
@property NSInteger compressFile;
@property NSInteger defaultView;
@property NSInteger dragDrop;
@property NSInteger endBlankSlide;
@property NSInteger filePropertiesPrompt;
@property NSInteger fullTextSearch;
@property NSInteger hideSpellingErrors;
@property NSInteger ignoreNumbersInWords;
@property NSInteger ignoreUppercase;
@property NSInteger linkSoundLimit;
@property NSInteger mruListActive;
@property NSInteger mruListSize;
@property NSInteger optionBitmap;
@property NSInteger rulerUnits;
@property NSInteger saveAutoRecovery;
@property NSInteger showVerticalRuler;
@property NSInteger slideShowControl;
@property NSInteger slideShowRightMouse;
@property NSInteger smartCutPaste;
@property NSInteger smartQuotes;
@property NSInteger undoLevelCount;
@property (copy) NSString *userInitials;
@property (copy) NSString *userName;
@property NSInteger wordSelection;


@end

// Every presentation
@interface PowerPointPresentation : PowerPointBaseObject

- (SBElementArray *) slides;
- (SBElementArray *) colorSchemes;
- (SBElementArray *) fonts;
- (SBElementArray *) documentWindows;
- (SBElementArray *) documentProperties;
- (SBElementArray *) customDocumentProperties;
- (SBElementArray *) designs;

@property (copy, readonly) PowerPointShape *defaultShape;
@property PowerPointPeBl eastAsianLineBreakLevel;  // Returns or sets the East Asian line break control level for the specified paragraph.
@property (copy, readonly) NSString *fullName;
@property (copy, readonly) PowerPointMaster *handoutMaster;
@property (readonly) BOOL hasTitleMaster;
@property PowerPointPDrT layoutDirection;
@property (copy, readonly) NSString *name;
@property (copy) NSString *noLineBreakAfter;
@property (copy) NSString *noLineBreakBefore;
@property (copy, readonly) PowerPointMaster *notesMaster;
@property (copy, readonly) PowerPointPageSetup *pageSetup;
@property (copy, readonly) NSString *path;
@property (copy, readonly) PowerPointPrintOptions *printOptions;
@property (readonly) BOOL readOnly;
@property (copy, readonly) PowerPointSaveAsMovieSettings *saveAsMovieSettings;
@property BOOL saved;
@property (copy, readonly) PowerPointMaster *slideMaster;
@property (copy, readonly) PowerPointSlideShowSettings *slideShowSettings;
@property (copy, readonly) PowerPointSlideShowWindow *slideShowWindow;
@property (copy, readonly) NSString *templateName;
@property (copy, readonly) PowerPointMaster *titleMaster;
@property (copy, readonly) PowerPointWebOptions *webOptions;

- (PowerPointDesign *) addDesignDesignName:(NSString *)DesignName index:(NSInteger)index;  // add a new design
- (void) applyTemplateFileName:(NSString *)fileName;  // Applies a theme or design template to the specified slide, master or presentation
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to printToFile:(NSString *)printToFile copies:(NSInteger)copies collate:(BOOL)collate showDialog:(BOOL)showDialog;
- (void) redoTimes:(NSInteger)times;
- (void) undoTimes:(NSInteger)times;
- (void) updateLinks;

@end

@interface PowerPointPresenterTool : PowerPointBaseObject

@property (copy, readonly) PowerPointSlide *currentPresenterSlide;
@property (copy) NSString *notesText;
@property NSInteger notesZoom;
@property (readonly) BOOL slideMiniature;

- (void) endShow;
- (void) next;
- (void) previous;
- (void) toggleSlideMiniature;

@end

// Every presenter view window
@interface PowerPointPresenterViewWindow : PowerPointBasicWindow

@property (readonly) BOOL active;
@property (readonly) double height;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointPresenterTool *presenterTool;
@property (readonly) double width;


@end

@interface PowerPointPrintOptions : PowerPointBaseObject

- (SBElementArray *) printRanges;

@property (copy, readonly) NSString *activePrinter;
@property BOOL collate;
@property BOOL fitToPage;
@property BOOL frameSlides;
@property NSInteger numberOfCopies;
@property PowerPointPtOt outputType;
@property PowerPointPrCt printColorType;
@property BOOL printFontsAsGraphics;
@property BOOL printHiddenSlides;
@property PowerPointRgTy rangeType;
@property (copy) NSString *slideShowName;


@end

// Every print range
@interface PowerPointPrintRange : PowerPointBaseObject

@property (readonly) NSInteger rangeEnd;
@property (readonly) NSInteger rangeStart;


@end

@interface PowerPointPropertyEffect : PowerPointBaseObject

- (SBElementArray *) animationPoints;

@property (copy, readonly) id ending;
@property PowerPointAnPr propertySetEffect;
@property (copy, readonly) id starting;


@end

@interface PowerPointRotatingEffect : PowerPointBaseObject

@property double rotating;


@end

// Every ruler level
@interface PowerPointRulerLevel : PowerPointBaseObject

@property double firstMargin;  // Returns or sets the first-line indent for the specified outline level, in points.
@property double leftMargin;  // Returns or sets the left indent for the specified outline level, in points.


@end

// Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.
@interface PowerPointRuler : PowerPointBaseObject

- (SBElementArray *) tabStops;
- (SBElementArray *) rulerLevels;


@end

@interface PowerPointSaveAsMovieSettings : PowerPointBaseObject

@property BOOL autoLoopEnabled;
@property (copy) NSString *backgroundSoundTrackFile;
@property NSInteger backgroundTrackSegmentEnd;
@property NSInteger backgroundTrackSegmentStart;
@property NSInteger backgroundTrackStart;
@property BOOL createMoviePreview;
@property BOOL forceAllInline;
@property BOOL includeNarrationAndSounds;
@property (copy) NSString *movieActors;
@property (copy) NSString *movieAuthor;
@property (copy) NSString *movieCopyright;
@property NSInteger movieFrameHeight;
@property NSInteger movieFrameWidth;
@property (copy) NSString *movieProducer;
@property PowerPointPMOt optimization;
@property BOOL showMovieController;
@property (copy) NSString *transitionDescription;
@property BOOL useSingleTransition;


@end

@interface PowerPointScaleEffect : PowerPointBaseObject

@property double byX;
@property double byY;
@property double fromX;
@property double fromY;
@property double toX;
@property double toY;


@end

// Represents the selection in the specified document window
@interface PowerPointSelection : PowerPointBaseObject

@property (copy, readonly) PowerPointShapeRange *childShapeRange;  // Returns the child shapes of a selection.
@property (readonly) BOOL hasChildShapeRange;  // Returns whether the selection contains child shapes.
@property (readonly) PowerPointPSEL selectionType;  // Represents the type of objects in a selection.
@property (copy, readonly) PowerPointShapeRange *shapeRange;  // Returns a collection of shapes selected on the specified slide.
@property (copy, readonly) PowerPointSlideRange *slideRange;  // Returns a collection of selected slides.
@property (copy, readonly) PowerPointTextRange *textRange;  // Returns the text and text properties of the selected text.

- (void) unselect;  // Cancels the current selection.

@end

// Every sequence
@interface PowerPointSequence : PowerPointBaseObject

- (SBElementArray *) effects;

- (PowerPointEffect *) addEffectFor:(PowerPointShape *)for_ fx:(PowerPointAnFX)fx level:(PowerPointAnLv)level trigger:(PowerPointAnTr)trigger index:(NSInteger)index;  // add an effect for a shape
- (PowerPointEffect *) convertToTextUnitEffectEffect:(PowerPointEffect *)Effect unit:(PowerPointAnTU)unit;  // convert an effect to a text unit effect

@end

@interface PowerPointSetEffect : PowerPointBaseObject

@property (copy) id ending;
@property PowerPointAnPr propertySetEffect;


@end

// A collection that represents a notes page or a slide range, which is a set of slides that can contain a single slide to as many as all slides in a presentation.
@interface PowerPointSlideRange : PowerPointBaseObject

- (SBElementArray *) slides;
- (SBElementArray *) shapes;
- (SBElementArray *) hyperlinks;

@property (copy) PowerPointColorScheme *colorScheme;  // Returns or sets the color scheme for the specified slide, slide range, or slide master.
@property (copy) PowerPointCustomLayout *customLayout;  // Returns the custom layout associated with the specified range of slides.
@property (copy) PowerPointDesign *design;
@property BOOL displayMasterShapes;  // Determines whether the specified range of slides displays the background objects on the slide master.
@property BOOL followMasterBackground;  // Determines whether the range of slides follows the slide master background.
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;  // Returns a collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides.
@property PowerPointPSLo layout;  // Returns or sets the slide layout.
@property (copy, readonly) PowerPointSlideRange *notesPage;  // Returns a slide range object that represents the notes pages for the specified slide or range of slides.
@property (readonly) NSInteger printSteps;
@property (readonly) NSInteger slideID;
@property (readonly) NSInteger slideIndex;
@property (copy, readonly) PowerPointMaster *slideMaster;  // Returns the slide master.
@property (readonly) NSInteger slideNumber;  // Returns the slide number.
@property (copy, readonly) PowerPointSlideShowTransition *slideShowTransitions;  // Returns the special effects for the specified slide transition.
@property (copy, readonly) PowerPointTimeline *timeline;  // Returns the animation timeline for the slide.


@end

@interface PowerPointSlideShowSettings : PowerPointBaseObject

- (SBElementArray *) namedSlideShows;

@property PowerPointPSaM advanceMode;
@property NSInteger endingSlide;
@property BOOL loopUntilStopped;
@property (copy) NSColor *pointerColor;
@property PowerPointMCoI pointerColorThemeIndex;  // Returns or sets the color of  default pointer color for a presentation.
@property PowerPointSRgT rangeType;
@property PowerPointPSSt showType;
@property (readonly) BOOL showWithAnimation;
@property BOOL showWithNarration;
@property (copy) NSString *slideShowName;
@property NSInteger startingSlide;

- (PowerPointSlideShowWindow *) runSlideShow;

@end

@interface PowerPointSlideShowTransition : PowerPointBaseObject

@property BOOL advanceOnClick;
@property BOOL advanceOnTime;
@property double advanceTime;
@property PowerPointPEeF entryEffect;
@property BOOL hidden;
@property BOOL loopSoundUntilNext;
@property (copy, readonly) PowerPointSoundEffect *soundEffectTransition;


@end

@interface PowerPointSlideShowView : PowerPointBaseObject

@property BOOL accelerationsEnabled;
@property (readonly) NSInteger currentShowPosition;
@property (readonly) BOOL isNamedShow;
@property (copy, readonly) PowerPointSlide *lastSlideViewed;
@property (copy) NSColor *pointerColor;
@property PowerPointMCoI pointerColorThemeIndex;  // Returns or sets the color of temporary pointer color for a view of a slide show.
@property PowerPointPSsP pointerType;
@property (readonly) double presentationElapsedTime;
@property double slideElapsedTime;
@property (copy, readonly) NSString *slideShowName;
@property PowerPointPShS slideState;
@property (readonly) NSInteger zoom;

- (void) exitSlideShow;
- (void) goToFirstSlide;
- (void) goToLastSlide;
- (void) goToNextSlide;
- (void) goToPreviousSlide;
- (void) resetSlideTime;

@end

// Every slide show window
@interface PowerPointSlideShowWindow : PowerPointBasicWindow

@property (readonly) BOOL active;
@property NSRect bounds;
@property double height;
@property (readonly) BOOL isFullScreen;
@property double leftPosition;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointSlideShowView *slideshowView;
@property double top;
@property double width;


@end

// Every slide
@interface PowerPointSlide : PowerPointBaseObject

- (SBElementArray *) shapes;
- (SBElementArray *) hyperlinks;

@property (copy, readonly) PowerPointShape *background;
@property (copy) PowerPointColorScheme *colorScheme;
@property (copy) PowerPointCustomLayout *customLayout;
@property (copy) PowerPointDesign *design;
@property BOOL displayMasterShapes;
@property BOOL followMasterBackground;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property PowerPointPSLo layout;
@property (copy, readonly) PowerPointSlide *notesPage;
@property (readonly) NSInteger printSteps;
@property (readonly) NSInteger slideID;
@property (readonly) NSInteger slideIndex;
@property (copy, readonly) PowerPointMaster *slideMaster;
@property (readonly) NSInteger slideNumber;
@property (copy, readonly) PowerPointSlideShowTransition *slideShowTransition;
@property (copy, readonly) PowerPointTimeline *timeline;

- (void) applyThemeColorSchemeFileName:(NSString *)fileName;  // Applies a theme color to the specified slide.
- (void) copyObject;
- (void) cutObject;

@end

@interface PowerPointSoundEffect : PowerPointBaseObject

@property (copy, readonly) NSString *name;
@property PowerPointPSnX soundType;

- (void) importSoundFileSoundFileName:(NSString *)soundFileName;
- (void) playSoundEffect;

@end

// Every tab stop
@interface PowerPointTabStop : PowerPointBaseObject

@property double tabPosition;  // Returns or sets the position of a tab stop relative to the left margin.
@property PowerPointPTSp tabStopType;  // Returns or sets the type of the tab stop object.


@end

// Every text style level
@interface PowerPointTextStyleLevel : PowerPointBaseObject

@property (copy, readonly) PowerPointFont *font;
@property (copy, readonly) PowerPointParagraphFormat *paragraphFormat;


@end

@interface PowerPointTextStyle : PowerPointBaseObject

- (SBElementArray *) textStyleLevels;

@property (copy, readonly) PowerPointRuler *ruler;
@property (copy, readonly) PowerPointTextFrame *textFrame;


@end

@interface PowerPointTimeline : PowerPointBaseObject

- (SBElementArray *) sequences;

@property (copy, readonly) PowerPointSequence *mainSequence;

- (PowerPointSequence *) addSequenceIndex:(NSInteger)index;  // add an interactive sequence

@end

@interface PowerPointTiming : PowerPointBaseObject

@property double acceleration;
@property BOOL autoreverse;
@property double deceleration;
@property double duration;
@property NSInteger repeatCount;
@property double repeatDuration;
@property PowerPointAnRs restart;
@property BOOL rewind;
@property BOOL smoothEnd;
@property BOOL smoothStart;
@property double speed;


@end

// Represents a single initial-capital autocorrect exception.
@interface PowerPointTwoInitialCapsException : PowerPointBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the TwoInitialCapsException.


@end

@interface PowerPointView : PowerPointBaseObject

@property BOOL displaySlideMiniature;
@property (copy) PowerPointSlide *slide;
@property (readonly) PowerPointPVTy viewType;
@property NSInteger zoom;
@property BOOL zoomToFit;

- (void) goToSlideNumber:(NSInteger)number;

@end

@interface PowerPointWebOptions : PowerPointBaseObject

@property BOOL allowPNG;
@property PowerPointPBuT buttonsType;
@property PowerPointMtEn encoding;
@property PowerPointFClr frameColors;
@property BOOL includeBinaryFile;
@property PowerPointPNBp navBarPlacement;
@property PowerPointPALO pageLayout;
@property BOOL supportIE4;
@property BOOL supportNN4;
@property BOOL supportOlderBrowsers;
@property (copy) NSString *webPageKeywords;
@property (copy) NSString *webPageTitle;


@end



/*
 * Drawing Suite
 */

// Every adjustment
@interface PowerPointAdjustment : PowerPointBaseObject

@property double adjustment_value;


@end

@interface PowerPointCalloutFormat : PowerPointBaseObject

@property BOOL accent;
@property PowerPointMCAt angle;
@property BOOL autoAttach;
@property (readonly) BOOL autoLength;
@property BOOL border;
@property (readonly) double calloutFormatLength;
@property PowerPointMCot calloutType;
@property (readonly) double drop;
@property (readonly) PowerPointMCDt dropType;
@property double gap;

- (void) presetDropDropType:(PowerPointMCDt)dropType;

@end

@interface PowerPointConnectorFormat : PowerPointBaseObject

@property (readonly) BOOL beginConnected;
@property (copy, readonly) PowerPointShape *beginConnectedShape;
@property (readonly) NSInteger beginConnectionSite;
@property PowerPointMCtT connectorType;
@property (readonly) BOOL endConnected;
@property (copy, readonly) PowerPointShape *endConnectedShape;
@property (readonly) NSInteger endConnectionSite;


@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface PowerPointFillFormat : PowerPointBaseObject

- (SBElementArray *) gradientStops;

@property (copy) NSColor *backColor;
@property PowerPointMCoI backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) PowerPointMFdT fillFormatType;
@property (copy) NSColor *foreColor;
@property PowerPointMCoI foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) PowerPointMGCt gradientColorType;
@property (readonly) double gradientDegree;
@property (readonly) PowerPointMGdS gradientStyle;
@property (readonly) NSInteger gradientVariant;
@property (readonly) PowerPointPpTy pattern;
@property (readonly) PowerPointMPGb presetGradientType;
@property (readonly) PowerPointMPzT presetTexture;
@property BOOL rotateWithObject;  // Returns or sets whether the fill rotates with the specified shape.
@property PowerPointMTtA textureAlignment;  // Returns or sets the texture alignment for the specified object.
@property double textureHorizontalScale;  // Returns or sets the texture alignment for the specified object.
@property (copy, readonly) NSString *textureName;
@property double textureOffsetX;  // Returns or sets the texture alignment for the specified object.
@property double textureOffsetY;  // Returns or sets the texture alignment for the specified object.
@property BOOL textureTile;  // Returns the texture tile style for the specified fill.
@property double textureVerticalScale;  // Returns or sets the texture alignment for the specified object.
@property double transparency;
@property BOOL visible;

- (void) deleteGradientStopStopIndex:(NSInteger)stopIndex;  // Removes a gradient stop.
- (void) insertGradientStopCustomColor:(NSColor *)customColor position:(double)position transparency:(double)transparency stopIndex:(NSInteger)stopIndex;  // Adds a stop to a gradient.

@end

// Represents the glow formatting for a shape or range of shapes
@interface PowerPointGlowFormat : PowerPointBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified glow format.
@property PowerPointMCoI colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Every gradient stop
@interface PowerPointGradientStop : PowerPointBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified the gradient stop.
@property PowerPointMCoI colorThemeIndex;  // Returns or sets the color for the specified gradient stop.
@property double position;  // Returns or sets the position for the specified gradient stop expressed as a percent.
@property double transparency;  // Returns or sets a value representing the transparency of the gradient fill expressed as a percent.


@end

@interface PowerPointLineFormat : PowerPointBaseObject

@property (copy) NSColor *backColor;
@property PowerPointMCoI backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property PowerPointMAhL beginArrowHeadLength;
@property PowerPointMAhS beginArrowheadStyle;
@property PowerPointMAhW beginArrowheadWidth;
@property PowerPointMlDs dashStyle;
@property PowerPointMAhL endArrowheadLength;
@property PowerPointMAhS endArrowheadStyle;
@property PowerPointMAhW endArrowheadWidth;
@property (copy) NSColor *foreColor;
@property PowerPointMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property PowerPointPpTy lineFormatPatterned;
@property PowerPointMLnS lineStyle;
@property double lineWeight;
@property double transparency;


@end

@interface PowerPointLinkFormat : PowerPointBaseObject

@property PowerPointPUdO autoUpdate;
@property (copy) NSString *sourceFullName;


@end

// Represents a Microsoft Office theme.
@interface PowerPointOfficeTheme : PowerPointBaseObject

@property (copy, readonly) PowerPointThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) PowerPointThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) PowerPointThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

@interface PowerPointPictureFormat : PowerPointBaseObject

@property double brightness;
@property PowerPointMPc colorType;
@property double contrast;
@property double cropBottom;
@property double cropLeft;
@property double cropRight;
@property double cropTop;
@property (copy) NSColor *transparencyColor;
@property BOOL transparentBackground;


@end

@interface PowerPointPlaceholderFormat : PowerPointBaseObject

@property (readonly) PowerPointMShp containedType;
@property (copy) NSString *placeholderName;
@property (readonly) PowerPointPPhd placeholderType;


@end

// Represents the reflection effect in Office graphics.
@interface PowerPointReflectionFormat : PowerPointBaseObject

@property PowerPointMRfT reflectionType;  // Returns or sets the type of the reflection format object.


@end

@interface PowerPointShadowFormat : PowerPointBaseObject

@property double blur;
@property (copy) NSColor *foreColor;
@property PowerPointMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;
@property double offsetX;
@property double offsetY;
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property PowerPointMSSt shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property PowerPointMSdT shadowType;
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;
@property BOOL visible;


@end

// A collection that represents a set of shapes on a document.
@interface PowerPointShapeRange : PowerPointBaseObject

- (SBElementArray *) adjustments;
- (SBElementArray *) shapes;

@property (copy, readonly) PowerPointAnimationSettings *animationSettings;  // Returns all the special effects that you can apply to the animation of the specified shape.
@property PowerPointMAsT autoShapeType;  // Returns or sets the shape type for AutoShapes other than a line, freeform drawing, or connector.
@property PowerPointMBSI backgroundStyle;  // Returns or sets the background style for the specified object.
@property PowerPointMBW blackAndWhiteMode;  // Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode.
@property (copy, readonly) PowerPointCalloutFormat *calloutFormat;  // Returns callout formatting properties for the specified line callout shapes.
@property (readonly) BOOL child;  // Returns whether all shapes in a shape range are child shapes of the same parent.
@property (readonly) NSInteger connectionSiteCount;  // Returns the number of connection sites on the specified shape.
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns the fill format properties for the specified object.
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns the glow format for the specified range of shapes.
@property (readonly) BOOL hasChild;
@property (readonly) BOOL hasTable;  // Returns whether the specified shape is a table.
@property (readonly) BOOL hasTextFrame;  // Returns whether the specified shape has a text frame.
@property double height;  // Returns or sets the height of the specified object.
@property (readonly) BOOL horizontalFlip;  // Returns whether the specified shape is flipped around the horizontal axis.
@property (readonly) BOOL isConnector;  // Determines whether the specified shape is a connector.
@property double leftPosition;  // Returns or sets a value equal to the distance from the left edge of the leftmost shape in the shape range to the left edge of the slide.
@property (copy, readonly) PowerPointLineFormat *lineFormat;  // Returns line format properties for the specified line or shape border.
@property (copy, readonly) PowerPointLinkFormat *linkFormat;  // Returns the properties for linked OLE objects.
@property BOOL lockAspectRatio;  // Determines whether the specified shape retains its original proportions when you resize it.
@property (readonly) PowerPointMedT mediaType;  // Returns the OLE media type.
@property (copy) NSString *name;  // Identifies the shape on a slide.
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns the reflection format for the specified range of shapes.
@property double rotation;  // Returns or sets the number of degrees that the specified shape is rotated around the z-axis.
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;  // Returns shadow formatting for specified shapes.
@property PowerPointMSSI shapeStyle;  // Returns or sets the shape style index for the specified object.
@property (readonly) PowerPointMShp shapeType;  // Represents the type of shape or shapes in a range of shapes.
@property (copy, readonly) PowerPointSoftEdgeFormat *softEdgeFormat;  // Returns the soft edge format for the specified range of shapes.
@property (copy, readonly) PowerPointTextFrame *textFrame;  // Returns the alignment and anchoring properties for the specified shape or master text style.
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns 3-D effect formatting properties for the specified shape.
@property double top;  // Returns or sets a value equal to the distance from the top edge of the topmost shape in the shape range to the top edge of the document.
@property (readonly) BOOL verticalFlip;  // Determines whether the specified shape is flipped around the vertical axis.
@property BOOL visible;  // Returns or sets the visibility of the specified object or the formatting applied to the specified object.
@property double width;  // Returns or sets the width of the specified object.
@property (readonly) NSInteger zOrderPosition;  // Returns the position of the specified shape in the z-order.

- (void) alignAlignType:(PowerPointMAlC)alignType relativeToSlide:(BOOL)relativeToSlide;  // Aligns the shapes in the specified range of shapes.
- (void) copyShapeRange;
- (void) cutShapeRange;
- (void) distributeDistributeType:(PowerPointMDsC)distributeType relativeToSlide:(BOOL)relativeToSlide;  // Evenly distributes the shapes in the specified range of shapes. You can specify whether you want to distribute the shapes horizontally or vertically and whether you want to distribute them over the entire slide or just over the space they originally occup
- (PowerPointShape *) group;  // Groups the shapes in the specified shape range.
- (PowerPointShape *) regroup;  // Regroups the group that the specified shape range belonged to previously.
- (PowerPointShapeRange *) ungroup;  // Ungroups any grouped shapes in the specified shape or range of shapes.

@end

// Every shape
@interface PowerPointShape : PowerPointBaseObject

- (SBElementArray *) shapes;
- (SBElementArray *) callouts;
- (SBElementArray *) connectors;
- (SBElementArray *) pictures;
- (SBElementArray *) lineShapes;
- (SBElementArray *) placeHolders;
- (SBElementArray *) textBoxes;
- (SBElementArray *) comments;
- (SBElementArray *) shapeTables;
- (SBElementArray *) adjustments;

@property (copy, readonly) PowerPointAnimationSettings *animationSettings;
@property PowerPointMAsT autoShapeType;
@property PowerPointMBSI backgroundStyle;  // Returns or sets the background style.
@property PowerPointMBW blackAndWhiteMode;
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger connectionSiteCount;
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified shape.
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (readonly) BOOL hasTable;
@property (readonly) BOOL hasTextFrame;
@property double height;
@property (readonly) BOOL horizontalFlip;
@property (readonly) BOOL isConnector;
@property double leftPosition;
@property (copy, readonly) PowerPointLineFormat *lineFormat;
@property (copy, readonly) PowerPointLinkFormat *linkFormat;
@property BOOL lockAspectRatio;
@property (readonly) PowerPointMedT mediaType;
@property (copy) NSString *name;
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property double rotation;
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;
@property PowerPointMSSI shapeStyle;  // Returns or sets the shape style corresponding to the Quick Styles.
@property (readonly) PowerPointMShp shapeType;
@property (copy, readonly) PowerPointSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) PowerPointTextFrame *textFrame;
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape.
@property double top;
@property (readonly) BOOL verticalFlip;
@property BOOL visible;
@property double width;
@property (copy, readonly) PowerPointWordArtFormat *wordArtFormat;  // Returns the formatting properties for a word art effect.
@property (readonly) NSInteger zOrderPosition;

- (void) copyShape;
- (void) cutShape;
- (void) saveAsPicturePictureType:(PowerPointMPiT)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format

@end

// Every callout
@interface PowerPointCallout : PowerPointShape

@property (readonly) PowerPointMCot calloutType;
@property (copy, readonly) PowerPointCalloutFormat *calloutFormat;


@end

// Every comment
@interface PowerPointComment : PowerPointShape


@end

// Every connector
@interface PowerPointConnector : PowerPointShape

@property (copy, readonly) PowerPointConnectorFormat *connectorFormat;
@property (readonly) PowerPointMCtT connectorType;


@end

// Every line shape
@interface PowerPointLineShape : PowerPointShape

@property double beginLineX;
@property double beginLineY;
@property double endLineX;
@property double endLineY;


@end

// Every media object
@interface PowerPointMediaObject : PowerPointShape

@property (copy, readonly) NSString *fileName;


@end

// Every picture
@interface PowerPointPicture : PowerPointShape

@property (copy, readonly) NSString *fileName;
@property (readonly) BOOL linkToFile;
@property (copy, readonly) PowerPointPictureFormat *pictureFormat;
@property (readonly) BOOL saveWithDocument;

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(PowerPointMSFr)scale;
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(PowerPointMSFr)scale;

@end

// Every place holder
@interface PowerPointPlaceHolder : PowerPointShape

@property (copy, readonly) PowerPointPlaceholderFormat *placeHolderFormat;
@property (readonly) PowerPointPPhd placeholderType;


@end

// Every shape table
@interface PowerPointShapeTable : PowerPointShape

@property (readonly) NSInteger numberOfColumns;
@property (readonly) NSInteger numberOfRows;
@property (copy, readonly) PowerPointTable *tableObject;


@end

// Represents the soft edge formatting for a shape or range of shapes
@interface PowerPointSoftEdgeFormat : PowerPointBaseObject

@property PowerPointMSET softEdgeType;  // Returns or sets the type soft edge format object.


@end

// Every text box
@interface PowerPointTextBox : PowerPointShape

@property (readonly) PowerPointTxOr textOrientation;


@end

// Represents a single text column.
@interface PowerPointTextColumn : PowerPointBaseObject

@property NSInteger columnNumber;  // Returns or sets the index of the text column object.
@property double spacing;  // Returns or sets the spacing between text columns in a text column object.
@property PowerPointPDrT textDirection;  // Returns or sets the direction of text in the text column object.


@end

@interface PowerPointTextFrame : PowerPointBaseObject

@property PowerPointPAtS autoSize;
@property (readonly) BOOL hasText;  // Returns whether the specified text frame has text.
@property PowerPointMHzA horizontalAnchor;
@property double marginBottom;
@property double marginLeft;
@property double marginRight;
@property double marginTop;
@property (readonly) PowerPointTxOr orientation;
@property PowerPointMPFo pathFormat;  // Returns or sets the path type for the specified text frame.
@property (copy, readonly) PowerPointRuler *ruler;
@property (copy, readonly) PowerPointTextColumn *textColumn;  // Returns the text column object that represents the columns within the text frame.
@property PowerPointTxOr textOrientation;
@property (copy, readonly) PowerPointTextRange *textRange;
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns the 3-D–effect formatting properties for the specified text.
@property PowerPointMVtA verticalAnchor;
@property PowerPointMWFo warpFormat;  // Returns or sets the warp type for the specified text frame.
@property (readonly) PowerPointMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.
@property BOOL wordWrap;  // Returns or sets text break lines within or past the boundaries of the shape.
@property PowerPointMPXF wordartFormat;  // Returns or sets the WordArt type for the specified text frame.


@end

// Represents the color scheme of an Office Theme
@interface PowerPointThemeColorScheme : PowerPointBaseObject

- (SBElementArray *) themeColors;

- (NSColor *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file.
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface PowerPointThemeColor : PowerPointBaseObject

@property (copy) NSColor *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) PowerPointMCSI themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface PowerPointThemeEffectScheme : PowerPointBaseObject

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file.

@end

// Represents the font scheme of a Microsoft Office theme.
@interface PowerPointThemeFontScheme : PowerPointBaseObject

- (SBElementArray *) minorThemeFonts;
- (SBElementArray *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Every theme font
@interface PowerPointThemeFont : PowerPointBaseObject

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Every major theme font
@interface PowerPointMajorThemeFont : PowerPointThemeFont


@end

// Every minor theme font
@interface PowerPointMinorThemeFont : PowerPointThemeFont


@end

// Represents a shape's three-dimensional formatting.
@interface PowerPointThreeDFormat : PowerPointBaseObject

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property PowerPointMBlT bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property PowerPointMBlT bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy) NSColor *contourColor;  // Returns or sets the color of the contour of an object or text.
@property PowerPointMCoI contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy) NSColor *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property PowerPointMCoI extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property PowerPointMExC extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property PowerPointMPzC presetCamera;  // Returns a constant that represents the camera preset.
@property PowerPointMExD presetExtrusionDirection;  // Sets or returns the direction that the extrusion's sweep path takes away from the extruded shape.
@property PowerPointMPLd presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property PowerPointMLtT presetLightingRig;  // Returns a constant that represents the lighting preset.
@property PowerPointMlSf presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property PowerPointMPMt presetMaterial;  // Returns or sets the extrusion surface material.
@property PowerPointM3DF presetThreeDFormat;  // Sets or returns the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

@interface PowerPointWordArtFormat : PowerPointBaseObject

@property BOOL fontBold;
@property BOOL fontItalic;
@property (copy) NSString *fontName;
@property BOOL kernedPairs;
@property BOOL normalizedHeight;
@property PowerPointMPTs presetShape;
@property BOOL rotatedChars;
@property PowerPointMTxA textAlignment;
@property double tracking;
@property PowerPointMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.
@property (copy) NSString *wordArtText;

- (void) toggleVerticalText;

@end



/*
 * Text Suite
 */

@interface PowerPointTextRange : PowerPointBaseObject

- (SBElementArray *) characters;
- (SBElementArray *) words;
- (SBElementArray *) sentences;
- (SBElementArray *) lines;
- (SBElementArray *) paragraphs;
- (SBElementArray *) textFlows;

@property (readonly) double boundsHeight;
@property (readonly) double boundsWidth;
@property (copy) NSString *content;
@property (copy, readonly) PowerPointFont *font;
@property NSInteger indentLevel;
@property (readonly) double leftBounds;
@property (readonly) NSInteger offset;
@property (copy, readonly) PowerPointParagraphFormat *paragraphFormat;
@property (readonly) NSInteger textLength;
@property (readonly) double topBounds;

- (void) addPeriodsTo;
- (void) changeCaseTo:(PowerPointPCgC)to;
- (void) copyTextRange;
- (void) cutTextRange;
- (NSArray *) getRotatedTextBounds;  // Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }
- (PowerPointActionSetting *) getTextActionSettingResult:(PowerPointPMtv)result;
- (void) insertTextTextRangeInsertWhere:(PowerPointMTiP)insertWhere newText:(NSString *)newText;  // Adds text to existing text.
- (void) pasteTextRange;
- (void) removePeriodsFrom;

@end

// Every character
@interface PowerPointCharacter : PowerPointTextRange


@end

// Every line
@interface PowerPointLine : PowerPointTextRange


@end

// Every paragraph
@interface PowerPointParagraph : PowerPointTextRange


@end

// Every sentence
@interface PowerPointSentence : PowerPointTextRange


@end

// Every text flow
@interface PowerPointTextFlow : PowerPointTextRange


@end

// Every word
@interface PowerPointWord : PowerPointTextRange


@end



/*
 * Table Suite
 */

// Every cell
@interface PowerPointCell : PowerPointBaseObject

@property (readonly) BOOL selected;
@property (copy, readonly) PowerPointShape *shape;

- (PowerPointLineFormat *) getBorderEdge:(PowerPointPBrT)edge;
- (void) mergeMergeWith:(PowerPointCell *)mergeWith;
- (void) splitNumberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns;

@end

// Every column
@interface PowerPointColumn : PowerPointBaseObject

- (SBElementArray *) cells;

@property double width;


@end

// Every row
@interface PowerPointRow : PowerPointBaseObject

- (SBElementArray *) cells;

@property double height;


@end

// Every table
@interface PowerPointTable : PowerPointBaseObject

- (SBElementArray *) columns;
- (SBElementArray *) rows;

@property PowerPointPDrT tableDirection;

- (PowerPointCell *) getCellFromRow:(NSInteger)row column:(NSInteger)column;

@end

