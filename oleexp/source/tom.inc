//tbShellLib - 100% Implemented

typedef enum {
    tomRTF = 1,
    tomText = 2,
    tomHTML = 3,
    tomWordDocument = 4,
    tomCreateNew = 16,
    tomCreateAlways = 32,
    tomOpenExisting = 48,
    tomOpenAlways = 64,
    tomTruncateExisting = 80,
    tomReadOnly = 256,
    tomShareDenyRead = 512,
    tomShareDenyWrite = 1024,
    tomPasteFile = 4096,
} tomOpenFlags;

typedef enum {
    tomLineSpacefloat = 0,
    tomLineSpace1pt5 = 1,
    tomLineSpaceDouble = 2,
    tomLineSpaceAtLeast = 3,
    tomLineSpaceExactly = 4,
    tomLineSpaceMultiple = 5,
} tomLineSpacingRules;

typedef enum {
    tomAlignLeft = 0,
    tomAlignCenter = 1,
    tomAlignRight = 2,
    tomAlignJustify = 3,
    tomAlignDecimal = 3,
    tomAlignBar = 4,
} tomAlign;

typedef enum {
    tomNoSelection = 0,
    tomSelectionIP = 1,
    tomSelectionNormal = 2,
    tomSelectionFrame = 3,
    tomSelectionColumn = 4,
    tomSelectionRow = 5,
    tomSelectionBlock = 6,
    tomSelectionInlineShape = 7,
    tomSelectionShape = 8,
} tomSelection;

typedef enum {
    tomFalse = 0,
    tomTrue = 0xffffffff,
    tomToggle = 0xff676982,
} tomBool;

typedef enum {
    tomAutoColor = 0xff676983,
} tomColor;

typedef enum {
    tomDefault = 0xff676984,
    tomUndefined = 0xff676981,
} tomResetFlags;

typedef enum {
    tomBackward = 0xc0000001,
    tomForward = 0x3fffffff,
} tomMove;

typedef enum {
    tomMove = 0,
    tomExtend = 1,
} tomExtend;

typedef enum {
    tomSelStartActive = 1,
    tomSelAtEOL = 2,
    tomSelOvertype = 4,
    tomSelActive = 8,
    tomSelReplace = 16,
} tomSelectionFlags;

typedef enum {
    tomEnd = 0,
    tomStart = 32,
} tomStartEnd;

typedef enum {
    tomCollapseEnd = 0,
    tomCollapseStart = 1,
} tomCollapse;

typedef enum {
    tomNone = 0,
    tomfloat = 1,
    tomWords = 2,
    tomDouble = 3,
    tomDotted = 4,
} tomUnderline;

typedef enum {
    tomSpaces = 0,
    tomDots = 1,
    tomDashes = 2,
    tomLines = 3,
} tomTabSpace;

typedef enum {
    tomTabBack = 0xfffffffd,
    tomTabNext = 0xfffffffe,
    tomTabHere = 0xffffffff,
} tomTabs;

typedef enum {
    tomListNone = 0,
    tomListBullet = 1,
    tomListNumberAsArabic = 2,
    tomListNumberAsLCLetter = 3,
    tomListNumberAsUCLetter = 4,
    tomListNumberAsLCRoman = 5,
    tomListNumberAsUCRoman = 6,
    tomListNumberAsSequence = 7,
    tomListParentheses = 0x00010000,
    tomListPeriod = 0x00020000,
    tomListPlain = 0x00030000,
} tomList;

typedef enum {
    tomCharacter = 1,
    tomWord = 2,
    tomSentence = 3,
    tomParagraph = 4,
    tomLine = 5,
    tomStory = 6,
    tomScreen = 7,
    tomSection = 8,
    tomColumn = 9,
    tomRow = 10,
    tomWindow = 11,
    tomCell = 12,
    tomCharFormat = 13,
    tomParaFormat = 14,
    tomTable = 15,
    tomObject = 16,
} tomUnits;

typedef enum {
    tomMatchWord = 2,
    tomMatchCase = 4,
    tomMatchPattern = 8,
} tomFindFlags;

typedef enum {
    tomUnknownStory = 0,
    tomMainTextStory = 1,
    tomFootnotesStory = 2,
    tomEndnotesStory = 3,
    tomCommentsStory = 4,
    tomTextFrameStory = 5,
    tomEvenPagesHeaderStory = 6,
    tomPrimaryHeaderStory = 7,
    tomEvenPagesFooterStory = 8,
    tomPrimaryFooterStory = 9,
    tomFirstPageHeaderStory = 10,
    tomFirstPageFooterStory = 11,
} tomStories;

typedef enum {
    tomNoAnimation = 0,
    tomLasVegasLights = 1,
    tomBlinkingBackground = 2,
    tomSparkleText = 3,
    tomMarchingBlackAnts = 4,
    tomMarchingRedAnts = 5,
    tomShimmer = 6,
    tomWipeDown = 7,
    tomWipeRight = 8,
    tomAnimationMax = 8,
} tomAnimation;

typedef enum {
    tomLowerCase = 0,
    tomUpperCase = 1,
    tomTitleCase = 2,
    tomSentenceCase = 4,
    tomToggleCase = 5,
} tomCase;

[
    odl,
    uuid(8CC497C0-A1DF-11CE-8098-00AA0047BE5D),
    version(1.0),
    dual,
    nonextensible,
    oleautomation
]
interface ITextDocument : IDispatch {

    [id(00000000), propget]
    HRESULT Name(
         [out, retval] BSTR* pName);

    [id(0x00000001), propget]
    HRESULT Selection(
         [out, retval] ITextSelection** ppSel);

    [id(0x00000002), propget]
    HRESULT StoryCount(
         [out, retval] long* pCount);

    [id(0x00000003), propget]
    HRESULT StoryRanges(
         [out, retval] ITextStoryRanges** ppStories);

    [id(0x00000004), propget]
    HRESULT Saved(
         [out, retval] long* pValue);

    [id(0x00000004), propput]
    HRESULT Saved(
         [in] long pValue);

    [id(0x00000005), propget]
    HRESULT DefaultTabStop(
         [out, retval] float *pValue);

    [id(0x00000005), propput]
    HRESULT DefaultTabStop(
         [in] float pValue);

    [id(0x00000006)]
    HRESULT New();

    [id(0x00000007)]
    HRESULT Open(
         [in] VARIANT* pVar, 
         [in] tomOpenFlags Flags, 
         [in, defaultvalue(0)] long CodePage);

    [id(0x00000008)]
    HRESULT Save(
         [in] VARIANT* pVar, 
         [in] tomOpenFlags Flags, 
         [in, defaultvalue(0)] long CodePage);

    [id(0x00000009)]
    HRESULT Freeze(
        [out, retval] long* pCount);

    [id(0x0000000a)]
    HRESULT Unfreeze(
        [out, retval] long* pCount);

    [id(0x0000000b)]
    HRESULT BeginEditCollection();

    [id(0x0000000c)]
    HRESULT EndEditCollection();

    [id(0x0000000d)]
    HRESULT Undo(
        [in] long Count, 
        [out, retval] long* prop);

    [id(0x0000000e)]
    HRESULT Redo(
        [in] long Count, 
        [out, retval] long* prop);

    [id(0x0000000f)]
    HRESULT Range(
        [in, defaultvalue(0)] long cp1, 
        [in, defaultvalue(-1)] long cp2, 
        [out, retval] ITextRange** ppRange);

    [id(0x00000010)]
    HRESULT RangeFromPoint(
        [in] long x, 
        [in] long y, 
        [out, retval] ITextRange** ppRange);

};

[
  odl,
  uuid(8CC497C2-A1DF-11CE-8098-00AA0047BE5D),
  version(1.0),
  dual,
  nonextensible,
  oleautomation
]
interface ITextRange : IDispatch {
    [id(00000000), propget]
    HRESULT Text([out, retval] BSTR* pbstr);

    [id(00000000), propput]
    HRESULT Text([in] BSTR pbstr);

    [id(0x00000201), propget]
    HRESULT Char([out, retval] long* pch);

    [id(0x00000201), propput]
    HRESULT Char([in] long pch);

    [id(0x00000202), propget]
    HRESULT Duplicate([out, retval] ITextRange** ppRange);

    [id(0x00000203), propget]
    HRESULT FormattedText([out, retval] ITextRange** ppRange);

    [id(0x00000203), propput]
    HRESULT FormattedText([in] ITextRange* ppRange);

    [id(0x00000204), propget]
    HRESULT Start([out, retval] long* pcpFirst);

    [id(0x00000204), propput]
    HRESULT Start([in] long pcpFirst);

    [id(0x00000205), propget]
    HRESULT End([out, retval] long* pcpLim);

    [id(0x00000205), propput]
    HRESULT End([in] long pcpLim);

    [id(0x00000206), propget]
    HRESULT Font([out, retval] ITextFont** pFont);

    [id(0x00000206), propput]
    HRESULT Font([in] ITextFont* pFont);

    [id(0x00000207), propget]
    HRESULT Para([out, retval] ITextPara** pPara);

    [id(0x00000207), propput]
    HRESULT Para([in] ITextPara* pPara);

    [id(0x00000208), propget]
    HRESULT StoryLength([out, retval] long* pcch);

    [id(0x00000209), propget]
    HRESULT StoryType([out, retval] long* pValue);

    [id(0x00000210)]
    HRESULT Collapse([in] long bStart);

    [id(0x00000211)]
    HRESULT Expand(
                    [in] long Unit, 
                    [out, retval] long* pDelta);

    [id(0x00000212)]
    HRESULT GetIndex(
                    [in] long Unit, 
                    [out, retval] long* pIndex);

    [id(0x00000213)]
    HRESULT SetIndex(
                    [in] long Unit, 
                    [in] long Index, 
                    [in] long Extend);

    [id(0x00000214)]
    HRESULT SetRange(
                    [in] long cpActive, 
                    [in] long cpOther);

    [id(0x00000215)]
    HRESULT InRange(
                    [in] ITextRange* pRange, 
                    [out, retval] long* pb);

    [id(0x00000216)]
    HRESULT InStory(
                    [in] ITextRange* pRange, 
                    [out, retval] long* pb);

    [id(0x00000217)]
    HRESULT IsEqual(
                    [in] ITextRange* pRange, 
                    [out, retval] long* pb);

    [id(0x00000218)]
    HRESULT Select();

    [id(0x00000219)]
    HRESULT StartOf(
                    [in] long Unit, 
                    [in] long Extend, 
                    [out, retval] long* pDelta);

    [id(0x00000220)]
    HRESULT EndOf(
                    [in] long Unit, 
                    [in] long Extend, 
                    [out, retval] long* pDelta);

    [id(0x00000221)]
    HRESULT Move(
                    [in] long Unit, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000222)]
    HRESULT MoveStart(
                    [in] long Unit, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000223)]
    HRESULT MoveEnd(
                    [in] long Unit, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000224)]
    HRESULT MoveWhile(
                    [in] VARIANT* Cset, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000225)]
    HRESULT MoveStartWhile(
                    [in] VARIANT* Cset, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000226)]
    HRESULT MoveEndWhile(
                    [in] VARIANT* Cset, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000227)]
    HRESULT MoveUntil(
                    [in] VARIANT* Cset, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000228)]
    HRESULT MoveStartUntil(
                    [in] VARIANT* Cset, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000229)]
    HRESULT MoveEndUntil(
                    [in] VARIANT* Cset, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000230)]
    HRESULT FindText(
                    [in] BSTR bstr, 
                    [in] long cch, 
                    [in] long Flags, 
                    [out, retval] long* pLength);

    [id(0x00000231)]
    HRESULT FindTextStart(
                    [in] BSTR bstr, 
                    [in] long cch, 
                    [in] long Flags, 
                    [out, retval] long* pLength);

    [id(0x00000232)]
    HRESULT FindTextEnd(
                    [in] BSTR bstr, 
                    [in] long cch, 
                    [in] long Flags, 
                    [out, retval] long* pLength);

    [id(0x00000233)]
    HRESULT Delete(
                    [in] long Unit, 
                    [in] long Count, 
                    [out, retval] long* pDelta);

    [id(0x00000234)]
    HRESULT Cut([out] VARIANT* pVar);

    [id(0x00000235)]
    HRESULT Copy([out] VARIANT* pVar);

    [id(0x00000236)]
    HRESULT Paste(
                    [in] VARIANT* pVar, 
                    [in] long Format);

    [id(0x00000237)]
    HRESULT CanPaste(
                    [in] VARIANT* pVar, 
                    [in] long Format, 
                    [out, retval] long* pb);

    [id(0x00000238)]
    HRESULT CanEdit([out, retval] long* pbCanEdit);

    [id(0x00000239)]
    HRESULT ChangeCase([in] long Type);

    [id(0x00000240)]
    HRESULT GetPoint(
                    [in] long Type, 
                    [out] long* px, 
                    [out] long* py);

    [id(0x00000241)]
    HRESULT SetPoint(
                    [in] long x, 
                    [in] long y, 
                    [in] long Type, 
                    [in] long Extend);

    [id(0x00000242)]
    HRESULT ScrollIntoView([in] long Value);

    [id(0x00000243)]
    HRESULT GetEmbeddedObject([out, retval] IUnknown** ppv);

};

[
  odl,
  uuid(8CC497C1-A1DF-11CE-8098-00AA0047BE5D),
  version(1.0),
  dual,
  nonextensible,
  oleautomation
]
interface ITextSelection : ITextRange {
    [id(0x00000101), propget]
    HRESULT Flags([out, retval] long* pFlags);

    [id(0x00000101), propput]
    HRESULT Flags([in] long pFlags);

    [id(0x00000102), propget]
    HRESULT Type([out, retval] long* pType);

    [id(0x00000103)]
    HRESULT MoveLeft(
                    [in] long Unit, 
                    [in] long Count, 
                    [in] long Extend, 
                    [out, retval] long* pDelta);

    [id(0x00000104)]
    HRESULT MoveRight(
                    [in] long Unit, 
                    [in] long Count, 
                    [in] long Extend, 
                    [out, retval] long* pDelta);

    [id(0x00000105)]
    HRESULT MoveUp(
                    [in] long Unit, 
                    [in] long Count, 
                    [in] long Extend, 
                    [out, retval] long* pDelta);

    [id(0x00000106)]
    HRESULT MoveDown(
                    [in] long Unit, 
                    [in] long Count, 
                    [in] long Extend, 
                    [out, retval] long* pDelta);

    [id(0x00000107)]
    HRESULT HomeKey(
                    [in] long Unit, 
                    [in] long Extend, 
                    [out, retval] long* pDelta);

    [id(0x00000108)]
    HRESULT EndKey(
                    [in] long Unit, 
                    [in] long Extend, 
                    [out, retval] long* pDelta);

    [id(0x00000109)]
    HRESULT TypeText([in] BSTR bstr);

};

[
  odl,
  uuid(8CC497C3-A1DF-11CE-8098-00AA0047BE5D),
  version(1.0),
  dual,
  nonextensible,
  oleautomation
]
interface ITextFont : IDispatch {
    [id(00000000), propget]
    HRESULT Duplicate([out, retval] ITextFont** ppFont);

    [id(00000000), propput]
    HRESULT Duplicate([in] ITextFont* ppFont);

    [id(0x00000301)]
    HRESULT CanChange([out, retval] long* pb);

    [id(0x00000302)]
    HRESULT IsEqual(
                    [in] ITextFont* pFont, 
                    [out, retval] long* pb);

    [id(0x00000303)]
    HRESULT Reset([in] long Value);

    [id(0x00000304), propget]
    HRESULT Style([out, retval] long* pValue);

    [id(0x00000304), propput]
    HRESULT Style([in] long pValue);

    [id(0x00000305), propget]
    HRESULT AllCaps([out, retval] long* pValue);

    [id(0x00000305), propput]
    HRESULT AllCaps([in] long pValue);

    [id(0x00000306), propget]
    HRESULT Animation([out, retval] long* pValue);

    [id(0x00000306), propput]
    HRESULT Animation([in] long pValue);

    [id(0x00000307), propget]
    HRESULT BackColor([out, retval] long* pValue);

    [id(0x00000307), propput]
    HRESULT BackColor([in] long pValue);

    [id(0x00000308), propget]
    HRESULT Bold([out, retval] long* pValue);

    [id(0x00000308), propput]
    HRESULT Bold([in] long pValue);

    [id(0x00000309), propget]
    HRESULT Emboss([out, retval] long* pValue);

    [id(0x00000309), propput]
    HRESULT Emboss([in] long pValue);

    [id(0x00000310), propget]
    HRESULT ForeColor([out, retval] long* pValue);

    [id(0x00000310), propput]
    HRESULT ForeColor([in] long pValue);

    [id(0x00000311), propget]
    HRESULT Hidden([out, retval] long* pValue);

    [id(0x00000311), propput]
    HRESULT Hidden([in] long pValue);

    [id(0x00000312), propget]
    HRESULT Engrave([out, retval] long* pValue);

    [id(0x00000312), propput]
    HRESULT Engrave([in] long pValue);

    [id(0x00000313), propget]
    HRESULT Italic([out, retval] long* pValue);

    [id(0x00000313), propput]
    HRESULT Italic([in] long pValue);

    [id(0x00000314), propget]
    HRESULT Kerning([out, retval] float* pValue);

    [id(0x00000314), propput]
    HRESULT Kerning([in] float pValue);

    [id(0x00000315), propget]
    HRESULT LanguageID([out, retval] long* pValue);

    [id(0x00000315), propput]
    HRESULT LanguageID([in] long pValue);

    [id(0x00000316), propget]
    HRESULT Name([out, retval] BSTR* pbstr);

    [id(0x00000316), propput]
    HRESULT Name([in] BSTR pbstr);

    [id(0x00000317), propget]
    HRESULT Outline([out, retval] long* pValue);

    [id(0x00000317), propput]
    HRESULT Outline([in] long pValue);

    [id(0x00000318), propget]
    HRESULT Position([out, retval] float* pValue);

    [id(0x00000318), propput]
    HRESULT Position([in] float pValue);

    [id(0x00000319), propget]
    HRESULT Protected([out, retval] long* pValue);

    [id(0x00000319), propput]
    HRESULT Protected([in] long pValue);

    [id(0x00000320), propget]
    HRESULT Shadow([out, retval] long* pValue);

    [id(0x00000320), propput]
    HRESULT Shadow([in] long pValue);

    [id(0x00000321), propget]
    HRESULT Size([out, retval] float* pValue);

    [id(0x00000321), propput]
    HRESULT Size([in] float pValue);

    [id(0x00000322), propget]
    HRESULT SmallCaps([out, retval] long* pValue);

    [id(0x00000322), propput]
    HRESULT SmallCaps([in] long pValue);

    [id(0x00000323), propget]
    HRESULT Spacing([out, retval] float* pValue);

    [id(0x00000323), propput]
    HRESULT Spacing([in] float pValue);

    [id(0x00000324), propget]
    HRESULT StrikeThrough([out, retval] long* pValue);

    [id(0x00000324), propput]
    HRESULT StrikeThrough([in] long pValue);

    [id(0x00000325), propget]
    HRESULT Subscript([out, retval] long* pValue);

    [id(0x00000325), propput]
    HRESULT Subscript([in] long pValue);

    [id(0x00000326), propget]
    HRESULT Superscript([out, retval] long* pValue);

    [id(0x00000326), propput]
    HRESULT Superscript([in] long pValue);

    [id(0x00000327), propget]
    HRESULT Underline([out, retval] long* pValue);

    [id(0x00000327), propput]
    HRESULT Underline([in] long pValue);

    [id(0x00000328), propget]
    HRESULT Weight([out, retval] long* pValue);

    [id(0x00000328), propput]
    HRESULT Weight([in] long pValue);

};

[
  odl,
  uuid(8CC497C4-A1DF-11CE-8098-00AA0047BE5D),
  version(1.0),
  dual,
  nonextensible,
  oleautomation
]
interface ITextPara : IDispatch {
    [id(00000000), propget]
    HRESULT Duplicate([out, retval] ITextPara** ppPara);

    [id(00000000), propput]
    HRESULT Duplicate([in] ITextPara* ppPara);

    [id(0x00000401)]
    HRESULT CanChange([out, retval] long* pb);

    [id(0x00000402)]
    HRESULT IsEqual(
                    [in] ITextPara* pPara, 
                    [out, retval] long* pb);

    [id(0x00000403)]
    HRESULT Reset([in] long Value);

    [id(0x00000404), propget]
    HRESULT Style([out, retval] long* pValue);

    [id(0x00000404), propput]
    HRESULT Style([in] long pValue);

    [id(0x00000405), propget]
    HRESULT Alignment([out, retval] long* pValue);

    [id(0x00000405), propput]
    HRESULT Alignment([in] long pValue);

    [id(0x00000406), propget]
    HRESULT Hyphenation([out, retval] long* pValue);

    [id(0x00000406), propput]
    HRESULT Hyphenation([in] long pValue);

    [id(0x00000407), propget]
    HRESULT FirstLineIndent([out, retval] float* pValue);

    [id(0x00000408), propget]
    HRESULT KeepTogether([out, retval] long* pValue);

    [id(0x00000408), propput]
    HRESULT KeepTogether([in] long pValue);

    [id(0x00000409), propget]
    HRESULT KeepWithNext([out, retval] long* pValue);

    [id(0x00000409), propput]
    HRESULT KeepWithNext([in] long pValue);

    [id(0x00000410), propget]
    HRESULT LeftIndent([out, retval] float* pValue);

    [id(0x00000411), propget]
    HRESULT LineSpacing([out, retval] float* pValue);

    [id(0x00000412), propget]
    HRESULT LineSpacingRule([out, retval] long* pValue);

    [id(0x00000413), propget]
    HRESULT ListAlignment([out, retval] long* pValue);

    [id(0x00000413), propput]
    HRESULT ListAlignment([in] long pValue);

    [id(0x00000414), propget]
    HRESULT ListLevelIndex([out, retval] long* pValue);

    [id(0x00000414), propput]
    HRESULT ListLevelIndex([in] long pValue);

    [id(0x00000415), propget]
    HRESULT ListStart([out, retval] long* pValue);

    [id(0x00000415), propput]
    HRESULT ListStart([in] long pValue);

    [id(0x00000416), propget]
    HRESULT ListTab([out, retval] float* pValue);

    [id(0x00000416), propput]
    HRESULT ListTab([in] float pValue);

    [id(0x00000417), propget]
    HRESULT ListType([out, retval] long* pValue);

    [id(0x00000417), propput]
    HRESULT ListType([in] long pValue);

    [id(0x00000418), propget]
    HRESULT NoLineNumber([out, retval] long* pValue);

    [id(0x00000418), propput]
    HRESULT NoLineNumber([in] long pValue);

    [id(0x00000419), propget]
    HRESULT PageBreakBefore([out, retval] long* pValue);

    [id(0x00000419), propput]
    HRESULT PageBreakBefore([in] long pValue);

    [id(0x00000420), propget]
    HRESULT RightIndent([out, retval] float* pValue);

    [id(0x00000420), propput]
    HRESULT RightIndent([in] float pValue);

    [id(0x00000421)]
    HRESULT SetIndents(
                    [in] float StartIndent, 
                    [in] float LeftIndent, 
                    [in] float RightIndent);

    [id(0x00000422)]
    HRESULT SetLineSpacing(
                    [in] long LineSpacingRule, 
                    [in] float LineSpacing);

    [id(0x00000423), propget]
    HRESULT SpaceAfter([out, retval] float* pValue);

    [id(0x00000423), propput]
    HRESULT SpaceAfter([in] float pValue);

    [id(0x00000424), propget]
    HRESULT SpaceBefore([out, retval] float* pValue);

    [id(0x00000424), propput]
    HRESULT SpaceBefore([in] float pValue);

    [id(0x00000425), propget]
    HRESULT WidowControl([out, retval] long* pValue);

    [id(0x00000425), propput]
    HRESULT WidowControl([in] long pValue);

    [id(0x00000426), propget]
    HRESULT TabCount([out, retval] long* pCount);

    [id(0x00000427)]
    HRESULT AddTab(
                    [in] float tbPos, 
                    [in] long tbAlign, 
                    [in] long tbLeader);

    [id(0x00000428)]
    HRESULT ClearAllTabs();

    [id(0x00000429)]
    HRESULT DeleteTab([in] float tbPos);

    [id(0x00000430)]
    HRESULT GetTab(
                    [in] long iTab, 
                    [out] float* ptbPos, 
                    [out] long* ptbAlign, 
                    [out] long* ptbLeader);

};

[
  odl,
  uuid(8CC497C5-A1DF-11CE-8098-00AA0047BE5D),
  version(1.0),
  dual,
  nonextensible,
  oleautomation
]
interface ITextStoryRanges : IDispatch {

    [id(0xfffffffc), restricted]
    HRESULT _NewEnum(
		[out, retval] VARIANT *ppunkEnum);


    [id(00000000)]
	HRESULT Item(
		[in] long Index, 
		[out, retval] ITextRange** ppRange);


    [id(0x00000002), propget]
    HRESULT Count(
		[out, retval] long* pCount);


};

interface ITextDisplays;
interface ITextFont2;
interface ITextPara2;
interface ITextSelection2;
interface ITextStoryRanges2;
interface ITextStrings;
interface ITextRange2;
interface ITextStory;
interface ITextRow;



    [
      odl,
      uuid(C241F5E0-7206-11D8-A2C7-00A0D1D6C6B3),
      version(2.0),
      dual,
      nonextensible,
      oleautomation
    ]
    interface ITextDocument2 : ITextDocument {
        [id(0x00000011), propget, helpstring("method GetCaretType")]
        HRESULT CaretType([out, retval] long* pValue);
        [id(0x00000011), propput, helpstring("method GetCaretType")]
        HRESULT CaretType([in] long pValue);
        [id(0x00000012), propget, helpstring("method GetDisplays")]
        HRESULT Displays([out, retval] ITextDisplays** ppDisplays);
        [id(0x00000013), propget, helpstring("method GetDocumentFont")]
        HRESULT DocumentFont([out, retval] ITextFont2** ppFont);
        [id(0x00000013), propput, helpstring("method GetDocumentFont")]
        HRESULT DocumentFont([in] ITextFont2* ppFont);
        [id(0x00000014), propget, helpstring("method GetDocumentPara")]
        HRESULT DocumentPara([out, retval] ITextPara2** ppPara);
        [id(0x00000014), propput, helpstring("method GetDocumentPara")]
        HRESULT DocumentPara([in] ITextPara2* ppPara);
        [id(0x00000015), propget, helpstring("method GetEastAsianFlags")]
        HRESULT EastAsianFlags([out, retval] long* pFlags);
        [id(0x00000016), propget, helpstring("method GetGenerator")]
        HRESULT Generator([out, retval] BSTR* pbstr);
        [id(0x00000017), propput, helpstring("method SetIMEInProgress")]
        HRESULT IMEInProgress([in] long rhs);
        [id(0x00000018), propget, helpstring("method GetNotificationMode")]
        HRESULT NotificationMode([out, retval] long* pValue);
        [id(0x00000018), propput, helpstring("method GetNotificationMode")]
        HRESULT NotificationMode([in] long pValue);
        [id(0x00000019), propget, helpstring("method GetSelection2")]
        HRESULT Selection2([out, retval] ITextSelection2** ppSel);
        [id(0x0000001a), propget, helpstring("method Selection2")]
        HRESULT StoryRanges2([out, retval] ITextStoryRanges2** ppStories);
        [id(0x0000001b), propget, helpstring("method GetTypographyOptions")]
        HRESULT TypographyOptions([out, retval] long* pOptions);
        [id(0x0000001c), propget, helpstring("method Version")]
        HRESULT Version([out, retval] long* pValue);
        [id(0x0000001d), propget, helpstring("method GetWindow")]
        HRESULT Window([out, retval] CURRENCY* pHwnd);
        [id(0x0000001e), helpstring("method AttachMsgFilter")]
        HRESULT AttachMsgFilter([in] IUnknown* pFilter);
        [id(0x0000001f), helpstring("method CheckTextLimit")]
        HRESULT CheckTextLimit(
                        [in] long Cch, 
                        [out] long* pcch);
        [id(0x00000020), helpstring("method GetCallManager")]
        HRESULT GetCallManager([out, retval] IUnknown** ppVoid);
        [id(0x00000021), helpstring("method GetClientRect")]
        HRESULT GetClientRect(
                        [in] long Type, 
                        [out] long* pLeft, 
                        [out] long* pTop, 
                        [out] long* pRight, 
                        [out] long* pBottom);
        [id(0x00000022), helpstring("method GetEffectColor")]
        HRESULT GetEffectColor(
                        [in] long Index, 
                        [out] long* pValue);
        [id(0x00000023), helpstring("method GetImmContext")]
        HRESULT GetImmContext([out, retval] CURRENCY* pContext);
        [id(0x00000024), helpstring("method GetPreferredFont")]
        HRESULT GetPreferredFont(
                        [in] long cp, 
                        [in] long CharRep, 
                        [in] long Options, 
                        [in] long curCharRep, 
                        [in] long curFontSize, 
                        [out] BSTR* pbstr, 
                        [out] long* pPitchAndFamily, 
                        [out] long* pNewFontSize);
        [id(0x00000025), helpstring("method GetProperty")]
        HRESULT GetProperty(
                        [in] long Type, 
                        [out] long* pValue);
        [id(0x00000026), helpstring("method GetStrings")]
        HRESULT GetStrings([out] ITextStrings** ppStrs);
        [id(0x00000027), helpstring("method Notify")]
        HRESULT Notify([in] long Notify);
        [id(0x00000028), helpstring("method Selection2")]
        HRESULT Range2(
                        [in] long cpActive, 
                        [in] long cpAnchor, 
                        [out, retval] ITextRange2** ppRange);
        [id(0x00000029), helpstring("method RangeFromPoint2")]
        HRESULT RangeFromPoint2(
                        [in] long x, 
                        [in] long y, 
                        [in] long Type, 
                        [out, retval] ITextRange2** ppRange);
        [id(0x0000002a), helpstring("method ReleaseCallManager")]
        HRESULT ReleaseCallManager([in] IUnknown* pVoid);
        [id(0x0000002b), helpstring("method ReleaseImmContext")]
        HRESULT ReleaseImmContext([in] CURRENCY Context);
        [id(0x0000002c), helpstring("method SetEffectColor")]
        HRESULT SetEffectColor(
                        [in] long Index, 
                        [in] long Value);
        [id(0x0000002d), helpstring("method SetProperty")]
        HRESULT SetProperty(
                        [in] long Type, 
                        [in] long Value);
        [id(0x0000002e), helpstring("method SetTypographyOptions")]
        HRESULT SetTypographyOptions(
                        [in] long Options, 
                        [in] long Mask);
        [id(0x0000002f), helpstring("method SysBeep")]
        HRESULT SysBeep();
        [id(0x00000030), helpstring("method Update")]
        HRESULT Update([in] long Value);
        [id(0x00000031), helpstring("method UpdateWindow")]
        HRESULT UpdateWindow();
        [id(0x00000032), helpstring("method GetMathProperties")]
        HRESULT GetMathProperties([out] long* pOptions);
        [id(0x00000033), helpstring("method SetMathProperties")]
        HRESULT SetMathProperties(
                        [in] long Options, 
                        [in] long Mask);
        [id(0x0000003c), propget]
        HRESULT ActiveStory([out, retval] ITextStory** ppStory);
        [id(0x0000003c), propput]
        HRESULT ActiveStory([in] ITextStory* ppStory);
        [id(0x0000003d), propget]
        HRESULT MainStory([out, retval] ITextStory** ppStory);
        [id(0x0000003e), propget]
        HRESULT NewStory([out, retval] ITextStory** ppStory);
        [id(0x00000042)]
        HRESULT GetStory(
                        [in] long Index, 
                        [out, retval] ITextStory** ppStory);
    };

    [
      odl,
      uuid(C241F5F2-7206-11D8-A2C7-00A0D1D6C6B3),
      version(2.0),
      dual,
      nonextensible,
      oleautomation
    ]
    interface ITextDisplays : IDispatch {
    };

    [
      odl,
      uuid(C241F5E3-7206-11D8-A2C7-00A0D1D6C6B3),
      version(2.0),
      dual,
      nonextensible,
      oleautomation
    ]
    interface ITextFont2 : ITextFont {
        [id(0x00000002), propget]
        HRESULT Count([out, retval] long* pCount);
        [id(0x00000329), propget]
        HRESULT AutoLigatures([out, retval] long* pValue);
        [id(0x00000329), propput]
        HRESULT AutoLigatures([in] long pValue);
        [id(0x0000032a), propget]
        HRESULT AutospaceAlpha([out, retval] long* pValue);
        [id(0x0000032a), propput]
        HRESULT AutospaceAlpha([in] long pValue);
        [id(0x0000032b), propget]
        HRESULT AutospaceNumeric([out, retval] long* pValue);
        [id(0x0000032b), propput]
        HRESULT AutospaceNumeric([in] long pValue);
        [id(0x0000032c), propget]
        HRESULT AutospaceParens([out, retval] long* pValue);
        [id(0x0000032c), propput]
        HRESULT AutospaceParens([in] long pValue);
        [id(0x0000032d), propget]
        HRESULT CharRep([out, retval] long* pValue);
        [id(0x0000032d), propput]
        HRESULT CharRep([in] long pValue);
        [id(0x0000032e), propget]
        HRESULT CompressionMode([out, retval] long* pValue);
        [id(0x0000032e), propput]
        HRESULT CompressionMode([in] long pValue);
        [id(0x0000032f), propget]
        HRESULT Cookie([out, retval] long* pValue);
        [id(0x0000032f), propput]
        HRESULT Cookie([in] long pValue);
        [id(0x00000330), propget]
        HRESULT DoubleStrike([out, retval] long* pValue);
        [id(0x00000330), propput]
        HRESULT DoubleStrike([in] long pValue);
        [id(0x00000331), propget]
        HRESULT Duplicate2([out, retval] ITextFont2** ppFont);
        [id(0x00000331), propput]
        HRESULT Duplicate2([in] ITextFont2* ppFont);
        [id(0x00000332), propget]
        HRESULT LinkType([out, retval] long* pValue);
        [id(0x00000333), propget]
        HRESULT MathZone([out, retval] long* pValue);
        [id(0x00000333), propput]
        HRESULT MathZone([in] long pValue);
        [id(0x00000334), propget]
        HRESULT ModWidthPairs([out, retval] long* pValue);
        [id(0x00000334), propput]
        HRESULT ModWidthPairs([in] long pValue);
        [id(0x00000335), propget]
        HRESULT ModWidthSpace([out, retval] long* pValue);
        [id(0x00000335), propput]
        HRESULT ModWidthSpace([in] long pValue);
        [id(0x00000336), propget]
        HRESULT OldNumbers([out, retval] long* pValue);
        [id(0x00000336), propput]
        HRESULT OldNumbers([in] long pValue);
        [id(0x00000337), propget]
        HRESULT Overlapping([out, retval] long* pValue);
        [id(0x00000337), propput]
        HRESULT Overlapping([in] long pValue);
        [id(0x00000338), propget]
        HRESULT PositionSubSuper([out, retval] long* pValue);
        [id(0x00000338), propput]
        HRESULT PositionSubSuper([in] long pValue);
        [id(0x00000339), propget]
        HRESULT Scaling([out, retval] long* pValue);
        [id(0x00000339), propput]
        HRESULT Scaling([in] long pValue);
        [id(0x0000033a), propget]
        HRESULT SpaceExtension([out, retval] float* pValue);
        [id(0x0000033a), propput]
        HRESULT SpaceExtension([in] float pValue);
        [id(0x0000033b), propget]
        HRESULT UnderlinePositionMode([out, retval] long* pValue);
        [id(0x0000033b), propput]
        HRESULT UnderlinePositionMode([in] long pValue);
        [id(0x00000340)]
        HRESULT GetEffects(
                        [out] long* pValue, 
                        [out] long* pMask);
        [id(0x00000341)]
        HRESULT GetEffects2(
                        [out] long* pValue, 
                        [out] long* pMask);
        [id(0x00000342)]
        HRESULT GetProperty(
                        [in] long Type, 
                        [out, retval] long* pValue);
        [id(0x00000343)]
        HRESULT GetPropertyInfo(
                        [in] long Index, 
                        [out] long* pType, 
                        [out] long* pValue);
        [id(0x00000344)]
        HRESULT IsEqual2(
                        [in] ITextFont2* pFont, 
                        [out, retval] long* pB);
        [id(0x00000345)]
        HRESULT SetEffects(
                        [in] long Value, 
                        [in] long Mask);
        [id(0x00000346)]
        HRESULT SetEffects2(
                        [in] long Value, 
                        [in] long Mask);
        [id(0x00000347)]
        HRESULT SetProperty(
                        [in] long Type, 
                        [in] long Value);
    };

    [
      odl,
      uuid(C241F5E4-7206-11D8-A2C7-00A0D1D6C6B3),
      version(2.0),
      dual,
      nonextensible,
      oleautomation
    ]
    interface ITextPara2 : ITextPara {
        [id(0x00000431), propget]
        HRESULT Borders([out, retval] IUnknown** ppBorders);
        [id(0x00000432), propget]
        HRESULT Duplicate2([out, retval] ITextPara2** ppPara);
        [id(0x00000432), propput]
        HRESULT Duplicate2([in] ITextPara2* ppPara);
        [id(0x00000433), propget]
        HRESULT FontAlignment([out, retval] long* pValue);
        [id(0x00000433), propput]
        HRESULT FontAlignment([in] long pValue);
        [id(0x00000434), propget]
        HRESULT HangingPunctuation([out, retval] long* pValue);
        [id(0x00000434), propput]
        HRESULT HangingPunctuation([in] long pValue);
        [id(0x00000435), propget]
        HRESULT SnapToGrid([out, retval] long* pValue);
        [id(0x00000435), propput]
        HRESULT SnapToGrid([in] long pValue);
        [id(0x00000436), propget]
        HRESULT TrimPunctuationAtStart([out, retval] long* pValue);
        [id(0x00000436), propput]
        HRESULT TrimPunctuationAtStart([in] long pValue);
        [id(0x00000440)]
        HRESULT GetEffects(
                        [out] long* pValue, 
                        [out] long* pMask);
        [id(0x00000441)]
        HRESULT GetProperty(
                        [in] long Type, 
                        [out, retval] long* pValue);
        [id(0x00000442)]
        HRESULT IsEqual2(
                        [in] ITextPara2* pPara, 
                        [out, retval] long* pB);
        [id(0x00000443)]
        HRESULT SetEffects(
                        [in] long Value, 
                        [in] long Mask);
        [id(0x00000444)]
        HRESULT SetProperty(
                        [in] long Type, 
                        [in] long Value);
    };

    [
      odl,
      uuid(C241F5E1-7206-11D8-A2C7-00A0D1D6C6B3),
      version(2.0),
      dual,
      nonextensible,
      oleautomation
    ]
    interface ITextSelection2 : ITextRange2 {
    };

    [
      odl,
      uuid(C241F5E2-7206-11D8-A2C7-00A0D1D6C6B3),
      version(2.0),
      dual,
      nonextensible,
      oleautomation
    ]
    interface ITextRange2 : ITextSelection {
        [id(0x00000244), propget]
        HRESULT Cch([out, retval] long* pcch);
        [id(0x00000245), propget]
        HRESULT Cells([out, retval] IUnknown** ppCells);
        [id(0x00000246), propget]
        HRESULT Column([out, retval] IUnknown** ppColumn);
        [id(0x00000247), propget]
        HRESULT Count([out, retval] long* pCount);
        [id(0x00000248), propget]
        HRESULT Duplicate2([out, retval] ITextRange2** ppRange);
        [id(0x00000249), propget]
        HRESULT Font2([out, retval] ITextFont2** ppFont);
        [id(0x00000249), propput]
        HRESULT Font2([in] ITextFont2* ppFont);
        [id(0x0000024a), propget]
        HRESULT FormattedText2([out, retval] ITextRange2** ppRange);
        [id(0x0000024a), propput]
        HRESULT FormattedText2([in] ITextRange2* ppRange);
        [id(0x0000024b), propget]
        HRESULT Gravity([out, retval] long* pValue);
        [id(0x0000024b), propput]
        HRESULT Gravity([in] long pValue);
        [id(0x0000024c), propget]
        HRESULT Para2([out, retval] ITextPara2** ppPara);
        [id(0x0000024c), propput]
        HRESULT Para2([in] ITextPara2* ppPara);
        [id(0x0000024d), propget]
        HRESULT Row([out, retval] ITextRow** ppRow);
        [id(0x0000024e), propget]
        HRESULT StartPara([out, retval] long* pValue);
        [id(0x0000024f), propget]
        HRESULT Table([out, retval] IUnknown** ppTable);
        [id(0x00000250), propget]
        HRESULT URL([out, retval] BSTR* pbstr);
        [id(0x00000250), propput]
        HRESULT URL([in] BSTR pbstr);
        [id(0x00000260)]
        HRESULT AddSubrange(
                        [in] long cp1, 
                        [in] long cp2, 
                        [in] long Activate);
        [id(0x00000261)]
        HRESULT BuildUpMath([in] long Flags);
        [id(0x00000262)]
        HRESULT DeleteSubrange(
                        [in] long cpFirst, 
                        [in] long cpLim);
        [id(0x00000263)]
        HRESULT Find(
                        [in] ITextRange2* pRange, 
                        [in] long Count, 
                        [in] long Flags, 
                        [out] long* pDelta);
        [id(0x00000264)]
        HRESULT GetChar2(
                        [out] long* pChar, 
                        [in] long Offset);
        [id(0x00000265)]
        HRESULT GetDropCap(
                        [out] long* pcLine, 
                        [out] long* pPosition);
        [id(0x00000266)]
        HRESULT GetInlineObject(
                        [out] long* pType, 
                        [out] long* pAlign, 
                        [out] long* pChar, 
                        [out] long* pChar1, 
                        [out] long* pChar2, 
                        [out] long* pCount, 
                        [out] long* pTeXStyle, 
                        [out] long* pcCol, 
                        [out] long* pLevel);
        [id(0x00000267)]
        HRESULT GetProperty(
                        [in] long Type, 
                        [out] long* pValue);
        [id(0x00000268)]
        HRESULT GetRect(
                        [in] long Type, 
                        [out] long* pLeft, 
                        [out] long* pTop, 
                        [out] long* pRight, 
                        [out] long* pBottom, 
                        [out] long* pHit);
        [id(0x00000269)]
        HRESULT GetSubrange(
                        [in] long iSubrange, 
                        [out] long* pcpFirst, 
                        [out] long* pcpLim);
        [id(0x0000026a)]
        HRESULT GetText2(
                        [in] long Flags, 
                        [out] BSTR* pbstr);
        [id(0x0000026b)]
        HRESULT HexToUnicode();
        [id(0x0000026c)]
        HRESULT InsertTable(
                        [in] long cCol, 
                        [in] long cRow, 
                        [in] long AutoFit);
        [id(0x0000026d)]
        HRESULT Linearize([in] long Flags);
        [id(0x0000026e)]
        HRESULT SetActiveSubrange(
                        [in] long cpAnchor, 
                        [in] long cpActive);
        [id(0x0000026f)]
        HRESULT SetDropCap(
                        [in] long cLine, 
                        [in] long Position);
        [id(0x00000270)]
        HRESULT SetProperty(
                        [in] long Type, 
                        [in] long Value);
        [id(0x00000271)]
        HRESULT SetText2(
                        [in] long Flags, 
                        [in] BSTR bstr);
        [id(0x00000272)]
        HRESULT UnicodeToHex();
        [id(0x00000273)]
        HRESULT SetInlineObject(
                        [in] long Type, 
                        [in] long Align, 
                        [in] long Char, 
                        [in] long Char1, 
                        [in] long Char2, 
                        [in] long Count, 
                        [in] long TeXStyle, 
                        [in] long cCol);
        [id(0x00000274)]
        HRESULT GetMathFunctionType(
                        [in] BSTR bstr, 
                        [out] long* pValue);
        [id(0x00000275)]
        HRESULT InsertImage(
                        [in] long width, 
                        [in] long Height, 
                        [in] long ascent, 
                        [in] long Type, 
                        [in] BSTR bstrAltText, 
                        [in] long pStream);
    };

    [
      odl,
      uuid(C241F5EF-7206-11D8-A2C7-00A0D1D6C6B3),
      version(2.0),
      dual,
      nonextensible,
      oleautomation
    ]
    interface ITextRow : IDispatch {
        [id(0x00000800), propget]
        HRESULT Alignment([out, retval] long* pValue);
        [id(0x00000800), propput]
        HRESULT Alignment([in] long pValue);
        [id(0x00000801), propget]
        HRESULT CellCount([out, retval] long* pValue);
        [id(0x00000801), propput]
        HRESULT CellCount([in] long pValue);
        [id(0x00000802), propget]
        HRESULT CellCountCache([out, retval] long* pValue);
        [id(0x00000802), propput]
        HRESULT CellCountCache([in] long pValue);
        [id(0x00000803), propget]
        HRESULT CellIndex([out, retval] long* pValue);
        [id(0x00000803), propput]
        HRESULT CellIndex([in] long pValue);
        [id(0x00000804), propget]
        HRESULT CellMargin([out, retval] long* pValue);
        [id(0x00000804), propput]
        HRESULT CellMargin([in] long pValue);
        [id(0x00000805), propget]
        HRESULT Height([out, retval] long* pValue);
        [id(0x00000805), propput]
        HRESULT Height([in] long pValue);
        [id(0x00000806), propget]
        HRESULT Indent([out, retval] long* pValue);
        [id(0x00000806), propput]
        HRESULT Indent([in] long pValue);
        [id(0x00000807), propget]
        HRESULT KeepTogether([out, retval] long* pValue);
        [id(0x00000807), propput]
        HRESULT KeepTogether([in] long pValue);
        [id(0x00000808), propget]
        HRESULT KeepWithNext([out, retval] long* pValue);
        [id(0x00000808), propput]
        HRESULT KeepWithNext([in] long pValue);
        [id(0x00000809), propget]
        HRESULT NestLevel([out, retval] long* pValue);
        [id(0x0000080a), propget]
        HRESULT RTL([out, retval] long* pValue);
        [id(0x0000080a), propput]
        HRESULT RTL([in] long pValue);
        [id(0x00000820), propget]
        HRESULT CellAlignment([out, retval] long* pValue);
        [id(0x00000820), propput]
        HRESULT CellAlignment([in] long pValue);
        [id(0x00000821), propget]
        HRESULT CellColorBack([out, retval] long* pValue);
        [id(0x00000821), propput]
        HRESULT CellColorBack([in] long pValue);
        [id(0x00000822), propget]
        HRESULT CellColorFore([out, retval] long* pValue);
        [id(0x00000822), propput]
        HRESULT CellColorFore([in] long pValue);
        [id(0x00000823), propget]
        HRESULT CellMergeFlags([out, retval] long* pValue);
        [id(0x00000823), propput]
        HRESULT CellMergeFlags([in] long pValue);
        [id(0x00000824), propget]
        HRESULT CellShading([out, retval] long* pValue);
        [id(0x00000824), propput]
        HRESULT CellShading([in] long pValue);
        [id(0x00000825), propget]
        HRESULT CellVerticalText([out, retval] long* pValue);
        [id(0x00000825), propput]
        HRESULT CellVerticalText([in] long pValue);
        [id(0x00000826), propget]
        HRESULT CellWidth([out, retval] long* pValue);
        [id(0x00000826), propput]
        HRESULT CellWidth([in] long pValue);
        [id(0x00000830)]
        HRESULT GetCellBorderColors(
                        [out] long* pcrLeft, 
                        [out] long* pcrTop, 
                        [out] long* pcrRight, 
                        [out] long* pcrBottom);
        [id(0x00000831)]
        HRESULT GetCellBorderWidths(
                        [out] long* pduLeft, 
                        [out] long* pduTop, 
                        [out] long* pduRight, 
                        [out] long* pduBottom);
        [id(0x00000832)]
        HRESULT SetCellBorderColors(
                        [in] long crLeft, 
                        [in] long crTop, 
                        [in] long crRight, 
                        [in] long crBottom);
        [id(0x00000833)]
        HRESULT SetCellBorderWidths(
                        [in] long duLeft, 
                        [in] long duTop, 
                        [in] long duRight, 
                        [in] long duBottom);
        [id(0x00000840)]
        HRESULT Apply(
                        [in] long cRow, 
                        [in] long Flags);
        [id(0x00000841)]
        HRESULT CanChange([out, retval] long* pValue);
        [id(0x00000842)]
        HRESULT GetProperty(
                        [in] long Type, 
                        [out] long* pValue);
        [id(0x00000843)]
        HRESULT Insert([in] long cRow);
        [id(0x00000844)]
        HRESULT IsEqual(
                        [in] ITextRow* pRow, 
                        [out, retval] long* pB);
        [id(0x00000845)]
        HRESULT Reset([in] long Value);
        [id(0x00000846)]
        HRESULT SetProperty(
                        [in] long Type, 
                        [in] long Value);
    };

    [
        odl,
        uuid(C241F5E5-7206-11D8-A2C7-00A0D1D6C6B3),
        dual
    ]
    interface ITextStoryRanges2 : ITextStoryRanges
    {
        [id(0x00000003)]
        HRESULT Item2([in] long Index, [out, retval] ITextRange2 **ppRange);
    };

    [
      odl,
      uuid(C241F5E7-7206-11D8-A2C7-00A0D1D6C6B3),
      dual,
      version(2.0)
    ]
    interface ITextStrings : IDispatch {
        [id(00000000)]
        HRESULT Item(
                        [in] long Index, 
                        [out, retval] ITextRange2** ppRange);
        [id(0x00000002), propget]
        HRESULT Count([out, retval] long* pCount);
        [id(0x00000003)]
        HRESULT Add([in] BSTR bstr);
        [id(0x00000004)]
        HRESULT Append(
                        [in] ITextRange2* pRange, 
                        [in] long iString);
        [id(0x00000005)]
        HRESULT Cat2([in] long iString);
        [id(0x00000006)]
        HRESULT CatTop2([in] BSTR bstr);
        [id(0x00000007)]
        HRESULT DeleteRange([in] ITextRange2* pRange);
        [id(0x00000008)]
        HRESULT EncodeFunction(
                        [in] long Type, 
                        [in] long Align, 
                        [in] long Char, 
                        [in] long Char1, 
                        [in] long Char2, 
                        [in] long Count, 
                        [in] long TeXStyle, 
                        [in] long cCol, 
                        [in] ITextRange2* pRange);
        [id(0x00000009)]
        HRESULT GetCch(
                        [in] long iString, 
                        [out] long* pcch);
        [id(0x0000000a)]
        HRESULT InsertNullStr([in] long iString);
        [id(0x0000000b)]
        HRESULT MoveBoundary(
                        [in] long iString, 
                        [in] long Cch);
        [id(0x0000000c)]
        HRESULT PrefixTop([in] BSTR bstr);
        [id(0x0000000d)]
        HRESULT Remove(
                        [in] long iString, 
                        [in] long cString);
        [id(0x0000000e)]
        HRESULT SetFormattedText(
                        [in] ITextRange2* pRangeD, 
                        [in] ITextRange2* pRangeS);
        [id(0x0000000f)]
        HRESULT SetOpCp(
                        [in] long iString, 
                        [in] long cp);
        [id(0x00000010)]
        HRESULT SuffixTop(
                        [in] BSTR bstr, 
                        [in] ITextRange2* pRange);
        [id(0x00000011)]
        HRESULT Swap();
    };

        [
      odl,
      uuid(C241F5F3-7206-11D8-A2C7-00A0D1D6C6B3),
      version(2.0)
    ]
    interface ITextStory : stdole.IUnknown {
        [propget]
        HRESULT _stdcall Active([out, retval] long* pValue);
        [propput]
        HRESULT _stdcall Active([in] long pValue);
        [propget]
        HRESULT _stdcall Display([out, retval] IUnknown** ppDisplay);
        [propget]
        HRESULT _stdcall Index([out, retval] long* pValue);
        [propget]
        HRESULT _stdcall Type([out, retval] long* pValue);
        [propput]
        HRESULT _stdcall Type([in] long pValue);
        HRESULT _stdcall GetProperty(
                        [in] long Type, 
                        [out] long* pValue);
        HRESULT _stdcall GetRange(
                        [in] long cpActive, 
                        [in] long cpAnchor, 
                        [out, retval] ITextRange2** ppRange);
        HRESULT _stdcall GetText(
                        [in] long Flags, 
                        [out, retval] BSTR* pbstr);
        HRESULT _stdcall SetFormattedText([in] IUnknown* pUnk);
        HRESULT _stdcall SetProperty(
                        [in] long Type, 
                        [in] long Value);
        HRESULT _stdcall SetText(
                        [in] long Flags, 
                        [in] BSTR bstr);
    };


typedef enum TXTBACKSTYLE {
	TXTBACK_TRANSPARENT = 0,		//@emem	background should show through
	TXTBACK_OPAQUE,					//@emem	erase background
} TXTBACKSTYLE;


/*
 *	TXTHITRESULT
 *
 *	@enum	Defines different hitresults
 */
typedef enum TXTHITRESULT {
	TXTHITRESULT_NOHIT			= 0,	//@emem	no hit
	TXTHITRESULT_TRANSPARENT	= 1,	//@emem point is within the text's rectangle, but 
										//in a transparent region
	TXTHITRESULT_CLOSE			= 2,	//@emem	point is close to the text
	TXTHITRESULT_HIT			= 3		//@emem dead-on hit
} TXTHITRESULT;

/*
 *	TXTNATURALSIZE
 *
 *	@enum	useful values for TxGetNaturalSize.
 *
  *	@xref <mf CTxtEdit::TxGetNaturalSize>
 */
typedef enum TXTNATURALSIZE {
	TXTNS_FITTOCONTENT2	  = 0,			//@emem Get size that fits indented content
	TXTNS_FITTOCONTENT	  = 1,			//@emem Get size that fits content
	TXTNS_ROUNDTOLINE	  = 2,			//@emem Round to nearest whole line
	TXTNS_FITTOCONTENT3	  = 3,			//@emem Get size that fits indented
										//		content + trailing whitespace
	TXTNS_FITTOCONTENTWSP = 4,			//@emem Get size that fits unindented
										//		content + trailing whitespace
	TXTNS_INCLUDELASTLINE = 0x40000000,	//@emem In plain-text ctrl, inc height
										//		of a final CR
	TXTNS_EMU			  = 0x80000000	//@emem EMUs, not pixels are used as
										//		measurement units (both ways)
} TXTNATURALSIZE;

/*
 *	TXTVIEW
 *
 *	@enum	useful values for TxDraw lViewId parameter
 *
  *	@xref <mf CTxtEdit::TxDraw>
 */
typedef enum TXTVIEW { 
	TXTVIEW_ACTIVE = 0,
	TXTVIEW_INACTIVE = -1
} TXTVIEW;


/*
 *	CHANGETYPE
 *
 *	@enum	used for CHANGENOTIFY.dwChangeType; indicates what happened 
 *			for a particular change.
 */
typedef enum CHANGETYPE
{
	CN_GENERIC		= 0,				//@emem Nothing special happened
	CN_TEXTCHANGED	= 1,				//@emem the text changed
	CN_NEWUNDO		= 2,				//@emem	A new undo action was added
	CN_NEWREDO		= 4					//@emem A new redo action was added
} CHANGETYPE;

/* 
 *	@struct CHANGENOTIFY  |
 *
 *	passed during an EN_CHANGE notification; contains information about
 *	what actually happened for a change.
 */
typedef struct CHANGENOTIFY {
	DWORD	dwChangeType;				//@field TEXT changed, etc
	long	pvCookieData; 				//@field cookie for the undo action 
										// associated with the change.
} CHANGENOTIFY;

    [
        odl,
        uuid(8D33F740-CF58-11CE-A89D-00AA006CADC5)
    ]
    interface ITextServices : stdole.IUnknown
    {
        HRESULT TxSendMessage([in] long msg, [in] long wParam, [in] LPVOID lParam, [out] long *plResult);
         HRESULT		TxDraw(	
							[in] DWORD dwDrawAspect,		
							[in] LONG  lindex,			
							[in] void *pvAspect,		 
							[in] DVTARGETDEVICE *ptd,									
							[in] HDC hdcDraw,			
							[in] HDC hicTargetDev,		 
							[in] RECT* lprcBounds,	
							[in] RECT* lprcWBounds,	
							[in] RECT* lprcUpdate,		
							[in] BOOL pfnContinue, 
							[in] DWORD dwContinue,
							[in] LONG lViewId);	

	HRESULT		TxGetHScroll(
							[out] LONG *plMin, 
							[out] LONG *plMax, 
							[out] LONG *plPos, 
							[out] LONG *plPage,
							[out] BOOL * pfEnabled );

	//@cmember Horizontal scrollbar support
	HRESULT		TxGetVScroll(
							[out] LONG *plMin, 
							[out] LONG *plMax, 
							[out] LONG *plPos, 
							[out] LONG *plPage, 
							[out] BOOL * pfEnabled );

	//@cmember Setcursor
	HRESULT 	OnTxSetCursor(
							[in] DWORD dwDrawAspect,		
							[in] LONG  lindex,			
							[in] void * pvAspect,		 
							[in] DVTARGETDEVICE * ptd,									
							[in] HDC hdcDraw,			
							[in] HDC hicTargetDev,		 
							[in] RECT *lprcClient, 
							[in] INT x, 
							[in] INT y);

	//@cmember Hit-test
	 HRESULT 	TxQueryHitPoint(
							[in] DWORD dwDrawAspect,		
							[in] LONG  lindex,			
							[in] void * pvAspect,		 
							[in] DVTARGETDEVICE * ptd,									
							[in] HDC hdcDraw,			
							[in] HDC hicTargetDev,		 
							[in] RECT *lprcClient, 
							[in] INT x, 
							[in] INT y, 
							[out] DWORD * pHitResult);

	//@cmember Inplace activate notification
	 HRESULT		OnTxInPlaceActivate([in] RECT * prcClient);

	//@cmember Inplace deactivate notification
	 HRESULT		OnTxInPlaceDeactivate();

	//@cmember UI activate notification
	 HRESULT		OnTxUIActivate();

	//@cmember UI deactivate notification
	 HRESULT		OnTxUIDeactivate();

	//@cmember Get text in control
	 HRESULT		TxGetText([out] BSTR *pbstrText);

	//@cmember Set text in control
	 HRESULT		TxSetText([in] long pszText);
	
	//@cmember Get x position of 
	 HRESULT		TxGetCurTargetX([out] long *x);
	//@cmember Get baseline position
	 HRESULT		TxGetBaseLinePos([out] long *pos);

	//@cmember Get Size to fit / Natural size
	 HRESULT		TxGetNaturalSize(
							[in] DWORD dwAspect,
							[in] HDC hdcDraw,
							[in] HDC hicTargetDev,
							[in] DVTARGETDEVICE *ptd,
							[in] DWORD dwMode, 	
							[in] SIZE *psizelExtent,
							[out] LONG *pwidth, 
							[out] LONG *pheight);

	//@cmember Drag & drop
	 HRESULT		TxGetDropTarget([in] IDropTarget **ppDropTarget );

	//@cmember Bulk bit property change notifications
	 HRESULT		OnTxPropertyBitsChange([in] DWORD dwMask, [in] DWORD dwBits);

	//@cmember Fetch the cached drawing size (logical not physical)
		HRESULT		TxGetCachedSize([out] DWORD *pdwWidth, [out] DWORD *pdwHeight);

    };

typedef enum CARET_FLAGS
{
	CARET_NONE		= 0,	// Normal Western caret (blinking vertical bar)
	CARET_CUSTOM	= 1,	// Adorned caret, only set currently for RTL
	CARET_RTL		= 2,
	CARET_ITALIC	= 32,
	CARET_NULL		= 64,	// Nondegenerate selection: use empty bitmap
	CARET_ROTATE90	= 128	// Rotate 90 degrees clockwise
} CARET_FLAGS;

typedef struct CARET_INFO
{
	long tstrct_union;
} CARET_INFO;


[
    odl,
    uuid(13E670F4-1A5A-11CF-ABEB-00AA00B65EA1)
]
interface ITextHost : stdole.IUnknown
{

	//@cmember Get the DC for the host
	HDC			TxGetDC();

	//@cmember Release the DC gotten from the host
	INT			TxReleaseDC([in] HDC hdc);
	
	//@cmember Show the scroll bar
	BOOL		TxShowScrollBar([in] INT fnBar, [in] BOOL fShow);

	//@cmember Enable the scroll bar
	BOOL		TxEnableScrollBar ([in] INT fuSBFlags, [in] INT fuArrowflags);

	//@cmember Set the scroll range
	BOOL		TxSetScrollRange(
							[in] INT fnBar, 
							[in] LONG nMinPos, 
							[in] INT nMaxPos, 
							[in] BOOL fRedraw);

	//@cmember Set the scroll position
	BOOL		TxSetScrollPos ([in] INT fnBar, [in] INT nPos, [in] BOOL fRedraw);

	//@cmember InvalidateRect
	void		TxInvalidateRect([in] RECT *prc, [in] BOOL fMode);

	//@cmember Send a WM_PAINT to the window
	void		TxViewChange([in] BOOL fUpdate);
	
	//@cmember Create the caret
	BOOL		TxCreateCaret([in] HBITMAP hbmp, [in] INT xWidth, [in] INT yHeight);

	//@cmember Show the caret
	BOOL		TxShowCaret([in] BOOL fShow);

	//@cmember Set the caret position
	BOOL		TxSetCaretPos([in] INT x, [in] INT y);

	//@cmember Create a timer with the specified timeout
	BOOL		TxSetTimer([in] UINT idTimer, [in] UINT uTimeout);

	//@cmember Destroy a timer
	void		TxKillTimer([in] UINT idTimer);

	//@cmember Scroll the content of the specified window's client area
	void		TxScrollWindowEx (
							[in] INT dx, 
							[in] INT dy, 
							[in] RECT *lprcScroll, 
							[in] RECT *lprcClip,
							[in] long hrgnUpdate, 
							[in] RECT* lprcUpdate, 
							[in] UINT fuScroll);
	
	//@cmember Get mouse capture
	void		TxSetCapture([in] BOOL fCapture);

	//@cmember Set the focus to the text window
	void		TxSetFocus();

	//@cmember Establish a new cursor shape
	void		TxSetCursor([in] HCURSOR hcur, [in] BOOL fText);

	//@cmember Converts screen coordinates of a specified point to the client coordinates 
	BOOL		TxScreenToClient ([in] POINT* lppt);

	//@cmember Converts the client coordinates of a specified point to screen coordinates
	BOOL		TxClientToScreen ([in] POINT* lppt);

	//@cmember Request host to activate text services
	HRESULT		TxActivate( [in, out] LONG * plOldState );

	//@cmember Request host to deactivate text services
	HRESULT		TxDeactivate( [in] LONG lNewState );

	//@cmember Retrieves the coordinates of a window's client area
	HRESULT		TxGetClientRect([in, out] RECT* prc);

	//@cmember Get the view rectangle relative to the inset
	HRESULT		TxGetViewInset([in, out] RECT* prc);

	//@cmember Get the default character format for the text
	HRESULT		TxGetCharFormat([in, out] long *ppCF );

	//@cmember Get the default paragraph format for the text
	HRESULT		TxGetParaFormat([in, out] long *ppPF);

	//@cmember Get the background color for the window
	COLORREF	TxGetSysColor([in] int nIndex);

	//@cmember Get the background (either opaque or transparent)
	HRESULT		TxGetBackStyle([in, out] TXTBACKSTYLE *pstyle);

	//@cmember Get the maximum length for the text
	HRESULT		TxGetMaxLength([in, out] DWORD *plength);

	//@cmember Get the bits representing requested scroll bars for the window
	HRESULT		TxGetScrollBars([in, out] DWORD *pdwScrollBar);

	//@cmember Get the character to display for password input
	HRESULT		TxGetPasswordChar([in, out]  short *pch);

	//@cmember Get the accelerator character
	HRESULT		TxGetAcceleratorPos([in, out] LONG *pcp);

	//@cmember Get the native size
	HRESULT		TxGetExtent([in, out] SIZE* lpExtent);

	//@cmember Notify host that default character format has changed
	HRESULT		OnTxCharFormatChange ([in]  long pCF);

	//@cmember Notify host that default paragraph format has changed
	HRESULT		OnTxParaFormatChange ([in]  long pPF);

	//@cmember Bulk access to bit properties
	HRESULT		TxGetPropertyBits([in] DWORD dwMask, [in, out] DWORD *pdwBits);

	//@cmember Notify host of events
	HRESULT		TxNotify([in] DWORD iNotify, [in] long pv);

	// East Asia Methods for getting the Input Context
	long		TxImmGetContext();
	void		TxImmReleaseContext([in] long himc );

	//@cmember Returns HIMETRIC size of the control bar.
	HRESULT		TxGetSelectionBarWidth ([in, out] LONG *lSelBarWidth);
};

[
    odl,
    uuid(13E670F5-1A5A-11CF-ABEB-00AA00B65EA1)
]
interface ITextHost2 : stdole.IUnknown
{
//ITextHost Methods
	//@cmember Get the DC for the host
	HDC			TxGetDC();

	//@cmember Release the DC gotten from the host
	INT			TxReleaseDC([in] HDC hdc);
	
	//@cmember Show the scroll bar
	BOOL		TxShowScrollBar([in] INT fnBar, [in] BOOL fShow);

	//@cmember Enable the scroll bar
	BOOL		TxEnableScrollBar ([in] INT fuSBFlags, [in] INT fuArrowflags);

	//@cmember Set the scroll range
	BOOL		TxSetScrollRange(
							[in] INT fnBar, 
							[in] LONG nMinPos, 
							[in] INT nMaxPos, 
							[in] BOOL fRedraw);

	//@cmember Set the scroll position
	BOOL		TxSetScrollPos ([in] INT fnBar, [in] INT nPos, [in] BOOL fRedraw);

	//@cmember InvalidateRect
	void		TxInvalidateRect([in] RECT *prc, [in] BOOL fMode);

	//@cmember Send a WM_PAINT to the window
	void		TxViewChange([in] BOOL fUpdate);
	
	//@cmember Create the caret
	BOOL		TxCreateCaret([in] HBITMAP hbmp, [in] INT xWidth, [in] INT yHeight);

	//@cmember Show the caret
	BOOL		TxShowCaret([in] BOOL fShow);

	//@cmember Set the caret position
	BOOL		TxSetCaretPos([in] INT x, [in] INT y);

	//@cmember Create a timer with the specified timeout
	BOOL		TxSetTimer([in] UINT idTimer, [in] UINT uTimeout);

	//@cmember Destroy a timer
	void		TxKillTimer([in] UINT idTimer);

	//@cmember Scroll the content of the specified window's client area
	void		TxScrollWindowEx (
							[in] INT dx, 
							[in] INT dy, 
							[in] RECT *lprcScroll, 
							[in] RECT *lprcClip,
							[in] long hrgnUpdate, 
							[in] RECT* lprcUpdate, 
							[in] UINT fuScroll);
	
	//@cmember Get mouse capture
	void		TxSetCapture([in] BOOL fCapture);

	//@cmember Set the focus to the text window
	void		TxSetFocus();

	//@cmember Establish a new cursor shape
	void		TxSetCursor([in] HCURSOR hcur, [in] BOOL fText);

	//@cmember Converts screen coordinates of a specified point to the client coordinates 
	BOOL		TxScreenToClient ([in] POINT* lppt);

	//@cmember Converts the client coordinates of a specified point to screen coordinates
	BOOL		TxClientToScreen ([in] POINT* lppt);

	//@cmember Request host to activate text services
	HRESULT		TxActivate( [in, out] LONG * plOldState );

	//@cmember Request host to deactivate text services
	HRESULT		TxDeactivate( [in] LONG lNewState );

	//@cmember Retrieves the coordinates of a window's client area
	HRESULT		TxGetClientRect([in, out] RECT* prc);

	//@cmember Get the view rectangle relative to the inset
	HRESULT		TxGetViewInset([in, out] RECT* prc);

	//@cmember Get the default character format for the text
	HRESULT		TxGetCharFormat([in, out] long *ppCF );

	//@cmember Get the default paragraph format for the text
	HRESULT		TxGetParaFormat([in, out] long *ppPF);

	//@cmember Get the background color for the window
	COLORREF	TxGetSysColor([in] int nIndex);

	//@cmember Get the background (either opaque or transparent)
	HRESULT		TxGetBackStyle([in, out] TXTBACKSTYLE *pstyle);

	//@cmember Get the maximum length for the text
	HRESULT		TxGetMaxLength([in, out] DWORD *plength);

	//@cmember Get the bits representing requested scroll bars for the window
	HRESULT		TxGetScrollBars([in, out] DWORD *pdwScrollBar);

	//@cmember Get the character to display for password input
	HRESULT		TxGetPasswordChar([in, out]  short *pch);

	//@cmember Get the accelerator character
	HRESULT		TxGetAcceleratorPos([in, out] LONG *pcp);

	//@cmember Get the native size
	HRESULT		TxGetExtent([in, out] SIZE* lpExtent);

	//@cmember Notify host that default character format has changed
	HRESULT		OnTxCharFormatChange ([in]  long pCF);

	//@cmember Notify host that default paragraph format has changed
	HRESULT		OnTxParaFormatChange ([in]  long pPF);

	//@cmember Bulk access to bit properties
	HRESULT		TxGetPropertyBits([in] DWORD dwMask, [in, out] DWORD *pdwBits);

	//@cmember Notify host of events
	HRESULT		TxNotify([in] DWORD iNotify, [in] long pv);

	// East Asia Methods for getting the Input Context
	long		TxImmGetContext();
	void		TxImmReleaseContext([in] long himc );

	//@cmember Returns HIMETRIC size of the control bar.
	HRESULT		TxGetSelectionBarWidth ([in, out] LONG *lSelBarWidth);



//ITextHost2 Methods

	BOOL		TxIsDoubleClickPending(); 

						//@cmember Get the overall window for this control	 
	HRESULT		TxGetWindow([in, out] HWND *phwnd);

						//@cmember Set control window to foreground
	HRESULT		TxSetForegroundWindow();

						//@cmember Set control window to foreground
	HPALETTE	TxGetPalette();

						//@cmember Get East Asian flags
	HRESULT		TxGetEastAsianFlags([in, out] LONG *pFlags);

						//@cmember Routes the cursor change to the winhost
	HCURSOR		TxSetCursor2([in] HCURSOR hcur, [in] BOOL bText);

						//@cmember Notification that text services is freed
	void		TxFreeTextServicesNotification();

						//@cmember Get Edit Style flags
	HRESULT		TxGetEditStyle([in] DWORD dwItem, [in, out] DWORD *pdwData);

						//@cmember Get Window Style bits
	HRESULT		TxGetWindowStyles([in, out] DWORD *pdwStyle, [in, out] DWORD *pdwExStyle);

						//@cmember Show / hide drop caret (D2D-only)
	HRESULT		TxShowDropCaret([in] BOOL fShow, [in] HDC hdc, [in] RECT* prc);

						//@cmember Destroy caret (D2D-only)
	HRESULT 	TxDestroyCaret();

						//@cmember Get Horizontal scroll extent
	HRESULT		TxGetHorzExtent([in, out] LONG *plHorzExtent);
};


interface ID2D1RenderTarget;

[
    odl,
    uuid(8D33F741-CF58-11CE-A89D-00AA006CADC5)
]
interface ITextServices2 : ITextServices
{
	//@cmember Get Size to fit / Natural size and single-line ascent
	HRESULT		TxGetNaturalSize2(
							[in] DWORD dwAspect,
							[in] HDC hdcDraw,
							[in] HDC hicTargetDev,
							[in] DVTARGETDEVICE *ptd,
							[in] DWORD dwMode, 	
							[in] SIZE *psizelExtent,
							[out] LONG *pwidth, 
							[out] LONG *pheight, 
							[out] LONG *pascent);

 	HRESULT		TxDrawD2D(
 							[in] ID2D1RenderTarget* pRenderTarget,
 							[in] RECT* lprcBounds,
 							[in] RECT* lprcUpdate,
 							[in] LONG lViewId);
};

[
    odl,
    uuid(23969a9d-8546-4032-a1bb-73750cbf3333)
]
interface IRichEditUiaInformation : stdole.IUnknown
{
    HRESULT GetBoundaryRectangle([in, out] RECT *pUiaRect);
    HRESULT IsVisible(void);
};

[
    odl,
    uuid(F2FB5CC0-B5A9-437F-9BA2-47632082269F)
]
interface IRicheditUiaOverrides : stdole.IUnknown
{
    HRESULT GetPropertyOverrideValue([in] long propertyId, [in, out] VARIANT *pRetValue);
};

