

/* created on 9/8/2012 9:28:14 AM from peg generator V1.0 using 'ExcelFormula.txt' as input*/

using Peg.Base;
using System;
using System.IO;

namespace ExcelFormula
{
    internal enum EExcelFormula
    {
        Formula = 1, Expression = 2, InfixTerms = 3, PreAndPostTerm = 4,
        Term = 5, RefInfixTerms = 6, RefTerm = 7, Constant = 8, RefConstant = 9,
        ErrorConstant = 10, LogicalConstant = 11, NumericalConstant = 12,
        SignificandPart = 13, WholeNumberPart = 14, FractionalPart = 15,
        ExponentPart = 16, StringConstant = 17, StringCharacter = 18, HighCharacter = 19,
        ArrayConstant = 20, ConstantListRows = 21, ConstantListRow = 22,
        InfixOperator = 23, ValueInfixOperator = 24, RefInfixOperator = 25,
        UnionOperator = 26, IntersectionOperator = 27, RangeOperator = 28,
        PostfixOperator = 29, PrefixOperator = 30, CellReference = 31, LocalCellReference = 32,
        ExternalCellReference = 33, BookPrefix = 34, BangReference = 35,
        SheetRangeReference = 36, SingleSheetPrefix = 37, SingleSheetReference = 38,
        SingleSheetArea = 39, SingleSheet = 40, SheetRange = 41, WorkbookIndex = 42,
        SheetName = 43, SheetNameCharacter = 44, SheetNameSpecial = 45,
        SheetNameBaseCharacter = 46, A1Reference = 47, A1Cell = 48, A1Area = 49,
        A1Column = 50, A1AbsoluteColumn = 51, A1RelativeColumn = 52, A1Row = 53,
        A1AbsoluteRow = 54, A1RelativeRow = 55, CellFunctionCall = 56, UserDefinedFunctionCall = 57,
        UserDefinedFunctionName = 58, ArgumentList = 59, Argument = 60,
        ArgumentExpression = 61, ArgumentInfixTerms = 62, ArgumentPreAndPostTerm = 63,
        ArgumentTerm = 64, ArgumentRefInfixTerms = 65, ArgumentRefTerm = 66,
        ArgumentInfixOperator = 67, RefArgumentInfixOperator = 68, NameReference = 69,
        ExternalName = 70, BangName = 71, Name = 72, NameStartCharacter = 73,
        NameCharacter = 74, StructureReference = 75, TableIdentifier = 76,
        TableName = 77, IntraTableReference = 78, InnerReference = 79, Keyword = 80,
        KeywordList = 81, ColumnRange = 82, Column = 83, SimpleColumnName = 84,
        EscapeColumnCharacter = 85, UnescapedColumnCharacter = 86, AnyNoSpaceColumnCharacter = 87,
        SpacedComma = 88, SpacedLBracket = 89, SpacedRBracket = 90, ws = 91
    };

    internal class ExcelFormula : PegCharParser
    {
        private static readonly OptimizedCharset optimizedCharset0 = new OptimizedCharset(new[] { new OptimizedCharset.Range('A', 'Z'), new OptimizedCharset.Range('a', 'z'), new OptimizedCharset.Range('0', '9'), new OptimizedCharset.Range(',', '.'), }, new char[] { '!', '"', '#', '$', '%', '&', '(', ')', '+', ';', '<', '=', '>', '@', '^', '_', '`', '{', '|', '}', '~', ' ' });

        private static readonly OptimizedCharset optimizedCharset1 = new OptimizedCharset(new[] { new OptimizedCharset.Range('A', 'Z'), new OptimizedCharset.Range('a', 'z'), new OptimizedCharset.Range('0', '9'), new OptimizedCharset.Range(',', '.'), }, new char[] { '!', '"', '#', '$', '%', '&', '(', ')', '*', '+', '/', ':', ';', '<', '=', '>', '?', '@', '\\', '^', '_', '`', '{', '|', '}', '~' });

        private static readonly OptimizedLiterals optimizedLiterals0 = new OptimizedLiterals(new[] { "<>", ">=", "<=", "^", "*", "/", "+", "-", "&", "=", "<", ">" });
        private static readonly EncodingClass encodingClass = EncodingClass.ascii;
        private static readonly UnicodeDetection unicodeDetection = UnicodeDetection.notApplicable;

        public ExcelFormula()
        {
        }

        public ExcelFormula(string src, TextWriter FerrOut) : base(src, FerrOut)
        {
        }

        public override string GetRuleNameFromId(int id)
        {
            try
            {
                var ruleEnum = (EExcelFormula)id;
                var s = ruleEnum.ToString();
                if (int.TryParse(s, out var val))
                {
                    return base.GetRuleNameFromId(id);
                }
                else
                {
                    return s;
                }
            }
            catch (Exception)
            {
                return base.GetRuleNameFromId(id);
            }
        }

        public override void GetProperties(out EncodingClass encoding, out UnicodeDetection detection)
        {
            encoding = encodingClass;
            detection = unicodeDetection;
        }

        /// <summary>
        /// Formula: Expression (!./FATAL\<"end of line expected">)
        /// </summary>
        /// <returns></returns>
        public bool Formula()
        {
            return And(() => Expression() && (Not(() => Any()) || Fatal("end of line expected")));
        }

        /// <summary>
        /// Expression: ws InfixTerms;
        /// </summary>
        /// <returns></returns>
        public bool Expression()
        {
            return And(() => ws() && InfixTerms());
        }

        /// <summary>
        /// InfixTerms: PreAndPostTerm (InfixOperator ws PreAndPostTerm)*;
        /// </summary>
        /// <returns></returns>
        public bool InfixTerms()
        {
            return And(() => PreAndPostTerm() && OptRepeat(() => And(() => InfixOperator() && ws() && PreAndPostTerm())));
        }

        /// <summary>
        /// PreAndPostTerm: (PrefixOperator ws)* Term (PostfixOperator ws)*;
        /// </summary>
        /// <returns></returns>
        public bool PreAndPostTerm()
        {
            return And(() => OptRepeat(() => And(() => PrefixOperator() && ws())) && Term() && OptRepeat(() => And(() => PostfixOperator() && ws())));
        }

        /// <summary>
        /// Term: (RefInfixTerms / '(' Expression ')' / Constant) ws;
        /// </summary>
        /// <returns></returns>
        public bool Term()
        {
            return And(() => (RefInfixTerms() || And(() => Char('(') && Expression() && Char(')')) || Constant()) && ws());
        }

        /// <summary>
        /// RefInfixTerms: RefTerm (RefInfixOperator ws RefTerm)*;
        /// </summary>
        /// <returns></returns>
        public bool RefInfixTerms()
        {
            return And(() => RefTerm() && OptRepeat(() => And(() => RefInfixOperator() && ws() && RefTerm())));
        }

        /// <summary>
        /// RefTerm: '(' ws RefInfixTerms ')' / RefConstant / CellFunctionCall / CellReference / UserDefinedFunctionCall / NameReference / StructureReference;
        /// </summary>
        /// <returns></returns>
        public bool RefTerm()
        {
            return And(() => Char('(') && ws() && RefInfixTerms() && Char(')')) || RefConstant() || CellFunctionCall() || CellReference() || UserDefinedFunctionCall() || NameReference() || StructureReference();
        }

        /// <summary>
        /// ^^Constant: ErrorConstant / LogicalConstant / NumericalConstant / StringConstant / ArrayConstant;
        /// </summary>
        /// <returns></returns>
        public bool Constant()
        {
            return TreeNT((int)EExcelFormula.Constant, () => ErrorConstant() || LogicalConstant() || NumericalConstant() || StringConstant() || ArrayConstant());
        }

        /// <summary>
        /// RefConstant: '#REF!';
        /// </summary>
        /// <returns></returns>
        public bool RefConstant()
        {
            return Char('#', 'R', 'E', 'F', '!');
        }

        /// <summary>
        /// ErrorConstant: RefConstant / '#DIV/0!' / '#N/A' / '#NAME?' / '#NULL!' / '#NUM!' / '#VALUE!' / '#GETTING_DATA';
        /// </summary>
        /// <returns></returns>
        public bool ErrorConstant()
        {
            return RefConstant() || Char('#', 'D', 'I', 'V', '/', '0', '!') || Char('#', 'N', '/', 'A') || Char('#', 'N', 'A', 'M', 'E', '?') || Char('#', 'N', 'U', 'L', 'L', '!') || Char('#', 'N', 'U', 'M', '!') || Char('#', 'V', 'A', 'L', 'U', 'E', '!') || Char("#GETTING_DATA");
        }

        /// <summary>
        /// LogicalConstant: 'FALSE' / 'TRUE';
        /// </summary>
        /// <returns></returns>
        public bool LogicalConstant()
        {
            return Char('F', 'A', 'L', 'S', 'E') || Char('T', 'R', 'U', 'E');
        }

        /// <summary>
        /// NumericalConstant: '-'? SignificandPart ExponentPart?;
        /// </summary>
        /// <returns></returns>
        public bool NumericalConstant()
        {
            return And(() => Option(() => Char('-')) && SignificandPart() && Option(() => ExponentPart()));
        }

        /// <summary>
        /// SignificandPart: WholeNumberPart FractionalPart? / FractionalPart;
        /// </summary>
        /// <returns></returns>
        public bool SignificandPart()
        {
            return And(() => WholeNumberPart() && Option(() => FractionalPart())) || FractionalPart();
        }

        /// <summary>
        /// WholeNumberPart: [0-9]+;
        /// </summary>
        /// <returns></returns>
        public bool WholeNumberPart()
        {
            return PlusRepeat(() => In('0', '9'));
        }

        /// <summary>
        /// FractionalPart: '.' [0-9]*;
        /// </summary>
        /// <returns></returns>
        public bool FractionalPart()
        {
            return And(() => Char('.') && OptRepeat(() => In('0', '9')));
        }

        /// <summary>
        /// ExponentPart: 'E' ('+' / '-')? [0-9]*;
        /// </summary>
        /// <returns></returns>
        public bool ExponentPart()
        {
            return And(() => Char('E') && Option(() => Char('+') || Char('-')) && OptRepeat(() => In('0', '9')));
        }

        /// <summary>
        /// StringConstant: '"' ('""'/StringCharacter)* '"';
        /// </summary>
        /// <returns></returns>
        public bool StringConstant()
        {
            return And(() => Char('"') && OptRepeat(() => Char('"', '"') || StringCharacter()) && Char('"'));
        }

        /// <summary>
        /// StringCharacter: [#-~] / '!' / ' ' / HighCharacter;
        /// </summary>
        /// <returns></returns>
        public bool StringCharacter()
        {
            return In('#', '~') || Char('!') || Char(' ') || HighCharacter();
        }

        /// <summary>
        /// *HighCharacter: [#x80-#xFFFF];
        /// </summary>
        /// <returns></returns>
        public bool HighCharacter()
        {
            return In('\u0080', '\uffff');
        }

        /// <summary>
        /// ^^ArrayConstant: '{' ConstantListRows '}';
        /// </summary>
        /// <returns></returns>
        public bool ArrayConstant()
        {
            return TreeNT((int)EExcelFormula.ArrayConstant, () => And(() => Char('{') && ConstantListRows() && Char('}')));
        }

        /// <summary>
        /// ConstantListRows: ConstantListRow (';' ConstantListRow)*;
        /// </summary>
        /// <returns></returns>
        public bool ConstantListRows()
        {
            return And(() => ConstantListRow() && OptRepeat(() => And(() => Char(';') && ConstantListRow())));
        }

        /// <summary>
        /// ^^ConstantListRow: Constant (',' Constant)*;
        /// </summary>
        /// <returns></returns>
        public bool ConstantListRow()
        {
            return TreeNT((int)EExcelFormula.ConstantListRow, () => And(() => Constant() && OptRepeat(() => And(() => Char(',') && Constant()))));
        }

        /// <summary>
        /// InfixOperator: RefInfixOperator / ValueInfixOperator;
        /// </summary>
        /// <returns></returns>
        public bool InfixOperator()
        {
            return RefInfixOperator() || ValueInfixOperator();
        }

        /// <summary>
        /// ^^ValueInfixOperator: '<>' / '>=' / '<=' / '^' / '*' / '/' / '+' / '-' / '&' / '=' / '<' / '>';
        /// </summary>
        /// <returns></returns>
        public bool ValueInfixOperator()
        {
            return TreeNT((int)EExcelFormula.ValueInfixOperator, () => OneOfLiterals(optimizedLiterals0));
        }

        /// <summary>
        /// RefInfixOperator: RangeOperator / UnionOperator / IntersectionOperator;
        /// </summary>
        /// <returns></returns>
        public bool RefInfixOperator()
        {
            return RangeOperator() || UnionOperator() || IntersectionOperator();
        }

        /// <summary>
        /// ^^UnionOperator: ',';
        /// </summary>
        /// <returns></returns>
        public bool UnionOperator()
        {
            return TreeNT((int)EExcelFormula.UnionOperator, () => Char(','));
        }

        /// <summary>
        /// ^^IntersectionOperator: ' ';
        /// </summary>
        /// <returns></returns>
        public bool IntersectionOperator()
        {
            return TreeNT((int)EExcelFormula.IntersectionOperator, () => Char(' '));
        }

        /// <summary>
        /// ^^RangeOperator: ':';
        /// </summary>
        /// <returns></returns>
        public bool RangeOperator()
        {
            return TreeNT((int)EExcelFormula.RangeOperator, () => Char(':'));
        }

        /// <summary>
        /// ^^PostfixOperator: '%';
        /// </summary>
        /// <returns></returns>
        public bool PostfixOperator()
        {
            return TreeNT((int)EExcelFormula.PostfixOperator, () => Char('%'));
        }

        /// <summary>
        /// ^^PrefixOperator: '+' / '-';
        /// </summary>
        /// <returns></returns>
        public bool PrefixOperator()
        {
            return TreeNT((int)EExcelFormula.PrefixOperator, () => Char('+') || Char('-'));
        }

        /// <summary>
        /// CellReference: ExternalCellReference / LocalCellReference;
        /// </summary>
        /// <returns></returns>
        public bool CellReference()
        {
            return ExternalCellReference() || LocalCellReference();
        }

        /// <summary>
        /// LocalCellReference: A1Reference;
        /// </summary>
        /// <returns></returns>
        public bool LocalCellReference()
        {
            return A1Reference();
        }

        /// <summary>
        /// ExternalCellReference: BangReference / SheetRangeReference / SingleSheetReference;
        /// </summary>
        /// <returns></returns>
        public bool ExternalCellReference()
        {
            return BangReference() || SheetRangeReference() || SingleSheetReference();
        }

        /// <summary>
        /// BookPrefix: WorkbookIndex '!';
        /// </summary>
        /// <returns></returns>
        public bool BookPrefix()
        {
            return And(() => WorkbookIndex() && Char('!'));
        }

        /// <summary>
        /// BangReference: '!' (A1Reference / '#REF!');
        /// </summary>
        /// <returns></returns>
        public bool BangReference()
        {
            return And(() => Char('!') && (A1Reference() || Char('#', 'R', 'E', 'F', '!')));
        }

        /// <summary>
        /// SheetRangeReference: SheetRange '!' A1Reference;
        /// </summary>
        /// <returns></returns>
        public bool SheetRangeReference()
        {
            return And(() => SheetRange() && Char('!') && A1Reference());
        }

        /// <summary>
        /// SingleSheetPrefix: SingleSheet '!';
        /// </summary>
        /// <returns></returns>
        public bool SingleSheetPrefix()
        {
            return And(() => SingleSheet() && Char('!'));
        }

        /// <summary>
        /// SingleSheetReference: SingleSheetPrefix (A1Reference / '#REF!');
        /// </summary>
        /// <returns></returns>
        public bool SingleSheetReference()
        {
            return And(() => SingleSheetPrefix() && (A1Reference() || Char('#', 'R', 'E', 'F', '!')));
        }

        /// <summary>
        /// SingleSheetArea: SingleSheetPrefix A1Area;
        /// </summary>
        /// <returns></returns>
        public bool SingleSheetArea()
        {
            return And(() => SingleSheetPrefix() && A1Area());
        }

        /// <summary>
        /// SingleSheet: WorkbookIndex? SheetName / '\'' WorkbookIndex? SheetNameSpecial '\'';
        /// </summary>
        /// <returns></returns>
        public bool SingleSheet()
        {
            return
                      And(() => Option(() => WorkbookIndex()) && SheetName()) || And(() => Char('\'') && Option(() => WorkbookIndex()) && SheetNameSpecial() && Char('\''));
        }

        /// <summary>
        /// SheetRange: WorkbookIndex? SheetName ':' SheetName / '\'' WorkbookIndex? SheetNameSpecial ':' SheetNameSpecial '\'';
        /// </summary>
        /// <returns></returns>
        public bool SheetRange()
        {
            return And(() => Option(() => WorkbookIndex()) && SheetName() && Char(':') && SheetName()) || And(() => Char('\'') && Option(() => WorkbookIndex()) && SheetNameSpecial() && Char(':') && SheetNameSpecial() && Char('\''));
        }

        /// <summary>
        /// ^^WorkbookIndex: '[' WholeNumberPart ']';
        /// </summary>
        /// <returns></returns>
        public bool WorkbookIndex()
        {
            return TreeNT((int)EExcelFormula.WorkbookIndex, () => And(() => Char('[') && WholeNumberPart() && Char(']')));
        }

        /// <summary>
        /// ^^SheetName: SheetNameCharacter+;
        /// </summary>
        /// <returns></returns>
        public bool SheetName()
        {
            return TreeNT((int)EExcelFormula.SheetName, () => PlusRepeat(() => SheetNameCharacter()));
        }

        /// <summary>
        /// SheetNameCharacter: [A-Za-z0-9._] / HighCharacter;
        /// </summary>
        /// <returns></returns>
        public bool SheetNameCharacter()
        {
            return (In('A', 'Z', 'a', 'z', '0', '9') || OneOf("._")) || HighCharacter();
        }

        /// <summary>
        /// ^^SheetNameSpecial: SheetNameBaseCharacter ('\'\''* SheetNameBaseCharacter)*;
        /// </summary>
        /// <returns></returns>
        public bool SheetNameSpecial()
        {
            return TreeNT((int)EExcelFormula.SheetNameSpecial, () => And(() => SheetNameBaseCharacter() && OptRepeat(() => And(() => OptRepeat(() => Char('\'', '\'')) && SheetNameBaseCharacter()))));
        }

        /// <summary>
        /// SheetNameBaseCharacter: [A-Za-z0-9!"#$%&()+,-.;<=>@^_`{|}~ ] / HighCharacter;
        /// </summary>
        /// <returns></returns>
        public bool SheetNameBaseCharacter()
        {
            return OneOf(optimizedCharset0) || HighCharacter();
        }

        /// <summary>
        /// ^^A1Reference: (A1Column ':' A1Column) / (A1Row ':' A1Row) / A1Area / A1Cell;
        /// </summary>
        /// <returns></returns>
        public bool A1Reference()
        {
            return TreeNT((int)EExcelFormula.A1Reference, () => And(() => A1Column() && Char(':') && A1Column()) || And(() => A1Row() && Char(':') && A1Row()) || A1Area() || A1Cell());
        }

        /// <summary>
        /// A1Cell: A1Column A1Row !NameCharacter;
        /// </summary>
        /// <returns></returns>
        public bool A1Cell()
        {
            return And(() => A1Column() && A1Row() && Not(() => NameCharacter()));
        }

        /// <summary>
        /// A1Area: A1Cell ':' A1Cell;
        /// </summary>
        /// <returns></returns>
        public bool A1Area()
        {
            return And(() => A1Cell() && Char(':') && A1Cell());
        }

        /// <summary>
        /// ^^A1Column: A1AbsoluteColumn / A1RelativeColumn;
        /// </summary>
        /// <returns></returns>
        public bool A1Column()
        {
            return TreeNT((int)EExcelFormula.A1Column, () => A1AbsoluteColumn() || A1RelativeColumn());
        }

        /// <summary>
        /// A1AbsoluteColumn: '$' A1RelativeColumn;
        /// </summary>
        /// <returns></returns>
        public bool A1AbsoluteColumn()
        {
            return And(() => Char('$') && A1RelativeColumn());
        }

        /// <summary>
        /// A1RelativeColumn: 'XF' [A-D] / 'X' [A-E] [A-Z] / [A-W][A-Z][A-Z] / [A-Z][A-Z] / [A-Z];
        /// </summary>
        /// <returns></returns>
        public bool A1RelativeColumn()
        {
            return
                      And(() => Char('X', 'F') && In('A', 'D')) || And(() => Char('X') && In('A', 'E') && In('A', 'Z')) || And(() => In('A', 'W') && In('A', 'Z') && In('A', 'Z')) || And(() => In('A', 'Z') && In('A', 'Z')) || In('A', 'Z');
        }

        /// <summary>
        /// ^^A1Row: A1AbsoluteRow / A1RelativeRow;
        /// </summary>
        /// <returns></returns>
        public bool A1Row()
        {
            return TreeNT((int)EExcelFormula.A1Row, () => A1AbsoluteRow() || A1RelativeRow());
        }

        /// <summary>
        /// A1AbsoluteRow: '$' A1RelativeRow;
        /// </summary>
        /// <returns></returns>
        public bool A1AbsoluteRow()
        {
            return And(() => Char('$') && A1RelativeRow());
        }

        /// <summary>
        /// A1RelativeRow: [1-9][0-9]*;
        /// </summary>
        /// <returns></returns>
        public bool A1RelativeRow()
        {
            return And(() => In('1', '9') && OptRepeat(() => In('0', '9')));
        }

        /// <summary>
        /// ^^CellFunctionCall: A1Cell '(' ArgumentList ')';
        /// </summary>
        /// <returns></returns>
        public bool CellFunctionCall()
        {
            return TreeNT((int)EExcelFormula.CellFunctionCall, () => And(() => A1Cell() && Char('(') && ArgumentList() && Char(')')));
        }

        /// <summary>
        /// ^^UserDefinedFunctionCall: UserDefinedFunctionName '(' ArgumentList ')';
        /// </summary>
        /// <returns></returns>
        public bool UserDefinedFunctionCall()
        {
            return TreeNT((int)EExcelFormula.UserDefinedFunctionCall, () => And(() => UserDefinedFunctionName() && Char('(') && ArgumentList() && Char(')')));
        }

        /// <summary>
        /// UserDefinedFunctionName: NameReference;
        /// </summary>
        /// <returns></returns>
        public bool UserDefinedFunctionName()
        {
            return NameReference();
        }

        /// <summary>
        /// ArgumentList: Argument (',' Argument)*;
        /// </summary>
        /// <returns></returns>
        public bool ArgumentList()
        {
            return And(() => Argument() && OptRepeat(() => And(() => Char(',') && Argument())));
        }

        /// <summary>
        /// Argument: ArgumentExpression / ws;
        /// </summary>
        /// <returns></returns>
        public bool Argument()
        {
            return ArgumentExpression() || ws();
        }

        /// <summary>
        /// ^^ArgumentExpression: ws ArgumentInfixTerms;
        /// </summary>
        /// <returns></returns>
        public bool ArgumentExpression()
        {
            return TreeNT((int)EExcelFormula.ArgumentExpression, () => And(() => ws() && ArgumentInfixTerms()));
        }

        /// <summary>
        /// ArgumentInfixTerms: ArgumentPreAndPostTerm (ArgumentInfixOperator ws ArgumentPreAndPostTerm)*;
        /// </summary>
        /// <returns></returns>
        public bool ArgumentInfixTerms()
        {
            return And(() => ArgumentPreAndPostTerm() && OptRepeat(() => And(() => ArgumentInfixOperator() && ws() && ArgumentPreAndPostTerm())));
        }

        /// <summary>
        /// ArgumentPreAndPostTerm: (PrefixOperator ws)* ArgumentTerm (PostfixOperator ws)*;
        /// </summary>
        /// <returns></returns>
        public bool ArgumentPreAndPostTerm()
        {
            return And(() => OptRepeat(() => And(() => PrefixOperator() && ws())) && ArgumentTerm() && OptRepeat(() => And(() => PostfixOperator() && ws())));
        }

        /// <summary>
        /// ArgumentTerm: (ArgumentRefInfixTerms / '(' Expression ')' / Constant) ws;
        /// </summary>
        /// <returns></returns>
        public bool ArgumentTerm()
        {
            return And(() => (ArgumentRefInfixTerms() || And(() => Char('(') && Expression() && Char(')')) || Constant()) && ws());
        }

        /// <summary>
        /// ArgumentRefInfixTerms: ArgumentRefTerm (RefArgumentInfixOperator ws ArgumentRefTerm)*;
        /// </summary>
        /// <returns></returns>
        public bool ArgumentRefInfixTerms()
        {
            return And(() => ArgumentRefTerm() && OptRepeat(() => And(() => RefArgumentInfixOperator() && ws() && ArgumentRefTerm())));
        }

        /// <summary>
        /// ArgumentRefTerm: '(' ws RefInfixTerms ')' / RefConstant / CellFunctionCall / CellReference / UserDefinedFunctionCall / NameReference / StructureReference;
        /// </summary>
        /// <returns></returns>
        public bool ArgumentRefTerm()
        {
            return And(() => Char('(') && ws() && RefInfixTerms() && Char(')')) || RefConstant() || CellFunctionCall() || CellReference() || UserDefinedFunctionCall() || NameReference() || StructureReference();
        }

        /// <summary>
        /// ArgumentInfixOperator: RefArgumentInfixOperator / ValueInfixOperator;
        /// </summary>
        /// <returns></returns>
        public bool ArgumentInfixOperator()
        {
            return RefArgumentInfixOperator() || ValueInfixOperator();
        }

        /// <summary>
        /// RefArgumentInfixOperator: RangeOperator / IntersectionOperator;
        /// </summary>
        /// <returns></returns>
        public bool RefArgumentInfixOperator()
        {
            return RangeOperator() || IntersectionOperator();
        }

        /// <summary>
        /// ^^NameReference: (ExternalName / Name) !'[';
        /// </summary>
        /// <returns></returns>
        public bool NameReference()
        {
            return TreeNT((int)EExcelFormula.NameReference, () => And(() => (ExternalName() || Name()) && Not(() => Char('['))));
        }

        /// <summary>
        /// ExternalName: BangName / (SingleSheetPrefix / BookPrefix) Name;
        /// </summary>
        /// <returns></returns>
        public bool ExternalName()
        {
            return BangName() || And(() => (SingleSheetPrefix() || BookPrefix()) && Name());
        }

        /// <summary>
        /// BangName: '!' Name;
        /// </summary>
        /// <returns></returns>
        public bool BangName()
        {
            return And(() => Char('!') && Name());
        }

        /// <summary>
        /// Name: NameStartCharacter NameCharacter*;
        /// </summary>
        /// <returns></returns>
        public bool Name()
        {
            return And(() => NameStartCharacter() && OptRepeat(() => NameCharacter()));
        }

        /// <summary>
        /// NameStartCharacter: [_\\A-Za-z] / HighCharacter;
        /// </summary>
        /// <returns></returns>
        public bool NameStartCharacter()
        {
            return (In('A', 'Z', 'a', 'z') || OneOf("_\\")) || HighCharacter();
        }

        /// <summary>
        /// NameCharacter: NameStartCharacter / [0-9] / '.' / '?' / HighCharacter;
        /// </summary>
        /// <returns></returns>
        public bool NameCharacter()
        {
            return NameStartCharacter() || In('0', '9') || Char('.') || Char('?') || HighCharacter();
        }

        /// <summary>
        /// ^^StructureReference: TableIdentifier? IntraTableReference;
        /// </summary>
        /// <returns></returns>
        public bool StructureReference()
        {
            return TreeNT((int)EExcelFormula.StructureReference, () => And(() => Option(() => TableIdentifier()) && IntraTableReference()));
        }

        /// <summary>
        /// TableIdentifier: BookPrefix? TableName;
        /// </summary>
        /// <returns></returns>
        public bool TableIdentifier()
        {
            return And(() => Option(() => BookPrefix()) && TableName());
        }

        /// <summary>
        /// TableName: Name;
        /// </summary>
        /// <returns></returns>
        public bool TableName()
        {
            return Name();
        }

        /// <summary>
        /// IntraTableReference: SpacedLBracket InnerReference SpacedRBracket / Keyword / '[' SimpleColumnName ']';
        /// </summary>
        /// <returns></returns>
        public bool IntraTableReference()
        {
            return And(() => SpacedLBracket() && InnerReference() && SpacedRBracket()) || Keyword() || And(() => Char('[') && SimpleColumnName() && Char(']'));
        }

        /// <summary>
        /// InnerReference: (KeywordList SpacedComma)? ColumnRange / KeywordList;
        /// </summary>
        /// <returns></returns>
        public bool InnerReference()
        {
            return
                      And(() =>
                          Option(() =>
                             And(() => KeywordList() && SpacedComma()))
                       && ColumnRange())
                   || KeywordList();
        }

        /// <summary>
        /// Keyword: '[#All]' / '[#Data]' / '[#Headers]' / '[#Totals]' / '[#This Row]';
        /// </summary>
        /// <returns></returns>
        public bool Keyword()
        {
            return Char('[', '#', 'A', 'l', 'l', ']') || Char('[', '#', 'D', 'a', 't', 'a', ']') || Char("[#Headers]") || Char("[#Totals]") || Char("[#This Row]");
        }

        /// <summary>
        /// KeywordList: '[#Headers]' SpacedComma '[#Data]' / '[#Data]' SpacedComma '[#Totals]' / Keyword;
        /// </summary>
        /// <returns></returns>
        public bool KeywordList()
        {
            return And(() => Char("[#Headers]") && SpacedComma() && Char('[', '#', 'D', 'a', 't', 'a', ']')) || And(() => Char('[', '#', 'D', 'a', 't', 'a', ']') && SpacedComma() && Char("[#Totals]")) || Keyword();
        }

        /// <summary>
        /// ColumnRange: Column (':' Column)?;
        /// </summary>
        /// <returns></returns>
        public bool ColumnRange()
        {
            return And(() => Column() && Option(() => And(() => Char(':') && Column())));
        }

        /// <summary>
        /// Column: '[' ws SimpleColumnName ws ']' / SimpleColumnName;
        /// </summary>
        /// <returns></returns>
        public bool Column()
        {
            return And(() => Char('[') && ws() && SimpleColumnName() && ws() && Char(']')) || SimpleColumnName();
        }

        /// <summary>
        /// SimpleColumnName: AnyNoSpaceColumnCharacter+ (ws AnyNoSpaceColumnCharacter+)*;
        /// </summary>
        /// <returns></returns>
        public bool SimpleColumnName()
        {
            return And(() => PlusRepeat(() => AnyNoSpaceColumnCharacter()) && OptRepeat(() => And(() => ws() && PlusRepeat(() => AnyNoSpaceColumnCharacter()))));
        }

        /// <summary>
        /// EscapeColumnCharacter: '\'' / '#' / '[' / ']';
        /// </summary>
        /// <returns></returns>
        public bool EscapeColumnCharacter()
        {
            return Char('\'') || Char('#') || Char('[') || Char(']');
        }

        /// <summary>
        /// UnescapedColumnCharacter: [A-Za-z0-9!"#$%&()*+,-./:;<=>?@\\^_`{|}~] / HighCharacter;
        /// </summary>
        /// <returns></returns>
        public bool UnescapedColumnCharacter()
        {
            return OneOf(optimizedCharset1) || HighCharacter();
        }

        /// <summary>
        /// AnyNoSpaceColumnCharacter: ('\'' EscapeColumnCharacter) / UnescapedColumnCharacter;
        /// </summary>
        /// <returns></returns>
        public bool AnyNoSpaceColumnCharacter()
        {
            return And(() => Char('\'') && EscapeColumnCharacter()) || UnescapedColumnCharacter();
        }

        /// <summary>
        /// SpacedComma: ' '? ',' ' '?;
        /// </summary>
        /// <returns></returns>
        public bool SpacedComma()
        {
            return And(() => Option(() => Char(' ')) && Char(',') && Option(() => Char(' ')));
        }

        /// <summary>
        /// SpacedLBracket: '[' ' '?;
        /// </summary>
        /// <returns></returns>
        public bool SpacedLBracket()
        {
            return And(() => Char('[') && Option(() => Char(' ')));
        }

        /// <summary>
        /// SpacedRBracket: ' '? ']';
        /// </summary>
        /// <returns></returns>
        public bool SpacedRBracket()
        {
            return And(() => Option(() => Char(' ')) && Char(']'));
        }

        /// <summary>
        /// ws: ' '*;
        /// </summary>
        /// <returns></returns>
        public bool ws()
        {
            return OptRepeat(() => Char(' '));
        }
    }
}