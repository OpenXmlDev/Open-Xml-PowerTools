/* created on 9/8/2012 9:28:14 AM from peg generator V1.0 using 'ExcelFormula.txt' as input*/

using Peg.Base;
using System;
using System.IO;
using System.Text;
namespace ExcelFormula
{
      
      enum EExcelFormula{Formula= 1, Expression= 2, InfixTerms= 3, PreAndPostTerm= 4, 
                          Term= 5, RefInfixTerms= 6, RefTerm= 7, Constant= 8, RefConstant= 9, 
                          ErrorConstant= 10, LogicalConstant= 11, NumericalConstant= 12, 
                          SignificandPart= 13, WholeNumberPart= 14, FractionalPart= 15, 
                          ExponentPart= 16, StringConstant= 17, StringCharacter= 18, HighCharacter= 19, 
                          ArrayConstant= 20, ConstantListRows= 21, ConstantListRow= 22, 
                          InfixOperator= 23, ValueInfixOperator= 24, RefInfixOperator= 25, 
                          UnionOperator= 26, IntersectionOperator= 27, RangeOperator= 28, 
                          PostfixOperator= 29, PrefixOperator= 30, CellReference= 31, LocalCellReference= 32, 
                          ExternalCellReference= 33, BookPrefix= 34, BangReference= 35, 
                          SheetRangeReference= 36, SingleSheetPrefix= 37, SingleSheetReference= 38, 
                          SingleSheetArea= 39, SingleSheet= 40, SheetRange= 41, WorkbookIndex= 42, 
                          SheetName= 43, SheetNameCharacter= 44, SheetNameSpecial= 45, 
                          SheetNameBaseCharacter= 46, A1Reference= 47, A1Cell= 48, A1Area= 49, 
                          A1Column= 50, A1AbsoluteColumn= 51, A1RelativeColumn= 52, A1Row= 53, 
                          A1AbsoluteRow= 54, A1RelativeRow= 55, CellFunctionCall= 56, UserDefinedFunctionCall= 57, 
                          UserDefinedFunctionName= 58, ArgumentList= 59, Argument= 60, 
                          ArgumentExpression= 61, ArgumentInfixTerms= 62, ArgumentPreAndPostTerm= 63, 
                          ArgumentTerm= 64, ArgumentRefInfixTerms= 65, ArgumentRefTerm= 66, 
                          ArgumentInfixOperator= 67, RefArgumentInfixOperator= 68, NameReference= 69, 
                          ExternalName= 70, BangName= 71, Name= 72, NameStartCharacter= 73, 
                          NameCharacter= 74, StructureReference= 75, TableIdentifier= 76, 
                          TableName= 77, IntraTableReference= 78, InnerReference= 79, Keyword= 80, 
                          KeywordList= 81, ColumnRange= 82, Column= 83, SimpleColumnName= 84, 
                          EscapeColumnCharacter= 85, UnescapedColumnCharacter= 86, AnyNoSpaceColumnCharacter= 87, 
                          SpacedComma= 88, SpacedLBracket= 89, SpacedRBracket= 90, ws= 91};
      class ExcelFormula : PegCharParser 
      {
        
         #region Input Properties
        public static EncodingClass encodingClass = EncodingClass.ascii;
        public static UnicodeDetection unicodeDetection = UnicodeDetection.notApplicable;
        #endregion Input Properties
        #region Constructors
        public ExcelFormula()
            : base()
        {
            
        }
        public ExcelFormula(string src,TextWriter FerrOut)
			: base(src,FerrOut)
        {
            
        }
        #endregion Constructors
        #region Overrides
        public override string GetRuleNameFromId(int id)
        {
            try
            {
                   EExcelFormula ruleEnum = (EExcelFormula)id;
                    string s= ruleEnum.ToString();
                    int val;
                    if( int.TryParse(s,out val) ){
                        return base.GetRuleNameFromId(id);
                    }else{
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
        #endregion Overrides
		#region Grammar Rules
        public bool Formula()    /*Formula: Expression (!./FATAL<"end of line expected">);*/
        {

           return And(()=>  
                     Expression()
                  && (    Not(()=> Any() ) || Fatal("end of line expected")) );
		}
        public bool Expression()    /*Expression: ws InfixTerms;*/
        {

           return And(()=>    ws() && InfixTerms() );
		}
        public bool InfixTerms()    /*InfixTerms: PreAndPostTerm (InfixOperator ws PreAndPostTerm)*;*/
        {

           return And(()=>  
                     PreAndPostTerm()
                  && OptRepeat(()=>    
                      And(()=>      
                               InfixOperator()
                            && ws()
                            && PreAndPostTerm() ) ) );
		}
        public bool PreAndPostTerm()    /*PreAndPostTerm: (PrefixOperator ws)* Term (PostfixOperator ws)*;*/
        {

           return And(()=>  
                     OptRepeat(()=> And(()=>    PrefixOperator() && ws() ) )
                  && Term()
                  && OptRepeat(()=> And(()=>    PostfixOperator() && ws() ) ) );
		}
        public bool Term()    /*Term: (RefInfixTerms / '(' Expression ')' / Constant) ws;*/
        {

           return And(()=>  
                     (    
                         RefInfixTerms()
                      || And(()=>    Char('(') && Expression() && Char(')') )
                      || Constant())
                  && ws() );
		}
        public bool RefInfixTerms()    /*RefInfixTerms: RefTerm (RefInfixOperator ws RefTerm)*;*/
        {

           return And(()=>  
                     RefTerm()
                  && OptRepeat(()=>    
                      And(()=>    RefInfixOperator() && ws() && RefTerm() ) ) );
		}
        public bool RefTerm()    /*RefTerm: '(' ws RefInfixTerms ')' / RefConstant / CellFunctionCall / CellReference / UserDefinedFunctionCall
	/ NameReference / StructureReference;*/
        {

           return   
                     And(()=>    
                         Char('(')
                      && ws()
                      && RefInfixTerms()
                      && Char(')') )
                  || RefConstant()
                  || CellFunctionCall()
                  || CellReference()
                  || UserDefinedFunctionCall()
                  || NameReference()
                  || StructureReference();
		}
        public bool Constant()    /*^^Constant: ErrorConstant / LogicalConstant / NumericalConstant / StringConstant / ArrayConstant;*/
        {

           return TreeNT((int)EExcelFormula.Constant,()=>
                  
                     ErrorConstant()
                  || LogicalConstant()
                  || NumericalConstant()
                  || StringConstant()
                  || ArrayConstant() );
		}
        public bool RefConstant()    /*RefConstant: '#REF!';*/
        {

           return Char('#','R','E','F','!');
		}
        public bool ErrorConstant()    /*ErrorConstant: RefConstant / '#DIV/0!' / '#N/A' / '#NAME?' / '#NULL!' / '#NUM!' / '#VALUE!' / '#GETTING_DATA';*/
        {

           return   
                     RefConstant()
                  || Char('#','D','I','V','/','0','!')
                  || Char('#','N','/','A')
                  || Char('#','N','A','M','E','?')
                  || Char('#','N','U','L','L','!')
                  || Char('#','N','U','M','!')
                  || Char('#','V','A','L','U','E','!')
                  || Char("#GETTING_DATA");
		}
        public bool LogicalConstant()    /*LogicalConstant: 'FALSE' / 'TRUE';*/
        {

           return     Char('F','A','L','S','E') || Char('T','R','U','E');
		}
        public bool NumericalConstant()    /*NumericalConstant: '-'? SignificandPart ExponentPart?;*/
        {

           return And(()=>  
                     Option(()=> Char('-') )
                  && SignificandPart()
                  && Option(()=> ExponentPart() ) );
		}
        public bool SignificandPart()    /*SignificandPart: WholeNumberPart FractionalPart? / FractionalPart;*/
        {

           return   
                     And(()=>    
                         WholeNumberPart()
                      && Option(()=> FractionalPart() ) )
                  || FractionalPart();
		}
        public bool WholeNumberPart()    /*WholeNumberPart: [0-9]+;*/
        {

           return PlusRepeat(()=> In('0','9') );
		}
        public bool FractionalPart()    /*FractionalPart: '.' [0-9]*;*/
        {

           return And(()=>    Char('.') && OptRepeat(()=> In('0','9') ) );
		}
        public bool ExponentPart()    /*ExponentPart: 'E' ('+' / '-')? [0-9]*;*/
        {

           return And(()=>  
                     Char('E')
                  && Option(()=>     Char('+') || Char('-') )
                  && OptRepeat(()=> In('0','9') ) );
		}
        public bool StringConstant()    /*StringConstant: '"' ('""'/StringCharacter)* '"';*/
        {

           return And(()=>  
                     Char('"')
                  && OptRepeat(()=>     Char('"','"') || StringCharacter() )
                  && Char('"') );
		}
        public bool StringCharacter()    /*StringCharacter: [#-~] / '!' / ' ' / HighCharacter;*/
        {

           return   
                     In('#','~')
                  || Char('!')
                  || Char(' ')
                  || HighCharacter();
		}
        public bool HighCharacter()    /*HighCharacter: [#x80-#xFFFF];*/
        {

           return In('\u0080','\uffff');
		}
        public bool ArrayConstant()    /*^^ArrayConstant: '{' ConstantListRows '}';*/
        {

           return TreeNT((int)EExcelFormula.ArrayConstant,()=>
                And(()=>    Char('{') && ConstantListRows() && Char('}') ) );
		}
        public bool ConstantListRows()    /*ConstantListRows: ConstantListRow (';' ConstantListRow)*;*/
        {

           return And(()=>  
                     ConstantListRow()
                  && OptRepeat(()=>    
                      And(()=>    Char(';') && ConstantListRow() ) ) );
		}
        public bool ConstantListRow()    /*^^ConstantListRow: Constant (',' Constant)*;*/
        {

           return TreeNT((int)EExcelFormula.ConstantListRow,()=>
                And(()=>  
                     Constant()
                  && OptRepeat(()=> And(()=>    Char(',') && Constant() ) ) ) );
		}
        public bool InfixOperator()    /*InfixOperator: RefInfixOperator / ValueInfixOperator;*/
        {

           return     RefInfixOperator() || ValueInfixOperator();
		}
        public bool ValueInfixOperator()    /*^^ValueInfixOperator: '<>' / '>=' / '<=' / '^' / '*' / '/' / '+' / '-' / '&' / '=' / '<' / '>';*/
        {

           return TreeNT((int)EExcelFormula.ValueInfixOperator,()=>
                OneOfLiterals(optimizedLiterals0) );
		}
        public bool RefInfixOperator()    /*RefInfixOperator: RangeOperator / UnionOperator / IntersectionOperator;*/
        {

           return   
                     RangeOperator()
                  || UnionOperator()
                  || IntersectionOperator();
		}
        public bool UnionOperator()    /*^^UnionOperator: ',';*/
        {

           return TreeNT((int)EExcelFormula.UnionOperator,()=>
                Char(',') );
		}
        public bool IntersectionOperator()    /*^^IntersectionOperator: ' ';*/
        {

           return TreeNT((int)EExcelFormula.IntersectionOperator,()=>
                Char(' ') );
		}
        public bool RangeOperator()    /*^^RangeOperator: ':';*/
        {

           return TreeNT((int)EExcelFormula.RangeOperator,()=>
                Char(':') );
		}
        public bool PostfixOperator()    /*^^PostfixOperator: '%';*/
        {

           return TreeNT((int)EExcelFormula.PostfixOperator,()=>
                Char('%') );
		}
        public bool PrefixOperator()    /*^^PrefixOperator: '+' / '-';*/
        {

           return TreeNT((int)EExcelFormula.PrefixOperator,()=>
                    Char('+') || Char('-') );
		}
        public bool CellReference()    /*CellReference: ExternalCellReference / LocalCellReference;*/
        {

           return     ExternalCellReference() || LocalCellReference();
		}
        public bool LocalCellReference()    /*LocalCellReference: A1Reference;*/
        {

           return A1Reference();
		}
        public bool ExternalCellReference()    /*ExternalCellReference: BangReference / SheetRangeReference / SingleSheetReference;*/
        {

           return   
                     BangReference()
                  || SheetRangeReference()
                  || SingleSheetReference();
		}
        public bool BookPrefix()    /*BookPrefix: WorkbookIndex '!';*/
        {

           return And(()=>    WorkbookIndex() && Char('!') );
		}
        public bool BangReference()    /*BangReference: '!' (A1Reference / '#REF!');*/
        {

           return And(()=>  
                     Char('!')
                  && (    A1Reference() || Char('#','R','E','F','!')) );
		}
        public bool SheetRangeReference()    /*SheetRangeReference: SheetRange '!' A1Reference;*/
        {

           return And(()=>    SheetRange() && Char('!') && A1Reference() );
		}
        public bool SingleSheetPrefix()    /*SingleSheetPrefix: SingleSheet '!';*/
        {

           return And(()=>    SingleSheet() && Char('!') );
		}
        public bool SingleSheetReference()    /*SingleSheetReference: SingleSheetPrefix (A1Reference / '#REF!');*/
        {

           return And(()=>  
                     SingleSheetPrefix()
                  && (    A1Reference() || Char('#','R','E','F','!')) );
		}
        public bool SingleSheetArea()    /*SingleSheetArea: SingleSheetPrefix A1Area;*/
        {

           return And(()=>    SingleSheetPrefix() && A1Area() );
		}
        public bool SingleSheet()    /*SingleSheet: WorkbookIndex? SheetName / '\'' WorkbookIndex? SheetNameSpecial '\'';*/
        {

           return   
                     And(()=>    
                         Option(()=> WorkbookIndex() )
                      && SheetName() )
                  || And(()=>    
                         Char('\'')
                      && Option(()=> WorkbookIndex() )
                      && SheetNameSpecial()
                      && Char('\'') );
		}
        public bool SheetRange()    /*SheetRange: WorkbookIndex? SheetName ':' SheetName / '\'' WorkbookIndex? SheetNameSpecial ':' SheetNameSpecial '\'';*/
        {

           return   
                     And(()=>    
                         Option(()=> WorkbookIndex() )
                      && SheetName()
                      && Char(':')
                      && SheetName() )
                  || And(()=>    
                         Char('\'')
                      && Option(()=> WorkbookIndex() )
                      && SheetNameSpecial()
                      && Char(':')
                      && SheetNameSpecial()
                      && Char('\'') );
		}
        public bool WorkbookIndex()    /*^^WorkbookIndex: '[' WholeNumberPart ']';*/
        {

           return TreeNT((int)EExcelFormula.WorkbookIndex,()=>
                And(()=>    Char('[') && WholeNumberPart() && Char(']') ) );
		}
        public bool SheetName()    /*^^SheetName: SheetNameCharacter+;*/
        {

           return TreeNT((int)EExcelFormula.SheetName,()=>
                PlusRepeat(()=> SheetNameCharacter() ) );
		}
        public bool SheetNameCharacter()    /*SheetNameCharacter: [A-Za-z0-9._] / HighCharacter;*/
        {

           return   
                     (In('A','Z', 'a','z', '0','9')||OneOf("._"))
                  || HighCharacter();
		}
        public bool SheetNameSpecial()    /*^^SheetNameSpecial: SheetNameBaseCharacter ('\'\''* SheetNameBaseCharacter)*;*/
        {

           return TreeNT((int)EExcelFormula.SheetNameSpecial,()=>
                And(()=>  
                     SheetNameBaseCharacter()
                  && OptRepeat(()=>    
                      And(()=>      
                               OptRepeat(()=> Char('\'','\'') )
                            && SheetNameBaseCharacter() ) ) ) );
		}
        public bool SheetNameBaseCharacter()    /*SheetNameBaseCharacter: [A-Za-z0-9!"#$%&()+,-.;<=>@^_`{|}~ ] / HighCharacter;*/
        {

           return     OneOf(optimizedCharset0) || HighCharacter();
		}
        public bool A1Reference()    /*^^A1Reference: (A1Column ':' A1Column) / (A1Row ':' A1Row) / A1Area / A1Cell;*/
        {

           return TreeNT((int)EExcelFormula.A1Reference,()=>
                  
                     And(()=>    A1Column() && Char(':') && A1Column() )
                  || And(()=>    A1Row() && Char(':') && A1Row() )
                  || A1Area()
                  || A1Cell() );
		}
        public bool A1Cell()    /*A1Cell: A1Column A1Row !NameCharacter;*/
        {

           return And(()=>  
                     A1Column()
                  && A1Row()
                  && Not(()=> NameCharacter() ) );
		}
        public bool A1Area()    /*A1Area: A1Cell ':' A1Cell;*/
        {

           return And(()=>    A1Cell() && Char(':') && A1Cell() );
		}
        public bool A1Column()    /*^^A1Column: A1AbsoluteColumn / A1RelativeColumn;*/
        {

           return TreeNT((int)EExcelFormula.A1Column,()=>
                    A1AbsoluteColumn() || A1RelativeColumn() );
		}
        public bool A1AbsoluteColumn()    /*A1AbsoluteColumn: '$' A1RelativeColumn;*/
        {

           return And(()=>    Char('$') && A1RelativeColumn() );
		}
        public bool A1RelativeColumn()    /*A1RelativeColumn: 'XF' [A-D] / 'X' [A-E] [A-Z] / [A-W][A-Z][A-Z] / [A-Z][A-Z] / [A-Z];*/
        {

           return   
                     And(()=>    Char('X','F') && In('A','D') )
                  || And(()=>    Char('X') && In('A','E') && In('A','Z') )
                  || And(()=>    In('A','W') && In('A','Z') && In('A','Z') )
                  || And(()=>    In('A','Z') && In('A','Z') )
                  || In('A','Z');
		}
        public bool A1Row()    /*^^A1Row: A1AbsoluteRow / A1RelativeRow;*/
        {

           return TreeNT((int)EExcelFormula.A1Row,()=>
                    A1AbsoluteRow() || A1RelativeRow() );
		}
        public bool A1AbsoluteRow()    /*A1AbsoluteRow: '$' A1RelativeRow;*/
        {

           return And(()=>    Char('$') && A1RelativeRow() );
		}
        public bool A1RelativeRow()    /*A1RelativeRow: [1-9][0-9]*;*/
        {

           return And(()=>    In('1','9') && OptRepeat(()=> In('0','9') ) );
		}
        public bool CellFunctionCall()    /*^^CellFunctionCall: A1Cell '(' ArgumentList ')';*/
        {

           return TreeNT((int)EExcelFormula.CellFunctionCall,()=>
                And(()=>  
                     A1Cell()
                  && Char('(')
                  && ArgumentList()
                  && Char(')') ) );
		}
        public bool UserDefinedFunctionCall()    /*^^UserDefinedFunctionCall: UserDefinedFunctionName '(' ArgumentList ')';*/
        {

           return TreeNT((int)EExcelFormula.UserDefinedFunctionCall,()=>
                And(()=>  
                     UserDefinedFunctionName()
                  && Char('(')
                  && ArgumentList()
                  && Char(')') ) );
		}
        public bool UserDefinedFunctionName()    /*UserDefinedFunctionName: NameReference;*/
        {

           return NameReference();
		}
        public bool ArgumentList()    /*ArgumentList: Argument (',' Argument)*;*/
        {

           return And(()=>  
                     Argument()
                  && OptRepeat(()=> And(()=>    Char(',') && Argument() ) ) );
		}
        public bool Argument()    /*Argument: ArgumentExpression / ws;*/
        {

           return     ArgumentExpression() || ws();
		}
        public bool ArgumentExpression()    /*^^ArgumentExpression: ws ArgumentInfixTerms;*/
        {

           return TreeNT((int)EExcelFormula.ArgumentExpression,()=>
                And(()=>    ws() && ArgumentInfixTerms() ) );
		}
        public bool ArgumentInfixTerms()    /*ArgumentInfixTerms: ArgumentPreAndPostTerm (ArgumentInfixOperator ws ArgumentPreAndPostTerm)*;*/
        {

           return And(()=>  
                     ArgumentPreAndPostTerm()
                  && OptRepeat(()=>    
                      And(()=>      
                               ArgumentInfixOperator()
                            && ws()
                            && ArgumentPreAndPostTerm() ) ) );
		}
        public bool ArgumentPreAndPostTerm()    /*ArgumentPreAndPostTerm: (PrefixOperator ws)* ArgumentTerm (PostfixOperator ws)*;*/
        {

           return And(()=>  
                     OptRepeat(()=> And(()=>    PrefixOperator() && ws() ) )
                  && ArgumentTerm()
                  && OptRepeat(()=> And(()=>    PostfixOperator() && ws() ) ) );
		}
        public bool ArgumentTerm()    /*ArgumentTerm: (ArgumentRefInfixTerms / '(' Expression ')' / Constant) ws;*/
        {

           return And(()=>  
                     (    
                         ArgumentRefInfixTerms()
                      || And(()=>    Char('(') && Expression() && Char(')') )
                      || Constant())
                  && ws() );
		}
        public bool ArgumentRefInfixTerms()    /*ArgumentRefInfixTerms: ArgumentRefTerm (RefArgumentInfixOperator ws ArgumentRefTerm)*;*/
        {

           return And(()=>  
                     ArgumentRefTerm()
                  && OptRepeat(()=>    
                      And(()=>      
                               RefArgumentInfixOperator()
                            && ws()
                            && ArgumentRefTerm() ) ) );
		}
        public bool ArgumentRefTerm()    /*ArgumentRefTerm: '(' ws RefInfixTerms ')' / RefConstant / CellFunctionCall / CellReference / UserDefinedFunctionCall
	/ NameReference / StructureReference;*/
        {

           return   
                     And(()=>    
                         Char('(')
                      && ws()
                      && RefInfixTerms()
                      && Char(')') )
                  || RefConstant()
                  || CellFunctionCall()
                  || CellReference()
                  || UserDefinedFunctionCall()
                  || NameReference()
                  || StructureReference();
		}
        public bool ArgumentInfixOperator()    /*ArgumentInfixOperator: RefArgumentInfixOperator / ValueInfixOperator;*/
        {

           return     RefArgumentInfixOperator() || ValueInfixOperator();
		}
        public bool RefArgumentInfixOperator()    /*RefArgumentInfixOperator: RangeOperator / IntersectionOperator;*/
        {

           return     RangeOperator() || IntersectionOperator();
		}
        public bool NameReference()    /*^^NameReference: (ExternalName / Name) !'[';*/
        {

           return TreeNT((int)EExcelFormula.NameReference,()=>
                And(()=>  
                     (    ExternalName() || Name())
                  && Not(()=> Char('[') ) ) );
		}
        public bool ExternalName()    /*ExternalName: BangName / (SingleSheetPrefix / BookPrefix) Name;*/
        {

           return   
                     BangName()
                  || And(()=>    
                         (    SingleSheetPrefix() || BookPrefix())
                      && Name() );
		}
        public bool BangName()    /*BangName: '!' Name;*/
        {

           return And(()=>    Char('!') && Name() );
		}
        public bool Name()    /*Name: NameStartCharacter NameCharacter*;*/
        {

           return And(()=>  
                     NameStartCharacter()
                  && OptRepeat(()=> NameCharacter() ) );
		}
        public bool NameStartCharacter()    /*NameStartCharacter: [_\\A-Za-z] / HighCharacter;*/
        {

           return   
                     (In('A','Z', 'a','z')||OneOf("_\\"))
                  || HighCharacter();
		}
        public bool NameCharacter()    /*NameCharacter: NameStartCharacter / [0-9] / '.' / '?' / HighCharacter;*/
        {

           return   
                     NameStartCharacter()
                  || In('0','9')
                  || Char('.')
                  || Char('?')
                  || HighCharacter();
		}
        public bool StructureReference()    /*^^StructureReference: TableIdentifier? IntraTableReference;*/
        {

           return TreeNT((int)EExcelFormula.StructureReference,()=>
                And(()=>  
                     Option(()=> TableIdentifier() )
                  && IntraTableReference() ) );
		}
        public bool TableIdentifier()    /*TableIdentifier: BookPrefix? TableName;*/
        {

           return And(()=>    Option(()=> BookPrefix() ) && TableName() );
		}
        public bool TableName()    /*TableName: Name;*/
        {

           return Name();
		}
        public bool IntraTableReference()    /*IntraTableReference: SpacedLBracket InnerReference SpacedRBracket / Keyword / '[' SimpleColumnName ']';*/
        {

           return   
                     And(()=>    
                         SpacedLBracket()
                      && InnerReference()
                      && SpacedRBracket() )
                  || Keyword()
                  || And(()=>    
                         Char('[')
                      && SimpleColumnName()
                      && Char(']') );
		}
        public bool InnerReference()    /*InnerReference: (KeywordList SpacedComma)? ColumnRange / KeywordList;*/
        {

           return   
                     And(()=>    
                         Option(()=>      
                            And(()=>    KeywordList() && SpacedComma() ) )
                      && ColumnRange() )
                  || KeywordList();
		}
        public bool Keyword()    /*Keyword: '[#All]' / '[#Data]' / '[#Headers]' / '[#Totals]' / '[#This Row]';*/
        {

           return   
                     Char('[','#','A','l','l',']')
                  || Char('[','#','D','a','t','a',']')
                  || Char("[#Headers]")
                  || Char("[#Totals]")
                  || Char("[#This Row]");
		}
        public bool KeywordList()    /*KeywordList: '[#Headers]' SpacedComma '[#Data]' / '[#Data]' SpacedComma '[#Totals]' / Keyword;*/
        {

           return   
                     And(()=>    
                         Char("[#Headers]")
                      && SpacedComma()
                      && Char('[','#','D','a','t','a',']') )
                  || And(()=>    
                         Char('[','#','D','a','t','a',']')
                      && SpacedComma()
                      && Char("[#Totals]") )
                  || Keyword();
		}
        public bool ColumnRange()    /*ColumnRange: Column (':' Column)?;*/
        {

           return And(()=>  
                     Column()
                  && Option(()=> And(()=>    Char(':') && Column() ) ) );
		}
        public bool Column()    /*Column: '[' ws SimpleColumnName ws ']' / SimpleColumnName;*/
        {

           return   
                     And(()=>    
                         Char('[')
                      && ws()
                      && SimpleColumnName()
                      && ws()
                      && Char(']') )
                  || SimpleColumnName();
		}
        public bool SimpleColumnName()    /*SimpleColumnName: AnyNoSpaceColumnCharacter+ (ws AnyNoSpaceColumnCharacter+)*;*/
        {

           return And(()=>  
                     PlusRepeat(()=> AnyNoSpaceColumnCharacter() )
                  && OptRepeat(()=>    
                      And(()=>      
                               ws()
                            && PlusRepeat(()=> AnyNoSpaceColumnCharacter() ) ) ) );
		}
        public bool EscapeColumnCharacter()    /*EscapeColumnCharacter: '\'' / '#' / '[' / ']';*/
        {

           return     Char('\'') || Char('#') || Char('[') || Char(']');
		}
        public bool UnescapedColumnCharacter()    /*UnescapedColumnCharacter: [A-Za-z0-9!"#$%&()*+,-./:;<=>?@\\^_`{|}~] / HighCharacter;*/
        {

           return     OneOf(optimizedCharset1) || HighCharacter();
		}
        public bool AnyNoSpaceColumnCharacter()    /*AnyNoSpaceColumnCharacter: ('\'' EscapeColumnCharacter) / UnescapedColumnCharacter;*/
        {

           return   
                     And(()=>    Char('\'') && EscapeColumnCharacter() )
                  || UnescapedColumnCharacter();
		}
        public bool SpacedComma()    /*SpacedComma: ' '? ',' ' '?;*/
        {

           return And(()=>  
                     Option(()=> Char(' ') )
                  && Char(',')
                  && Option(()=> Char(' ') ) );
		}
        public bool SpacedLBracket()    /*SpacedLBracket: '[' ' '?;*/
        {

           return And(()=>    Char('[') && Option(()=> Char(' ') ) );
		}
        public bool SpacedRBracket()    /*SpacedRBracket: ' '? ']';*/
        {

           return And(()=>    Option(()=> Char(' ') ) && Char(']') );
		}
        public bool ws()    /*ws: ' '*;*/
        {

           return OptRepeat(()=> Char(' ') );
		}
		#endregion Grammar Rules

        #region Optimization Data 
        internal static OptimizedCharset optimizedCharset0;
        internal static OptimizedCharset optimizedCharset1;
        
        internal static OptimizedLiterals optimizedLiterals0;
        
        static ExcelFormula()
        {
            {
               OptimizedCharset.Range[] ranges = new OptimizedCharset.Range[]
                  {new OptimizedCharset.Range('A','Z'),
                   new OptimizedCharset.Range('a','z'),
                   new OptimizedCharset.Range('0','9'),
                   new OptimizedCharset.Range(',','.'),
                   };
               char[] oneOfChars = new char[]    {'!','"','#','$','%'
                                                  ,'&','(',')','+',';'
                                                  ,'<','=','>','@','^'
                                                  ,'_','`','{','|','}'
                                                  ,'~',' '};
               optimizedCharset0= new OptimizedCharset(ranges,oneOfChars);
            }
            
            {
               OptimizedCharset.Range[] ranges = new OptimizedCharset.Range[]
                  {new OptimizedCharset.Range('A','Z'),
                   new OptimizedCharset.Range('a','z'),
                   new OptimizedCharset.Range('0','9'),
                   new OptimizedCharset.Range(',','.'),
                   };
               char[] oneOfChars = new char[]    {'!','"','#','$','%'
                                                  ,'&','(',')','*','+'
                                                  ,'/',':',';','<','='
                                                  ,'>','?','@','\\','^'
                                                  ,'_','`','{','|','}'
                                                  ,'~'};
               optimizedCharset1= new OptimizedCharset(ranges,oneOfChars);
            }
            
            
            {
               string[] literals=
               { "<>",">=","<=","^","*","/","+","-",
                  "&","=","<",">" };
               optimizedLiterals0= new OptimizedLiterals(literals);
            }

            
        }
        #endregion Optimization Data 
           }
}
