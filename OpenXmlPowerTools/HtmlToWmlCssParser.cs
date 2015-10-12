/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

***************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace OpenXmlPowerTools.HtmlToWml.CSS
{
    public class CssAttribute
    {
        private string m_operand;
        private CssAttributeOperator? m_op = null;
        private string m_val;

        public string Operand
        {
            get {
                return m_operand;
            }
            set {
                m_operand = value;
            }
        }

        public CssAttributeOperator? Operator
        {
            get {
                return m_op;
            }
            set {
                m_op = value;
            }
        }

        public string CssOperatorString
        {
            get {
                if (this.m_op.HasValue)
                {
                    return this.m_op.Value.ToString();
                }
                else
                {
                    return null;
                }
            }
            set {
                this.m_op = (CssAttributeOperator)Enum.Parse(typeof(CssAttributeOperator), value);
            }
        }

        public string Value
        {
            get {
                return m_val;
            }
            set {
                m_val = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("[{0}", m_operand);
            if (m_op.HasValue)
            {
                switch (m_op.Value)
                {
                    case CssAttributeOperator.Equals:
                        sb.Append("=");
                        break;
                    case CssAttributeOperator.InList:
                        sb.Append("~=");
                        break;
                    case CssAttributeOperator.Hyphenated:
                        sb.Append("|=");
                        break;
                    case CssAttributeOperator.BeginsWith:
                        sb.Append("$=");
                        break;
                    case CssAttributeOperator.EndsWith:
                        sb.Append("^=");
                        break;
                    case CssAttributeOperator.Contains:
                        sb.Append("*=");
                        break;
                }
                sb.Append(m_val);
            }
            sb.Append("]");
            return sb.ToString();
        }
    }

    public enum CssAttributeOperator
    {
        Equals,
        InList,
        Hyphenated,
        EndsWith,
        BeginsWith,
        Contains,
    }

    public enum CssCombinator
    {
        ChildOf,
        PrecededImmediatelyBy,
        PrecededBy,
    }

    public class CssDocument : ItfRuleSetContainer
    {
        private List<CssDirective> m_dirs = new List<CssDirective>();
        private List<CssRuleSet> m_rulesets = new List<CssRuleSet>();

        public List<CssDirective> Directives
        {
            get {
                return m_dirs;
            }
            set {
                m_dirs = value;
            }
        }

        public List<CssRuleSet> RuleSets
        {
            get {
                return m_rulesets;
            }
            set {
                m_rulesets = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (CssDirective cssDir in m_dirs)
            {
                sb.AppendFormat("{0}" + Environment.NewLine, cssDir.ToString());
            }
            if (sb.Length > 0)
            {
                sb.Append(Environment.NewLine);
            }
            foreach (CssRuleSet rules in m_rulesets)
            {
                sb.AppendFormat("{0}" + Environment.NewLine,
                    rules.ToString());
            }
            return sb.ToString();
        }
    }

    public class CssDeclaration
    {
        private string m_name;
        private CssExpression m_expression;
        private bool m_important;

        public string Name
        {
            get {
                return m_name;
            }
            set {
                m_name = value;
            }
        }

        public bool Important
        {
            get {
                return m_important;
            }
            set {
                m_important = value;
            }
        }

        public CssExpression Expression
        {
            get {
                return m_expression;
            }
            set {
                m_expression = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{0}: {1}{2}",
                m_name,
                m_expression.ToString(),
                m_important ? " !important" : "");
            return sb.ToString();
        }
    }

    public class CssDirective : ItfDeclarationContainer, ItfRuleSetContainer
    {
        private CssDirectiveType m_type;
        private string m_name;
        private CssExpression m_expression;
        private List<CssMedium> m_mediums = new List<CssMedium>();
        private List<CssDirective> m_directives = new List<CssDirective>();
        private List<CssRuleSet> m_rulesets = new List<CssRuleSet>();
        private List<CssDeclaration> m_declarations = new List<CssDeclaration>();

        public CssDirectiveType Type
        {
            get {
                return this.m_type;
            }
            set {
                this.m_type = value;
            }
        }

        public string Name
        {
            get {
                return this.m_name;
            }
            set {
                this.m_name = value;
            }
        }

        public CssExpression Expression
        {
            get {
                return this.m_expression;
            }
            set {
                this.m_expression = value;
            }
        }

        public List<CssMedium> Mediums
        {
            get {
                return this.m_mediums;
            }
            set {
                this.m_mediums = value;
            }
        }

        public List<CssDirective> Directives
        {
            get {
                return this.m_directives;
            }
            set {
                this.m_directives = value;
            }
        }

        public List<CssRuleSet> RuleSets
        {
            get {
                return this.m_rulesets;
            }
            set {
                this.m_rulesets = value;
            }
        }

        public List<CssDeclaration> Declarations
        {
            get {
                return this.m_declarations;
            }
            set {
                this.m_declarations = value;
            }
        }

        public override string ToString()
        {
            return ToString(0);
        }

        public string ToString(int indentLevel)
        {
            string start = "".PadRight(indentLevel, '\t');

            switch (m_type)
            {
                case CssDirectiveType.Charset:
                    return ToCharSetString(start);
                case CssDirectiveType.Page:
                    return ToPageString(start);
                case CssDirectiveType.Media:
                    return ToMediaString(indentLevel);
                case CssDirectiveType.Import:
                    return ToImportString();
                case CssDirectiveType.FontFace:
                    return ToFontFaceString(start);
            }

            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("{0} ", m_name);

            if (m_expression != null)
            {
                sb.AppendFormat("{0} ", m_expression);
            }

            bool first = true;
            foreach (CssMedium med in m_mediums)
            {
                if (first)
                {
                    first = false;
                    sb.Append(" ");
                }
                else
                {
                    sb.Append(", ");
                }
                sb.Append(med.ToString());
            }

            bool HasBlock = (this.m_declarations.Count > 0 || this.m_directives.Count > 0 || this.m_rulesets.Count > 0);

            if (!HasBlock)
            {
                sb.Append(";");
                return sb.ToString();
            }

            sb.Append(" {" + Environment.NewLine + start);

            foreach (CssDirective dir in m_directives)
            {
                sb.AppendFormat("{0}" + Environment.NewLine, dir.ToCharSetString(start + "\t"));
            }

            foreach (CssRuleSet rules in m_rulesets)
            {
                sb.AppendFormat("{0}" + Environment.NewLine, rules.ToString(indentLevel + 1));
            }

            first = true;
            foreach (CssDeclaration decl in m_declarations)
            {
                if (first)
                {
                    first = false;
                }
                else
                {
                    sb.Append(";");
                }
                sb.Append(Environment.NewLine + "\t" + start);
                sb.Append(decl.ToString());
            }

            sb.Append(Environment.NewLine + "}");
            return sb.ToString();
        }

        private string ToFontFaceString(string start)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("@font-face {");

            bool first = true;
            foreach (CssDeclaration decl in m_declarations)
            {
                if (first)
                {
                    first = false;
                }
                else
                {
                    sb.Append(";");
                }
                sb.Append(Environment.NewLine + "\t" + start);
                sb.Append(decl.ToString());
            }

            sb.Append(Environment.NewLine + "}");
            return sb.ToString();
        }

        private string ToImportString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("@import ");
            if (m_expression != null)
            {
                sb.AppendFormat("{0} ", m_expression);
            }
            bool first = true;
            foreach (CssMedium med in m_mediums)
            {
                if (first)
                {
                    first = false;
                    sb.Append(" ");
                }
                else
                {
                    sb.Append(", ");
                }
                sb.Append(med.ToString());
            }
            sb.Append(";");
            return sb.ToString();
        }

        private string ToMediaString(int indentLevel)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("@media");

            bool first = true;
            foreach (CssMedium medium in m_mediums)
            {
                if (first)
                {
                    first = false;
                    sb.Append(" ");
                }
                else
                {
                    sb.Append(", ");
                }
                sb.Append(medium.ToString());
            }
            sb.Append(" {" + Environment.NewLine);

            foreach (CssRuleSet ruleset in m_rulesets)
            {
                sb.AppendFormat("{0}" + Environment.NewLine, ruleset.ToString(indentLevel + 1));
            }

            sb.Append("}");
            return sb.ToString();
        }

        private string ToPageString(string start)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("@page ");
            if (m_expression != null)
            {
                sb.AppendFormat("{0} ", m_expression);
            }
            sb.Append("{" + Environment.NewLine);

            bool first = true;
            foreach (CssDeclaration decl in m_declarations)
            {
                if (first)
                {
                    first = false;
                }
                else
                {
                    sb.Append(";");
                }
                sb.Append(Environment.NewLine + "\t" + start);
                sb.Append(decl.ToString());
            }

            sb.Append("}");
            return sb.ToString();
        }

        private string ToCharSetString(string start)
        {
            return string.Format("{2}{0} {1}", 
                m_name, 
                m_expression.ToString(), 
                start);
        }
    }

    public enum CssDirectiveType
    {
        Media,
        Import,
        Charset,
        Page,
        FontFace,
        Namespace,
        Other,
    }

    public class CssExpression
    {
        private List<CssTerm> m_terms = new List<CssTerm>();

        public List<CssTerm> Terms
        {
            get {
                return m_terms;
            }
            set {
                m_terms = value;
            }
        }

        public bool IsNotAuto
        {
            get
            {
                return (this != null && this.ToString() != "auto");
            }
        }

        public bool IsAuto
        {
            get
            {
                return (this != null && this.ToString() == "auto");
            }
        }

        public bool IsNotNormal
        {
            get
            {
                return (this != null && this.ToString() != "normal");
            }
        }

        public bool IsNormal
        {
            get
            {
                return (this != null && this.ToString() == "normal");
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            bool first = true;
            foreach (CssTerm term in m_terms)
            {
                if (first)
                {
                    first = false;
                }
                else
                {
                    sb.AppendFormat("{0} ", 
                        term.Separator.HasValue ? term.Separator.Value.ToString() : "");
                }
                sb.Append(term.ToString());
            }
            return sb.ToString();
        }

        public static implicit operator string(CssExpression e)
        {
            return e.ToString();
        }

        public static explicit operator double(CssExpression e)
        {
            return double.Parse(e.Terms.First().Value);
        }

        public static explicit operator Emu(CssExpression e)
        {
            return Emu.PointsToEmus(double.Parse(e.Terms.First().Value));
        }

        // will only be called on expression that is in terms of points
        public static explicit operator TPoint(CssExpression e)
        {
            return new TPoint(double.Parse(e.Terms.First().Value));
        }

        // will only be called on expression that is in terms of points
        public static explicit operator Twip(CssExpression length)
        {
            if (length.Terms.Count() == 1)
            {
                CssTerm term = length.Terms.First();
                if (term.Unit == CssUnit.PT)
                {
                    double ptValue;
                    if (double.TryParse(term.Value.ToString(), out ptValue))
                    {
                        if (term.Sign == '-')
                            ptValue = -ptValue;
                        return new Twip((long)(ptValue * 20));
                    }
                }
            }
            return 0;
        }
    }

    public class CssFunction
    {
        private string m_name;
        private CssExpression m_expression;

        public string Name
        {
            get {
                return m_name;
            }
            set {
                m_name = value;
            }
        }

        public CssExpression Expression
        {
            get {
                return m_expression;
            }
            set {
                m_expression = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{0}(", m_name);
            if (m_expression != null)
            {
                bool first = true;
                foreach (CssTerm t in m_expression.Terms)
                {
                    if (first)
                    {
                        first = false;
                    }
                    else if (!t.Value.EndsWith("="))
                    {
                        sb.Append(", ");
                    }

                    bool quote = false;
                    if (t.Type == CssTermType.String && !t.Value.EndsWith("="))
                    {
                        quote = true;
                    }
                    if (quote)
                    {
                        sb.Append("'");
                    }
                    sb.Append(t.ToString());
                    if (quote)
                    {
                        sb.Append("'");
                    }
                }
            }
            sb.Append(")");
            return sb.ToString();
        }
    }

    public interface ItfDeclarationContainer
    {
        List<CssDeclaration> Declarations { get; set; }
    }

    public interface ItfRuleSetContainer
    {
        List<CssRuleSet> RuleSets { get; set; }
    }

    public interface ItfSelectorContainer
    {
        List<CssSelector> Selectors { get; set; }
    }

    public enum CssMedium
    {
        all,
        aural,
        braille,
        embossed,
        handheld,
        print,
        projection,
        screen,
        tty,
        tv
    }

    public class CssPropertyValue
    {
        private CssValueType m_type;
        private CssUnit m_unit;
        private string m_value;

        public CssValueType Type
        {
            get {
                return this.m_type;
            }
            set {
                this.m_type = value;
            }
        }

        public CssUnit Unit
        {
            get {
                return this.m_unit;
            }
            set {
                this.m_unit = value;
            }
        }

        public string Value
        {
            get {
                return this.m_value;
            }
            set {
                this.m_value = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder(m_value);
            if (m_type == CssValueType.Unit)
            {
                sb.Append(m_unit.ToString().ToLower());
            }
            sb.Append(" [");
            sb.Append(m_type.ToString());
            sb.Append("]");
            return sb.ToString();
        }

        public bool IsColor
        {
            get
            {
                if (((m_type == CssValueType.Hex) 
                    || (m_type == CssValueType.String && m_value.StartsWith("#"))) 
                    && (m_value.Length == 6 || (m_value.Length == 7 && m_value.StartsWith("#"))))
                {
                    bool hex = true;
                    foreach (char c in m_value)
                    {
                        if (!char.IsDigit(c)
                            && c != '#'
                            && c != 'a'
                            && c != 'A'
                            && c != 'b'
                            && c != 'B'
                            && c != 'c'
                            && c != 'C'
                            && c != 'd'
                            && c != 'D'
                            && c != 'e'
                            && c != 'E'
                            && c != 'f'
                            && c != 'F'
                        )
                        {
                            return false;
                        }
                    }
                    return hex;
                }
                else if (m_type == CssValueType.String)
                {
                    bool number = true;
                    foreach (char c in m_value)
                    {
                        if (!char.IsDigit(c))
                        {
                            number = false;
                            break;
                        }
                    }
                    if (number) { return false; }

                    try
                    {
                        KnownColor kc = (KnownColor)Enum.Parse(typeof(KnownColor), m_value, true);
                        return true;
                    }
                    catch { }
                }
                return false;
            }
        }

        public Color ToColor()
        {
            string hex = "000000";
            if (m_type == CssValueType.Hex)
            {
                if (m_value.Length == 7 && m_value.StartsWith("#"))
                {
                    hex = m_value.Substring(1);
                }
                else if (m_value.Length == 6)
                {
                    hex = m_value;
                }
            }
            else
            {
                try
                {
                    KnownColor kc = (KnownColor)Enum.Parse(typeof(KnownColor), m_value, true);
                    Color c = Color.FromKnownColor(kc);
                    return c;
                }
                catch { }
            }
            int r = ConvertFromHex(hex.Substring(0, 2));
            int g = ConvertFromHex(hex.Substring(2, 2));
            int b = ConvertFromHex(hex.Substring(4));
            return Color.FromArgb(r, g, b);
        }

        private int ConvertFromHex(string input)
        {
            int val;
            int result = 0;
            for (int i = 0; i < input.Length; i++)
            {
                string chunk = input.Substring(i, 1).ToUpper();
                switch (chunk)
                {
                    case "A":
                        val = 10; 
                        break;
                    case "B":
                        val = 11; 
                        break;
                    case "C":
                        val = 12; 
                        break;
                    case "D":
                        val = 13; 
                        break;
                    case "E":
                        val = 14; 
                        break;
                    case "F":
                        val = 15; 
                        break;
                    default:
                        val = int.Parse(chunk); 
                        break;
                }
                if (i == 0)
                {
                    result += val * 16;
                }
                else
                {
                    result += val;
                }
            }
            return result;
        }
    }

    public class CssRuleSet : ItfDeclarationContainer
    {
        private List<CssSelector> m_selectors = new List<CssSelector>();
        private List<CssDeclaration> m_declarations = new List<CssDeclaration>();

        public List<CssSelector> Selectors
        {
            get {
                return m_selectors;
            }
            set {
                m_selectors = value;
            }
        }

        public List<CssDeclaration> Declarations
        {
            get {
                return m_declarations;
            }
            set {
                m_declarations = value;
            }
        }

        public override string ToString()
        {
            return ToString(0);
        }

        public string ToString(int indentLevel)
        {
            string start = "";
            for (int i = 0; i < indentLevel; i++)
            {
                start += "\t";
            }

            StringBuilder sb = new StringBuilder();
            bool first = true;
            foreach (CssSelector sel in m_selectors)
            {
                if (first) 
                { 
                    first = false; 
                    sb.Append(start); 
                } 
                else 
                { 
                    sb.Append(", "); 
                }
                sb.Append(sel.ToString());
            }
            sb.Append(" {" + Environment.NewLine);
            sb.Append(start);

            foreach (CssDeclaration dec in m_declarations)
            {
                sb.AppendFormat("\t{0};" + Environment.NewLine + "{1}", dec.ToString(), start);
            }

            sb.Append("}");
            return sb.ToString();
        }
    }

    public class CssSelector
    {
        private List<CssSimpleSelector> m_simpleSelectors = new List<CssSimpleSelector>();

        public List<CssSimpleSelector> SimpleSelectors
        {
            get {
                return m_simpleSelectors;
            }
            set {
                m_simpleSelectors = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            bool first = true;
            foreach (CssSimpleSelector ss in m_simpleSelectors)
            {
                if (first)
                {
                    first = false;
                }
                else
                {
                    sb.Append(" ");
                }
                sb.Append(ss.ToString());
            }
            return sb.ToString();
        }
    }

    public class CssSimpleSelector
    {
        private CssCombinator? m_combinator = null;
        private string m_elementname;
        private string m_id;
        private string m_cls;
        private CssAttribute m_attribute;
        private string m_pseudo;
        private CssFunction m_function;
        private CssSimpleSelector m_child;

        public CssCombinator? Combinator
        {
            get {
                return m_combinator;
            }
            set {
                m_combinator = value;
            }
        }
        public string CombinatorString
        {
            get {
                if (this.m_combinator.HasValue)
                {
                    return m_combinator.ToString();
                }
                else
                {
                    return null;
                }
            }
            set {
                this.m_combinator = (CssCombinator)Enum.Parse(typeof(CssCombinator), value);
            }
        }

        public string ElementName
        {
            get {
                return m_elementname;
            }
            set {
                m_elementname = value;
            }
        }

        public string ID
        {
            get {
                return m_id;
            }
            set {
                m_id = value;
            }
        }

        public string Class
        {
            get {
                return m_cls;
            }
            set {
                m_cls = value;
            }
        }

        public string Pseudo
        {
            get {
                return m_pseudo;
            }
            set {
                m_pseudo = value;
            }
        }

        public CssAttribute Attribute
        {
            get {
                return m_attribute;
            }
            set {
                m_attribute = value;
            }
        }

        public CssFunction Function
        {
            get {
                return m_function;
            }
            set {
                m_function = value;
            }
        }

        public CssSimpleSelector Child
        {
            get {
                return m_child;
            }
            set {
                m_child = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            if (m_combinator.HasValue)
            {
                switch (m_combinator.Value)
                {
                    case OpenXmlPowerTools.HtmlToWml.CSS.CssCombinator.PrecededImmediatelyBy: 
                        sb.Append(" + "); 
                        break;
                    case OpenXmlPowerTools.HtmlToWml.CSS.CssCombinator.ChildOf: 
                        sb.Append(" > "); 
                        break;
                    case OpenXmlPowerTools.HtmlToWml.CSS.CssCombinator.PrecededBy: 
                        sb.Append(" ~ "); 
                        break;
                }
            }
            if (m_elementname != null)
            {
                sb.Append(m_elementname);
            }
            if (m_id != null)
            {
                sb.AppendFormat("#{0}", m_id);
            }
            if (m_cls != null)
            {
                sb.AppendFormat(".{0}", m_cls);
            }
            if (m_pseudo != null)
            {
                sb.AppendFormat(":{0}", m_pseudo);
            }
            if (m_attribute != null)
            {
                sb.Append(m_attribute.ToString());
            }
            if (m_function != null)
            {
                sb.Append(m_function.ToString());
            }
            if (m_child != null)
            {
                if (m_child.ElementName != null)
                {
                    sb.Append(" ");
                }
                sb.Append(m_child.ToString());
            }
            return sb.ToString();
        }
    }

    public class CssTag
    {
        private CssTagType m_tagtype;
        private string m_name;
        private string m_cls;
        private string m_pseudo;
        private string m_id;
        private char m_parentrel = '\0';
        private CssTag m_subtag;
        private List<string> m_attribs = new List<string>();

        public CssTagType TagType
        {
            get {
                return m_tagtype;
            }
            set {
                m_tagtype = value;
            }
        }

        public bool IsIDSelector
        {
            get {
                return m_id != null;
            }
        }

        public bool HasName
        {
            get {
                return m_name != null;
            }
        }

        public bool HasClass
        {
            get {
                return m_cls != null;
            }
        }

        public bool HasPseudoClass
        {
            get {
                return m_pseudo != null;
            }
        }

        public string Name
        {
            get {
                return m_name;
            }
            set {
                m_name = value;
            }
        }

        public string Class
        {
            get {
                return m_cls;
            }
            set {
                m_cls = value;
            }
        }

        public string Pseudo
        {
            get {
                return m_pseudo;
            }
            set {
                m_pseudo = value;
            }
        }

        public string Id
        {
            get {
                return m_id;
            }
            set {
                m_id = value;
            }
        }

        public char ParentRelationship
        {
            get {
                return m_parentrel;
            }
            set {
                m_parentrel = value;
            }
        }

        public CssTag SubTag
        {
            get {
                return m_subtag;
            }
            set {
                m_subtag = value;
            }
        }

        public List<string> Attributes
        {
            get {
                return m_attribs;
            }
            set {
                m_attribs = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder(ToShortString());

            if (m_subtag != null)
            {
                sb.Append(" ");
                sb.Append(m_subtag.ToString());
            }
            return sb.ToString();
        }

        public string ToShortString()
        {
            StringBuilder sb = new StringBuilder();
            if (m_parentrel != '\0')
            {
                sb.AppendFormat("{0} ", m_parentrel.ToString());
            }
            if (HasName)
            {
                sb.Append(m_name);
            }
            foreach (string atr in m_attribs)
            {
                sb.AppendFormat("[{0}]", atr);
            }
            if (HasClass)
            {
                sb.Append(".");
                sb.Append(m_cls);
            }
            if (IsIDSelector)
            {
                sb.Append("#");
                sb.Append(m_id);
            }
            if (HasPseudoClass)
            {
                sb.Append(":");
                sb.Append(m_pseudo);
            }
            return sb.ToString();
        }
    }

    [Flags]
    public enum CssTagType
    {
        Named = 1,
        Classed = 2,
        IDed = 4,
        Pseudoed = 8,
        Directive = 16
    }

    public class CssTerm
    {
        private char? m_separator;
        private char? m_sign;
        private CssTermType m_type;
        private string m_val;
        private CssUnit? m_unit;
        private CssFunction m_function;

        public char? Separator
        {
            get {
                return m_separator;
            }
            set {
                m_separator = value;
            }
        }
        public string SeparatorChar
        {
            get {
                return m_separator.HasValue ? this.m_separator.Value.ToString() : null;
            }
            set {
                m_separator = !string.IsNullOrEmpty(value) ? value[0] : '\0';
            }
        }

        public char? Sign
        {
            get {
                return m_sign;
            }
            set {
                m_sign = value;
            }
        }
        public string SignChar
        {
            get {
                return this.m_sign.HasValue ? this.m_sign.Value.ToString() : null;
            }
            set {
                this.m_sign = !string.IsNullOrEmpty(value) ? value[0] : '\0';
            }
        }

        public CssTermType Type
        {
            get {
                return m_type;
            }
            set {
                m_type = value;
            }
        }

        public string Value
        {
            get {
                return m_val;
            }
            set {
                m_val = value;
            }
        }

        public CssUnit? Unit
        {
            get {
                return m_unit;
            }
            set {
                m_unit = value;
            }
        }
        public string UnitString
        {
            get {
                if (this.m_unit.HasValue)
                {
                    return this.m_unit.ToString();
                }
                else
                {
                    return null;
                }
            }
            set {
                this.m_unit = (CssUnit)Enum.Parse(typeof(CssUnit), value);
            }
        }

        public CssFunction Function
        {
            get {
                return m_function;
            }
            set {
                m_function = value;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            if (m_type == CssTermType.Function)
            {
                sb.Append(m_function.ToString());
            }
            else if (m_type == CssTermType.Url)
            {
                sb.AppendFormat("url('{0}')", m_val);
            }
            else if (m_type == CssTermType.Unicode)
            {
                sb.AppendFormat("U\\{0}", m_val.ToUpper());
            }
            else if (m_type == CssTermType.Hex)
            {
                sb.Append(m_val.ToUpper());
            }
            else
            {
                if (m_sign.HasValue)
                {
                    sb.Append(m_sign.Value);
                }
                sb.Append(m_val);
                if (m_unit.HasValue)
                {
                    if (m_unit.Value == OpenXmlPowerTools.HtmlToWml.CSS.CssUnit.Percent)
                    {
                        sb.Append("%");
                    }
                    else
                    {
                        sb.Append(CssUnitOutput.ToString(m_unit.Value));
                    }
                }
            }

            return sb.ToString();
        }

        public bool IsColor
        {
            get
            {
                if (((m_type == CssTermType.Hex)
                    || (m_type == CssTermType.String && m_val.StartsWith("#")))
                    && (m_val.Length == 6 || m_val.Length == 3 || ((m_val.Length == 7 || m_val.Length == 4)
                    && m_val.StartsWith("#"))))
                {
                    bool hex = true;
                    foreach (char c in m_val)
                    {
                        if (!char.IsDigit(c)
                            && c != '#'
                            && c != 'a'
                            && c != 'A'
                            && c != 'b'
                            && c != 'B'
                            && c != 'c'
                            && c != 'C'
                            && c != 'd'
                            && c != 'D'
                            && c != 'e'
                            && c != 'E'
                            && c != 'f'
                            && c != 'F'
                        )
                        {
                            return false;
                        }
                    }
                    return hex;
                }
                else if (m_type == CssTermType.String)
                {
                    bool number = true;
                    foreach (char c in m_val)
                    {
                        if (!char.IsDigit(c))
                        {
                            number = false;
                            break;
                        }
                    }
                    if (number) {
                        return false;
                    }

                    KnownColor kc;
                    if (Enum.TryParse(m_val, true, out kc))
                    {
                        return true;
                    }
                }
                else if (m_type == CssTermType.Function)
                {
                    if ((m_function.Name.ToLower().Equals("rgb") && m_function.Expression.Terms.Count == 3)
                        || (m_function.Name.ToLower().Equals("rgba") && m_function.Expression.Terms.Count == 4)
                        )
                    {
                        for (int i = 0; i < m_function.Expression.Terms.Count; i++)
                        {
                            if (m_function.Expression.Terms[i].Type != CssTermType.Number) 
                            { 
                                return false; 
                            }
                        }
                        return true;
                    }
                    else if ((m_function.Name.ToLower().Equals("hsl") && m_function.Expression.Terms.Count == 3)
                      || (m_function.Name.ToLower().Equals("hsla") && m_function.Expression.Terms.Count == 4)
                      )
                    {
                        for (int i = 0; i < m_function.Expression.Terms.Count; i++)
                        {
                            if (m_function.Expression.Terms[i].Type != CssTermType.Number) 
                            { 
                                return false; 
                            }
                        }
                        return true;
                    }
                }
                return false;
            }
        }

        private int GetRGBValue(CssTerm t)
        {
            try
            {
                if (t.Unit.HasValue && t.Unit.Value == OpenXmlPowerTools.HtmlToWml.CSS.CssUnit.Percent)
                {
                    return (int)(255f * float.Parse(t.Value) / 100f);
                }
                return int.Parse(t.Value);
            }
            catch { }
            return 0;
        }

        private int GetHueValue(CssTerm t)
        {
            try
            {
                return (int)(float.Parse(t.Value) * 255f / 360f);
            }
            catch { }
            return 0;
        }

        public Color ToColor()
        {
            string hex = "000000";
            if (m_type == CssTermType.Hex)
            {
                if ((m_val.Length == 7 || m_val.Length == 4) && m_val.StartsWith("#"))
                {
                    hex = m_val.Substring(1);
                }
                else if (m_val.Length == 6 || m_val.Length == 3)
                {
                    hex = m_val;
                }
            }
            else if (m_type == CssTermType.Function)
            {
                if ((m_function.Name.ToLower().Equals("rgb") && m_function.Expression.Terms.Count == 3)
                    || (m_function.Name.ToLower().Equals("rgba") && m_function.Expression.Terms.Count == 4)
                    )
                {
                    int fr = 0, fg = 0, fb = 0;
                    for (int i = 0; i < m_function.Expression.Terms.Count; i++)
                    {
                        if (m_function.Expression.Terms[i].Type != CssTermType.Number) 
                        { 
                            return Color.Black; 
                        }
                        switch (i)
                        {
                            case 0: fr = GetRGBValue(m_function.Expression.Terms[i]); 
                                break;
                            case 1: fg = GetRGBValue(m_function.Expression.Terms[i]); 
                                break;
                            case 2: fb = GetRGBValue(m_function.Expression.Terms[i]); 
                                break;
                        }
                    }
                    return Color.FromArgb(fr, fg, fb);
                }
                else if ((m_function.Name.ToLower().Equals("hsl") && m_function.Expression.Terms.Count == 3)
                  || (m_function.Name.Equals("hsla") && m_function.Expression.Terms.Count == 4)
                  )
                {
                    int h = 0, s = 0, v = 0;
                    for (int i = 0; i < m_function.Expression.Terms.Count; i++)
                    {
                        if (m_function.Expression.Terms[i].Type != CssTermType.Number) { return Color.Black; }
                        switch (i)
                        {
                            case 0: h = GetHueValue(m_function.Expression.Terms[i]); 
                                break;
                            case 1: s = GetRGBValue(m_function.Expression.Terms[i]); 
                                break;
                            case 2: v = GetRGBValue(m_function.Expression.Terms[i]); 
                                break;
                        }
                    }
                    HueSatVal hsv = new HueSatVal(h, s, v);
                    return hsv.Color;
                }
            }
            else
            {
                try
                {
                    KnownColor kc = (KnownColor)Enum.Parse(typeof(KnownColor), m_val, true);
                    Color c = Color.FromKnownColor(kc);
                    return c;
                }
                catch { }
            }
            if (hex.Length == 3)
            {
                string temp = "";
                foreach (char c in hex)
                {
                    temp += c.ToString() + c.ToString();
                }
                hex = temp;
            }
            int r = ConvertFromHex(hex.Substring(0, 2));
            int g = ConvertFromHex(hex.Substring(2, 2));
            int b = ConvertFromHex(hex.Substring(4));
            return Color.FromArgb(r, g, b);
        }
        private int ConvertFromHex(string input)
        {
            int val;
            int result = 0;
            for (int i = 0; i < input.Length; i++)
            {
                string chunk = input.Substring(i, 1).ToUpper();
                switch (chunk)
                {
                    case "A":
                        val = 10; 
                        break;
                    case "B":
                        val = 11; 
                        break;
                    case "C":
                        val = 12; 
                        break;
                    case "D":
                        val = 13; 
                        break;
                    case "E":
                        val = 14; 
                        break;
                    case "F":
                        val = 15; 
                        break;
                    default:
                        val = int.Parse(chunk); 
                        break;
                }
                if (i == 0)
                {
                    result += val * 16;
                }
                else
                {
                    result += val;
                }
            }
            return result;
        }
    }

    public enum CssTermType
    {
        Number,
        Function,
        String,
        Url,
        Unicode,
        Hex
    }

    public enum CssUnit
    {
        None,
        Percent,
        EM,
        EX,
        PX,
        GD,
        REM,
        VW,
        VH,
        VM,
        CH,
        MM,
        CM,
        IN,
        PT,
        PC,
        DEG,
        GRAD,
        RAD,
        TURN,
        MS,
        S,
        Hz,
        kHz,
    }

    public static class CssUnitOutput
    {
        public static string ToString(CssUnit u)
        {
            if (u == CssUnit.Percent)
            {
                return "%";
            }
            else if (u == CssUnit.Hz || u == CssUnit.kHz)
            {
                return u.ToString();
            }
            else if (u == CssUnit.None)
            {
                return "";
            }
            return u.ToString().ToLower();
        }
    }

    public enum CssValueType
    {
        String,
        Hex,
        Unit,
        Percent,
        Url,
        Function
    }

    public class CssParser
    {
        private List<string> m_errors = new List<string>();
        private CssDocument m_doc;

        public CssDocument ParseText(string content)
        {
            MemoryStream mem = new MemoryStream();
            byte[] bytes = ASCIIEncoding.ASCII.GetBytes(content);
            mem.Write(bytes, 0, bytes.Length);
            try
            {
                return ParseStream(mem);
            }
            catch (OpenXmlPowerToolsException e)
            {
                string msg = e.Message + ".  CSS => " + content;
                throw new OpenXmlPowerToolsException(msg);
            }
        }

        // following method should be private, as it does not properly re-throw OpenXmlPowerToolsException
        private CssDocument ParseStream(Stream stream)
        {
            Scanner scanner = new Scanner(stream);
            Parser parser = new Parser(scanner);
            parser.Parse();
            m_doc = parser.CssDoc;
            return m_doc;
        }

        public CssDocument CSSDocument
        {
            get { return m_doc; }
        }

        public List<string> Errors
        {
            get { return m_errors; }
        }
    }

    // Hue Sat and Val values from 0 - 255.
    internal struct HueSatVal
    {
        private int m_hue;
        private int m_sat;
        private int m_val;
        public HueSatVal(int h, int s, int v)
        {
            m_hue = h;
            m_sat = s;
            m_val = v;
        }
        public HueSatVal(Color color)
        {
            m_hue = 0;
            m_sat = 0;
            m_val = 0;
            ConvertFromRGB(color);
        }
        public int Hue
        {
            get {
                return m_hue;
            }
            set {
                m_hue = value;
            }
        }
        public int Saturation
        {
            get {
                return m_sat;
            }
            set {
                m_sat = value;
            }
        }
        public int Value
        {
            get {
                return m_val;
            }
            set {
                m_val = value;
            }
        }
        public Color Color
        {
            get {
                return ConvertToRGB();
            }
            set {
                ConvertFromRGB(value);
            }
        }
        private void ConvertFromRGB(Color color)
        {
            double min; double max; double delta;
            double r = (double)color.R / 255.0d;
            double g = (double)color.G / 255.0d;
            double b = (double)color.B / 255.0d;
            double h; double s; double v;

            min = Math.Min(Math.Min(r, g), b);
            max = Math.Max(Math.Max(r, g), b);
            v = max;
            delta = max - min;
            if (max == 0 || delta == 0)
            {
                s = 0;
                h = 0;
            }
            else
            {
                s = delta / max;
                if (r == max)
                {
                    h = (60D * ((g - b) / delta)) % 360.0d;
                }
                else if (g == max)
                {
                    h = 60D * ((b - r) / delta) + 120.0d;
                }
                else
                {
                    h = 60D * ((r - g) / delta) + 240.0d;
                }
            }
            if (h < 0)
            {
                h += 360.0d;
            }

            Hue = (int)(h / 360.0d * 255.0d);
            Saturation = (int)(s * 255.0d);
            Value = (int)(v * 255.0d);
        }

        private Color ConvertToRGB()
        {
            double h;
            double s;
            double v;
            double r = 0;
            double g = 0;
            double b = 0;

            h = ((double)Hue / 255.0d * 360.0d) % 360.0d;
            s = (double)Saturation / 255.0d;
            v = (double)Value / 255.0d;

            if (s == 0)
            {
                r = v;
                g = v;
                b = v;
            }
            else
            {
                double p;
                double q;
                double t;

                double fractionalPart;
                int sectorNumber;
                double sectorPos;

                sectorPos = h / 60.0d;
                sectorNumber = (int)(Math.Floor(sectorPos));

                fractionalPart = sectorPos - sectorNumber;

                p = v * (1.0d - s);
                q = v * (1.0d - (s * fractionalPart));
                t = v * (1.0d - (s * (1.0d - fractionalPart)));

                switch (sectorNumber)
                {
                    case 0:
                        r = v;
                        g = t;
                        b = p;
                        break;
                    case 1:
                        r = q;
                        g = v;
                        b = p;
                        break;
                    case 2:
                        r = p;
                        g = v;
                        b = t;
                        break;
                    case 3:
                        r = p;
                        g = q;
                        b = v;
                        break;
                    case 4:
                        r = t;
                        g = p;
                        b = v;
                        break;
                    case 5:
                        r = v;
                        g = p;
                        b = q;
                        break;
                }
            }
            return Color.FromArgb((int)(r * 255.0d), (int)(g * 255.0d), (int)(b * 255.0d));
        }

        public static bool operator !=(HueSatVal left, HueSatVal right)
        {
            return !(left == right);
        }

        public static bool operator ==(HueSatVal left, HueSatVal right)
        {
            return (left.Hue == right.Hue && left.Value == right.Value && left.Saturation == right.Saturation);
        }

        public override bool Equals(object obj)
        {
            return this == (HueSatVal)obj;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }

    public class Parser
    {
        public const int c_EOF = 0;
        public const int c_ident = 1;
        public const int c_newline = 2;
        public const int c_digit = 3;
        public const int c_whitespace = 4;
        public const int c_maxT = 49;

        const bool T = true;
        const bool x = false;
        const int minErrDist = 2;

        public Scanner m_scanner;
        public Errors m_errors;

        public CssToken m_lastRecognizedToken;
        public CssToken m_lookaheadToken;
        int errDist = minErrDist;

        public CssDocument CssDoc;

        bool IsInHex(string value)
        {
            if (value.Length == 7)
            {
                return false;
            }
            if (value.Length + m_lookaheadToken.m_tokenValue.Length > 7)
            {
                return false;
            }
            var hexes = new List<string>
            {
                "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "a", "b", "c", "d", "e", "f"
            };
            foreach (char c in m_lookaheadToken.m_tokenValue)
            {
                if (!hexes.Contains(c.ToString()))
                {
                    return false;
                }
            }
            return true;
        }

        bool IsUnitOfLength()
        {
            if (m_lookaheadToken.m_tokenKind != 1)
            {
                return false;
            }
            System.Collections.Generic.List<string> units = new System.Collections.Generic.List<string>(
                new string[]
                {
                    "em", "ex", "px", "gd", "rem", "vw", "vh", "vm", "ch", "mm", "cm", "in", "pt", "pc", "deg", "grad", "rad", "turn", "ms", "s", "hz", "khz"
                });
            return units.Contains(m_lookaheadToken.m_tokenValue.ToLower());
        }

        bool IsNumber()
        {
            if (m_lookaheadToken.m_tokenValue.Length > 0)
            {
                return char.IsDigit(m_lookaheadToken.m_tokenValue[0]);
            }
            return false;
        }

        public Parser(Scanner scanner)
        {
            this.m_scanner = scanner;
            m_errors = new Errors();
        }

        void SyntaxErr(int n)
        {
            if (errDist >= minErrDist) 
                m_errors.SyntaxError(m_lookaheadToken.m_tokenLine, m_lookaheadToken.m_tokenColumn, n);
            errDist = 0;
        }

        public void SemanticErr(string msg)
        {
            if (errDist >= minErrDist)
                m_errors.SemanticError(m_lastRecognizedToken.m_tokenLine, m_lastRecognizedToken.m_tokenColumn, msg);
            errDist = 0;
        }

        void Get()
        {
            for (;;)
            {
                m_lastRecognizedToken = m_lookaheadToken;
                m_lookaheadToken = m_scanner.Scan();
                if (m_lookaheadToken.m_tokenKind <= c_maxT) 
                { 
                    ++errDist; 
                    break; 
                }

                m_lookaheadToken = m_lastRecognizedToken;
            }
        }

        void Expect(int n)
        {
            if (m_lookaheadToken.m_tokenKind == n)
                Get();
            else
            {
                SyntaxErr(n);
            }
        }

        bool StartOf(int s)
        {
            return set[s, m_lookaheadToken.m_tokenKind];
        }

        void ExpectWeak(int n, int follow)
        {
            if (m_lookaheadToken.m_tokenKind == n) 
                Get();
            else
            {
                SyntaxErr(n);
                while (!StartOf(follow)) 
                    Get();
            }
        }


        bool WeakSeparator(int n, int syFol, int repFol)
        {
            int kind = m_lookaheadToken.m_tokenKind;
            if (kind == n) 
            { 
                Get(); 
                return true; 
            }
            else if (StartOf(repFol)) 
            { 
                return false; 
            }
            else
            {
                SyntaxErr(n);
                while (!(set[syFol, kind] || set[repFol, kind] || set[0, kind]))
                {
                    Get();
                    kind = m_lookaheadToken.m_tokenKind;
                }
                return StartOf(syFol);
            }
        }


        void Css3()
        {
            CssDoc = new CssDocument();
            CssRuleSet rset = null;
            CssDirective dir = null;

            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            while (m_lookaheadToken.m_tokenKind == 5 || m_lookaheadToken.m_tokenKind == 6)
            {
                if (m_lookaheadToken.m_tokenKind == 5)
                {
                    Get();
                }
                else
                {
                    Get();
                }
            }
            while (StartOf(1))
            {
                if (StartOf(2))
                {
                    RuleSet(out rset);
                    CssDoc.RuleSets.Add(rset);
                }
                else
                {
                    Directive(out dir);
                    CssDoc.Directives.Add(dir);
                }
                while (m_lookaheadToken.m_tokenKind == 5 || m_lookaheadToken.m_tokenKind == 6)
                {
                    if (m_lookaheadToken.m_tokenKind == 5)
                    {
                        Get();
                    }
                    else
                    {
                        Get();
                    }
                }
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
            }
        }

        void RuleSet(out CssRuleSet rset)
        {
            rset = new CssRuleSet();
            CssSelector sel = null;
            CssDeclaration dec = null;

            Selector(out sel);
            rset.Selectors.Add(sel);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            while (m_lookaheadToken.m_tokenKind == 25)
            {
                Get();
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
                Selector(out sel);
                rset.Selectors.Add(sel);
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
            }
            Expect(26);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            if (StartOf(3))
            {
                Declaration(out dec);
                rset.Declarations.Add(dec);
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
                while (m_lookaheadToken.m_tokenKind == 27)
                {
                    Get();
                    while (m_lookaheadToken.m_tokenKind == 4)
                    {
                        Get();
                    }
                    if (m_lookaheadToken.m_tokenValue.Equals("}")) 
                    { 
                        Get(); 
                        return; 
                    }

                    Declaration(out dec);
                    rset.Declarations.Add(dec);
                    while (m_lookaheadToken.m_tokenKind == 4)
                    {
                        Get();
                    }
                }
                if (m_lookaheadToken.m_tokenKind == 27)
                {
                    Get();
                    while (m_lookaheadToken.m_tokenKind == 4)
                    {
                        Get();
                    }
                }
            }
            Expect(28);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
        }

        void Directive(out CssDirective dir)
        {
            dir = new CssDirective();
            CssDeclaration dec = null;
            CssRuleSet rset = null;
            CssExpression exp = null;
            CssDirective dr = null;
            string ident = null;
            CssMedium m;

            Expect(23);
            dir.Name = "@";
            if (m_lookaheadToken.m_tokenKind == 24)
            {
                Get();
                dir.Name += "-";
            }
            Identity(out ident);
            dir.Name += ident;
            switch (dir.Name.ToLower())
            {
                case "@media": 
                    dir.Type = CssDirectiveType.Media; 
                    break;
                case "@import": 
                    dir.Type = CssDirectiveType.Import; 
                    break;
                case "@charset": 
                    dir.Type = CssDirectiveType.Charset; 
                    break;
                case "@page": 
                    dir.Type = CssDirectiveType.Page; 
                    break;
                case "@font-face": 
                    dir.Type = CssDirectiveType.FontFace; 
                    break;
                case "@namespace": 
                    dir.Type = CssDirectiveType.Namespace; 
                    break;
                default: 
                    dir.Type = CssDirectiveType.Other; 
                    break;
            }

            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            if (StartOf(4))
            {
                if (StartOf(5))
                {
                    Medium(out m);
                    dir.Mediums.Add(m);
                    while (m_lookaheadToken.m_tokenKind == 4)
                    {
                        Get();
                    }
                    while (m_lookaheadToken.m_tokenKind == 25)
                    {
                        Get();
                        while (m_lookaheadToken.m_tokenKind == 4)
                        {
                            Get();
                        }
                        Medium(out m);
                        dir.Mediums.Add(m);
                        while (m_lookaheadToken.m_tokenKind == 4)
                        {
                            Get();
                        }
                    }
                }
                else
                {
                    Exprsn(out exp);
                    dir.Expression = exp;
                    while (m_lookaheadToken.m_tokenKind == 4)
                    {
                        Get();
                    }
                }
            }
            if (m_lookaheadToken.m_tokenKind == 26)
            {
                Get();
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
                if (StartOf(6))
                {
                    while (StartOf(1))
                    {
                        if (dir.Type == CssDirectiveType.Page || dir.Type == CssDirectiveType.FontFace)
                        {
                            Declaration(out dec);
                            dir.Declarations.Add(dec);
                            while (m_lookaheadToken.m_tokenKind == 4)
                            {
                                Get();
                            }
                            while (m_lookaheadToken.m_tokenKind == 27)
                            {
                                Get();
                                while (m_lookaheadToken.m_tokenKind == 4)
                                {
                                    Get();
                                }
                                if (m_lookaheadToken.m_tokenValue.Equals("}")) 
                                { 
                                    Get(); 
                                    return; 
                                }
                                Declaration(out dec);
                                dir.Declarations.Add(dec);
                                while (m_lookaheadToken.m_tokenKind == 4)
                                {
                                    Get();
                                }
                            }
                            if (m_lookaheadToken.m_tokenKind == 27)
                            {
                                Get();
                                while (m_lookaheadToken.m_tokenKind == 4)
                                {
                                    Get();
                                }
                            }
                        }
                        else if (StartOf(2))
                        {
                            RuleSet(out rset);
                            dir.RuleSets.Add(rset);
                            while (m_lookaheadToken.m_tokenKind == 4)
                            {
                                Get();
                            }
                        }
                        else
                        {
                            Directive(out dr);
                            dir.Directives.Add(dr);
                            while (m_lookaheadToken.m_tokenKind == 4)
                            {
                                Get();
                            }
                        }
                    }
                }
                Expect(28);
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
            }
            else if (m_lookaheadToken.m_tokenKind == 27)
            {
                Get();
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
            }
            else SyntaxErr(50);
        }

        void QuotedString(out string qs)
        {
            qs = "";
            if (m_lookaheadToken.m_tokenKind == 7)
            {
                Get();
                while (StartOf(7))
                {
                    Get();
                    qs += m_lastRecognizedToken.m_tokenValue;
                    if (m_lookaheadToken.m_tokenValue.Equals("'") && !m_lastRecognizedToken.m_tokenValue.Equals("\\")) 
                    { 
                        break; 
                    }
                }
                Expect(7);
            }
            else if (m_lookaheadToken.m_tokenKind == 8)
            {
                Get();
                while (StartOf(8))
                {
                    Get();
                    qs += m_lastRecognizedToken.m_tokenValue;
                    if (m_lookaheadToken.m_tokenValue.Equals("\"") && !m_lastRecognizedToken.m_tokenValue.Equals("\\")) 
                    { 
                        break; 
                    }
                }
                Expect(8);
            }
            else SyntaxErr(51);

        }

        void URI(out string url)
        {
            url = "";
            Expect(9);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            if (m_lookaheadToken.m_tokenKind == 10)
            {
                Get();
            }
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            if (m_lookaheadToken.m_tokenKind == 7 || m_lookaheadToken.m_tokenKind == 8)
            {
                QuotedString(out url);
            }
            else if (StartOf(9))
            {
                while (StartOf(10))
                {
                    Get();
                    url += m_lastRecognizedToken.m_tokenValue;
                    if (m_lookaheadToken.m_tokenValue.Equals(")")) 
                    { 
                        break; 
                    }
                }
            }
            else SyntaxErr(52);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            if (m_lookaheadToken.m_tokenKind == 11)
            {
                Get();
            }
        }

        void Medium(out CssMedium m)
        {
            m = CssMedium.all;
            switch (m_lookaheadToken.m_tokenKind)
            {
                case 12:
                    {
                        Get();
                        m = CssMedium.all;
                        break;
                    }
                case 13:
                    {
                        Get();
                        m = CssMedium.aural;
                        break;
                    }
                case 14:
                    {
                        Get();
                        m = CssMedium.braille;
                        break;
                    }
                case 15:
                    {
                        Get();
                        m = CssMedium.embossed;
                        break;
                    }
                case 16:
                    {
                        Get();
                        m = CssMedium.handheld;
                        break;
                    }
                case 17:
                    {
                        Get();
                        m = CssMedium.print;
                        break;
                    }
                case 18:
                    {
                        Get();
                        m = CssMedium.projection;
                        break;
                    }
                case 19:
                    {
                        Get();
                        m = CssMedium.screen;
                        break;
                    }
                case 20:
                    {
                        Get();
                        m = CssMedium.tty;
                        break;
                    }
                case 21:
                    {
                        Get();
                        m = CssMedium.tv;
                        break;
                    }
                default: SyntaxErr(53); break;
            }
        }

        void Identity(out string ident)
        {
            ident = "";
            switch (m_lookaheadToken.m_tokenKind)
            {
                case 1:
                    {
                        Get();
                        break;
                    }
                case 22:
                    {
                        Get();
                        break;
                    }
                case 9:
                    {
                        Get();
                        break;
                    }
                case 12:
                    {
                        Get();
                        break;
                    }
                case 13:
                    {
                        Get();
                        break;
                    }
                case 14:
                    {
                        Get();
                        break;
                    }
                case 15:
                    {
                        Get();
                        break;
                    }
                case 16:
                    {
                        Get();
                        break;
                    }
                case 17:
                    {
                        Get();
                        break;
                    }
                case 18:
                    {
                        Get();
                        break;
                    }
                case 19:
                    {
                        Get();
                        break;
                    }
                case 20:
                    {
                        Get();
                        break;
                    }
                case 21:
                    {
                        Get();
                        break;
                    }
                default: SyntaxErr(54); break;
            }
            ident += m_lastRecognizedToken.m_tokenValue;
        }

        void Exprsn(out CssExpression exp)
        {
            exp = new CssExpression();
            char? sep = null;
            CssTerm trm = null;

            Term(out trm);
            exp.Terms.Add(trm);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            while (StartOf(11))
            {
                if (m_lookaheadToken.m_tokenKind == 25 || m_lookaheadToken.m_tokenKind == 46)
                {
                    if (m_lookaheadToken.m_tokenKind == 46)
                    {
                        Get();
                        sep = '/';
                    }
                    else
                    {
                        Get();
                        sep = ',';
                    }
                    while (m_lookaheadToken.m_tokenKind == 4)
                    {
                        Get();
                    }
                }
                Term(out trm);
                if (sep.HasValue) 
                { 
                    trm.Separator = sep.Value; 
                }
                exp.Terms.Add(trm);
                sep = null;

                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
            }
        }

        void Declaration(out CssDeclaration dec)
        {
            dec = new CssDeclaration();
            CssExpression exp = null;
            string ident = "";

            if (m_lookaheadToken.m_tokenKind == 24)
            {
                Get();
                dec.Name += "-";
            }
            Identity(out ident);
            dec.Name += ident;
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            Expect(43);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            Exprsn(out exp);
            dec.Expression = exp;
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            if (m_lookaheadToken.m_tokenKind == 44)
            {
                Get();
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
                Expect(45);
                dec.Important = true;
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
            }
        }

        void Selector(out CssSelector sel)
        {
            sel = new CssSelector();
            CssSimpleSelector ss = null;
            CssCombinator? cb = null;

            SimpleSelector(out ss);
            sel.SimpleSelectors.Add(ss);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            while (StartOf(12))
            {
                if (m_lookaheadToken.m_tokenKind == 29 || m_lookaheadToken.m_tokenKind == 30 || m_lookaheadToken.m_tokenKind == 31)
                {
                    if (m_lookaheadToken.m_tokenKind == 29)
                    {
                        Get();
                        cb = CssCombinator.PrecededImmediatelyBy;
                    }
                    else if (m_lookaheadToken.m_tokenKind == 30)
                    {
                        Get();
                        cb = CssCombinator.ChildOf;
                    }
                    else
                    {
                        Get();
                        cb = CssCombinator.PrecededBy;
                    }
                }
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
                SimpleSelector(out ss);
                if (cb.HasValue) 
                { 
                    ss.Combinator = cb.Value; 
                }
                sel.SimpleSelectors.Add(ss);

                cb = null;
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
            }
        }

        void SimpleSelector(out CssSimpleSelector ss)
        {
            ss = new CssSimpleSelector();
            ss.ElementName = "";
            string psd = null;
            OpenXmlPowerTools.HtmlToWml.CSS.CssAttribute atb = null;
            CssSimpleSelector parent = ss;
            string ident = null;

            if (StartOf(3))
            {
                if (m_lookaheadToken.m_tokenKind == 24)
                {
                    Get();
                    ss.ElementName += "-";
                }
                Identity(out ident);
                ss.ElementName += ident;
            }
            else if (m_lookaheadToken.m_tokenKind == 32)
            {
                Get();
                ss.ElementName = "*";
            }
            else if (StartOf(13))
            {
                if (m_lookaheadToken.m_tokenKind == 33)
                {
                    Get();
                    if (m_lookaheadToken.m_tokenKind == 24)
                    {
                        Get();
                        ss.ID = "-";
                    }
                    Identity(out ident);
                    if (ss.ID == null) 
                    { 
                        ss.ID = ident; 
                    } 
                    else 
                    { 
                        ss.ID += ident; 
                    }
                }
                else if (m_lookaheadToken.m_tokenKind == 34)
                {
                    Get();
                    ss.Class = "";
                    if (m_lookaheadToken.m_tokenKind == 24)
                    {
                        Get();
                        ss.Class += "-";
                    }
                    Identity(out ident);
                    ss.Class += ident;
                }
                else if (m_lookaheadToken.m_tokenKind == 35)
                {
                    Attrib(out atb);
                    ss.Attribute = atb;
                }
                else
                {
                    Pseudo(out psd);
                    ss.Pseudo = psd;
                }
            }
            else SyntaxErr(55);
            while (StartOf(13))
            {
                CssSimpleSelector child = new CssSimpleSelector();
                if (m_lookaheadToken.m_tokenKind == 33)
                {
                    Get();
                    if (m_lookaheadToken.m_tokenKind == 24)
                    {
                        Get();
                        child.ID = "-";
                    }
                    Identity(out ident);
                    if (child.ID == null)
                    {
                        child.ID = ident;
                    }
                    else
                    {
                        child.ID += "-";
                    }
                }
                else if (m_lookaheadToken.m_tokenKind == 34)
                {
                    Get();
                    child.Class = "";
                    if (m_lookaheadToken.m_tokenKind == 24)
                    {
                        Get();
                        child.Class += "-";
                    }
                    Identity(out ident);
                    child.Class += ident;
                }
                else if (m_lookaheadToken.m_tokenKind == 35)
                {
                    Attrib(out atb);
                    child.Attribute = atb;
                }
                else
                {
                    Pseudo(out psd);
                    child.Pseudo = psd;
                }
                parent.Child = child;
                parent = child;

            }
        }

        void Attrib(out OpenXmlPowerTools.HtmlToWml.CSS.CssAttribute atb)
        {
            atb = new OpenXmlPowerTools.HtmlToWml.CSS.CssAttribute();
            atb.Value = "";
            string quote = null;
            string ident = null;

            Expect(35);
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            Identity(out ident);
            atb.Operand = ident;
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            if (StartOf(14))
            {
                switch (m_lookaheadToken.m_tokenKind)
                {
                    case 36:
                        {
                            Get();
                            atb.Operator = CssAttributeOperator.Equals;
                            break;
                        }
                    case 37:
                        {
                            Get();
                            atb.Operator = CssAttributeOperator.InList;
                            break;
                        }
                    case 38:
                        {
                            Get();
                            atb.Operator = CssAttributeOperator.Hyphenated;
                            break;
                        }
                    case 39:
                        {
                            Get();
                            atb.Operator = CssAttributeOperator.EndsWith;
                            break;
                        }
                    case 40:
                        {
                            Get();
                            atb.Operator = CssAttributeOperator.BeginsWith;
                            break;
                        }
                    case 41:
                        {
                            Get();
                            atb.Operator = CssAttributeOperator.Contains;
                            break;
                        }
                }
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
                if (StartOf(3))
                {
                    if (m_lookaheadToken.m_tokenKind == 24)
                    {
                        Get();
                        atb.Value += "-";
                    }
                    Identity(out ident);
                    atb.Value += ident;
                }
                else if (m_lookaheadToken.m_tokenKind == 7 || m_lookaheadToken.m_tokenKind == 8)
                {
                    QuotedString(out quote);
                    atb.Value = quote;
                }
                else SyntaxErr(56);
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
            }
            Expect(42);
        }

        void Pseudo(out string pseudo)
        {
            pseudo = "";
            CssExpression exp = null;
            string ident = null;

            Expect(43);
            if (m_lookaheadToken.m_tokenKind == 43)
            {
                Get();
            }
            while (m_lookaheadToken.m_tokenKind == 4)
            {
                Get();
            }
            if (m_lookaheadToken.m_tokenKind == 24)
            {
                Get();
                pseudo += "-";
            }
            Identity(out ident);
            pseudo += ident;
            if (m_lookaheadToken.m_tokenKind == 10)
            {
                Get();
                pseudo += m_lastRecognizedToken.m_tokenValue;
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
                Exprsn(out exp);
                pseudo += exp.ToString();
                while (m_lookaheadToken.m_tokenKind == 4)
                {
                    Get();
                }
                Expect(11);
                pseudo += m_lastRecognizedToken.m_tokenValue;
            }
        }

        void Term(out CssTerm trm)
        {
            trm = new CssTerm();
            string val = "";
            CssExpression exp = null;
            string ident = null;

            if (m_lookaheadToken.m_tokenKind == 7 || m_lookaheadToken.m_tokenKind == 8)
            {
                QuotedString(out val);
                trm.Value = val; trm.Type = CssTermType.String;
            }
            else if (m_lookaheadToken.m_tokenKind == 9)
            {
                URI(out val);
                trm.Value = val;
                trm.Type = CssTermType.Url;
            }
            else if (m_lookaheadToken.m_tokenKind == 47)
            {
                Get();
                Identity(out ident);
                trm.Value = "U\\" + ident;
                trm.Type = CssTermType.Unicode;
            }
            else if (m_lookaheadToken.m_tokenKind == 33)
            {
                HexValue(out val);
                trm.Value = val;
                trm.Type = CssTermType.Hex;
            }
            else if (StartOf(15))
            {
                bool minus = false;
                if (m_lookaheadToken.m_tokenKind == 24)
                {
                    Get();
                    minus = true;
                }
                if (StartOf(16))
                {
                    Identity(out ident);
                    trm.Value = ident;
                    trm.Type = CssTermType.String;
                    if (minus)
                    {
                        trm.Value = "-" + trm.Value;
                    }
                    if (StartOf(17))
                    {
                        while (m_lookaheadToken.m_tokenKind == 34 || m_lookaheadToken.m_tokenKind == 36 || m_lookaheadToken.m_tokenKind == 43)
                        {
                            if (m_lookaheadToken.m_tokenKind == 43)
                            {
                                Get();
                                trm.Value += m_lastRecognizedToken.m_tokenValue;
                                if (StartOf(18))
                                {
                                    if (m_lookaheadToken.m_tokenKind == 43)
                                    {
                                        Get();
                                        trm.Value += m_lastRecognizedToken.m_tokenValue;
                                    }
                                    if (m_lookaheadToken.m_tokenKind == 24)
                                    {
                                        Get();
                                        trm.Value += m_lastRecognizedToken.m_tokenValue;
                                    }
                                    Identity(out ident);
                                    trm.Value += ident;
                                }
                                else if (m_lookaheadToken.m_tokenKind == 33)
                                {
                                    HexValue(out val);
                                    trm.Value += val;
                                }
                                else if (StartOf(19))
                                {
                                    while (m_lookaheadToken.m_tokenKind == 3)
                                    {
                                        Get();
                                        trm.Value += m_lastRecognizedToken.m_tokenValue;
                                    }
                                    if (m_lookaheadToken.m_tokenKind == 34)
                                    {
                                        Get();
                                        trm.Value += ".";
                                        while (m_lookaheadToken.m_tokenKind == 3)
                                        {
                                            Get();
                                            trm.Value += m_lastRecognizedToken.m_tokenValue;
                                        }
                                    }
                                }
                                else SyntaxErr(57);
                            }
                            else if (m_lookaheadToken.m_tokenKind == 34)
                            {
                                Get();
                                trm.Value += m_lastRecognizedToken.m_tokenValue;
                                if (m_lookaheadToken.m_tokenKind == 24)
                                {
                                    Get();
                                    trm.Value += m_lastRecognizedToken.m_tokenValue;
                                }
                                Identity(out ident);
                                trm.Value += ident;
                            }
                            else
                            {
                                Get();
                                trm.Value += m_lastRecognizedToken.m_tokenValue;
                                if (m_lookaheadToken.m_tokenKind == 24)
                                {
                                    Get();
                                    trm.Value += m_lastRecognizedToken.m_tokenValue;
                                }
                                if (StartOf(16))
                                {
                                    Identity(out ident);
                                    trm.Value += ident;
                                }
                                else if (StartOf(19))
                                {
                                    while (m_lookaheadToken.m_tokenKind == 3)
                                    {
                                        Get();
                                        trm.Value += m_lastRecognizedToken.m_tokenValue;
                                    }
                                }
                                else SyntaxErr(58);
                            }
                        }
                    }
                    if (m_lookaheadToken.m_tokenKind == 10)
                    {
                        Get();
                        while (m_lookaheadToken.m_tokenKind == 4)
                        {
                            Get();
                        }
                        Exprsn(out exp);
                        CssFunction func = new CssFunction();
                        func.Name = trm.Value;
                        func.Expression = exp;
                        trm.Value = null;
                        trm.Function = func;
                        trm.Type = CssTermType.Function;

                        while (m_lookaheadToken.m_tokenKind == 4)
                        {
                            Get();
                        }
                        Expect(11);
                    }
                }
                else if (StartOf(15))
                {
                    if (m_lookaheadToken.m_tokenKind == 29)
                    {
                        Get();
                        trm.Sign = '+';
                    }
                    if (minus) { trm.Sign = '-'; }
                    while (m_lookaheadToken.m_tokenKind == 3)
                    {
                        Get();
                        val += m_lastRecognizedToken.m_tokenValue;
                    }
                    if (m_lookaheadToken.m_tokenKind == 34)
                    {
                        Get();
                        val += m_lastRecognizedToken.m_tokenValue;
                        while (m_lookaheadToken.m_tokenKind == 3)
                        {
                            Get();
                            val += m_lastRecognizedToken.m_tokenValue;
                        }
                    }
                    if (StartOf(20))
                    {
                        if (m_lookaheadToken.m_tokenValue.ToLower().Equals("n"))
                        {
                            Expect(22);
                            val += m_lastRecognizedToken.m_tokenValue;
                            if (m_lookaheadToken.m_tokenKind == 24 || m_lookaheadToken.m_tokenKind == 29)
                            {
                                if (m_lookaheadToken.m_tokenKind == 29)
                                {
                                    Get();
                                    val += m_lastRecognizedToken.m_tokenValue;
                                }
                                else
                                {
                                    Get();
                                    val += m_lastRecognizedToken.m_tokenValue;
                                }
                                Expect(3);
                                val += m_lastRecognizedToken.m_tokenValue;
                                while (m_lookaheadToken.m_tokenKind == 3)
                                {
                                    Get();
                                    val += m_lastRecognizedToken.m_tokenValue;
                                }
                            }
                        }
                        else if (m_lookaheadToken.m_tokenKind == 48)
                        {
                            Get();
                            trm.Unit = CssUnit.Percent;
                        }
                        else
                        {
                            if (IsUnitOfLength())
                            {
                                Identity(out ident);
                                try
                                {
                                    trm.Unit = (CssUnit)Enum.Parse(typeof(CssUnit), ident, true);
                                }
                                catch
                                {
                                    m_errors.SemanticError(m_lastRecognizedToken.m_tokenLine, m_lastRecognizedToken.m_tokenColumn, string.Format("Unrecognized unit '{0}'", ident));
                                }

                            }
                        }
                    }
                    trm.Value = val; trm.Type = CssTermType.Number;
                }
                else SyntaxErr(59);
            }
            else SyntaxErr(60);
        }

        void HexValue(out string val)
        {
            val = "";
            bool found = false;

            Expect(33);
            val += m_lastRecognizedToken.m_tokenValue;
            if (StartOf(19))
            {
                while (m_lookaheadToken.m_tokenKind == 3)
                {
                    Get();
                    val += m_lastRecognizedToken.m_tokenValue;
                }
            }
            else if (IsInHex(val))
            {
                Expect(1);
                val += m_lastRecognizedToken.m_tokenValue; found = true;
            }
            else SyntaxErr(61);
            if (!found && IsInHex(val))
            {
                Expect(1);
                val += m_lastRecognizedToken.m_tokenValue;
            }
        }

        public void Parse()
        {
            m_lookaheadToken = new CssToken();
            m_lookaheadToken.m_tokenValue = "";
            Get();
            Css3();
            Expect(0);
        }

        static readonly bool[,] set = {
		    {T,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x},
		    {x,T,x,x, x,x,x,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,T, T,x,x,x, x,x,x,x, T,T,T,T, x,x,x,x, x,x,x,T, x,x,x,x, x,x,x},
		    {x,T,x,x, x,x,x,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,x, T,x,x,x, x,x,x,x, T,T,T,T, x,x,x,x, x,x,x,T, x,x,x,x, x,x,x},
		    {x,T,x,x, x,x,x,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,x, T,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x},
		    {x,T,x,T, T,x,x,T, T,T,x,x, T,T,T,T, T,T,T,T, T,T,T,x, T,T,T,T, x,T,x,x, x,T,T,x, x,x,x,x, x,x,x,x, x,x,T,T, T,x,x},
		    {x,x,x,x, x,x,x,x, x,x,x,x, T,T,T,T, T,T,T,T, T,T,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x},
		    {x,T,x,x, x,x,x,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,T, T,x,x,x, T,x,x,x, T,T,T,T, x,x,x,x, x,x,x,T, x,x,x,x, x,x,x},
		    {x,T,T,T, T,T,T,x, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,x},
		    {x,T,T,T, T,T,T,T, x,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,x},
		    {x,T,T,T, T,T,T,T, T,T,x,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,x},
		    {x,T,T,T, x,T,T,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,x},
		    {x,T,x,T, T,x,x,T, T,T,x,x, T,T,T,T, T,T,T,T, T,T,T,x, T,T,x,x, x,T,x,x, x,T,T,x, x,x,x,x, x,x,x,x, x,x,T,T, T,x,x},
		    {x,T,x,x, T,x,x,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,x, T,x,x,x, x,T,T,T, T,T,T,T, x,x,x,x, x,x,x,T, x,x,x,x, x,x,x},
		    {x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,T,T,T, x,x,x,x, x,x,x,T, x,x,x,x, x,x,x},
		    {x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, T,T,T,T, T,T,x,x, x,x,x,x, x,x,x},
		    {x,T,x,T, T,x,x,T, T,T,x,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,x,x, T,T,T,T, x,x,x,x, x,x,x,T, T,x,T,T, T,x,x},
		    {x,T,x,x, x,x,x,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x},
		    {x,x,x,x, x,x,x,x, x,x,T,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,T,x, T,x,x,x, x,x,x,T, x,x,x,x, x,x,x},
		    {x,T,x,x, x,x,x,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,x, T,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,T, x,x,x,x, x,x,x},
		    {x,T,x,T, T,x,x,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,T,T, T,T,x,x, T,T,T,T, T,x,x,x, x,x,x,T, T,x,T,T, T,x,x},
		    {x,T,x,x, x,x,x,x, x,T,x,x, T,T,T,T, T,T,T,T, T,T,T,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, x,x,x,x, T,x,x}
	    };
    }


    public class Errors
    {
        public int numberOfErrorsDetected = 0;
        public string errMsgFormat = "CssParser error: line {0} col {1}: {2}";

        public virtual void SyntaxError(int line, int col, int n)
        {
            string s;
            switch (n)
            {
                case 0: s = "EOF expected";
                    break;
                case 1: s = "identifier expected";
                    break;
                case 2: s = "newline expected";
                    break;
                case 3: s = "digit expected";
                    break;
                case 4: s = "whitespace expected";
                    break;
                case 5: s = "\"<!--\" expected";
                    break;
                case 6: s = "\"-->\" expected";
                    break;
                case 7: s = "\"\'\" expected";
                    break;
                case 8: s = "\"\"\" expected";
                    break;
                case 9: s = "\"url\" expected";
                    break;
                case 10: s = "\"(\" expected";
                    break;
                case 11: s = "\")\" expected";
                    break;
                case 12: s = "\"all\" expected";
                    break;
                case 13: s = "\"aural\" expected";
                    break;
                case 14: s = "\"braille\" expected";
                    break;
                case 15: s = "\"embossed\" expected";
                    break;
                case 16: s = "\"handheld\" expected";
                    break;
                case 17: s = "\"print\" expected";
                    break;
                case 18: s = "\"projection\" expected";
                    break;
                case 19: s = "\"screen\" expected";
                    break;
                case 20: s = "\"tty\" expected";
                    break;
                case 21: s = "\"tv\" expected";
                    break;
                case 22: s = "\"n\" expected";
                    break;
                case 23: s = "\"@\" expected";
                    break;
                case 24: s = "\"-\" expected";
                    break;
                case 25: s = "\",\" expected";
                    break;
                case 26: s = "\"{\" expected";
                    break;
                case 27: s = "\";\" expected";
                    break;
                case 28: s = "\"}\" expected";
                    break;
                case 29: s = "\"+\" expected";
                    break;
                case 30: s = "\">\" expected";
                    break;
                case 31: s = "\"~\" expected";
                    break;
                case 32: s = "\"*\" expected";
                    break;
                case 33: s = "\"#\" expected";
                    break;
                case 34: s = "\".\" expected";
                    break;
                case 35: s = "\"[\" expected";
                    break;
                case 36: s = "\"=\" expected";
                    break;
                case 37: s = "\"~=\" expected";
                    break;
                case 38: s = "\"|=\" expected";
                    break;
                case 39: s = "\"$=\" expected";
                    break;
                case 40: s = "\"^=\" expected";
                    break;
                case 41: s = "\"*=\" expected";
                    break;
                case 42: s = "\"]\" expected";
                    break;
                case 43: s = "\":\" expected";
                    break;
                case 44: s = "\"!\" expected";
                    break;
                case 45: s = "\"important\" expected";
                    break;
                case 46: s = "\"/\" expected";
                    break;
                case 47: s = "\"U\\\\\" expected";
                    break;
                case 48: s = "\"%\" expected";
                    break;
                case 49: s = "??? expected";
                    break;
                case 50: s = "invalid directive";
                    break;
                case 51: s = "invalid QuotedString";
                    break;
                case 52: s = "invalid URI";
                    break;
                case 53: s = "invalid medium";
                    break;
                case 54: s = "invalid identity";
                    break;
                case 55: s = "invalid simpleselector";
                    break;
                case 56: s = "invalid attrib";
                    break;
                case 57: s = "invalid term";
                    break;
                case 58: s = "invalid term";
                    break;
                case 59: s = "invalid term";
                    break;
                case 60: s = "invalid term";
                    break;
                case 61: s = "invalid HexValue";
                    break;

                default: s = "error " + n;
                    break;
            }
            var errorString = string.Format(errMsgFormat, line, col, s);
            throw new OpenXmlPowerToolsException(errorString);
        }

        public virtual void SemanticError(int line, int col, string s)
        {
            var errorString = string.Format(errMsgFormat, line, col, s);
            throw new OpenXmlPowerToolsException(errorString);
        }

        public virtual void SemanticError(string s)
        {
            throw new OpenXmlPowerToolsException(s);
        }

        public virtual void Warning(int line, int col, string s)
        {
            var errorString = string.Format(errMsgFormat, line, col, s);
            throw new OpenXmlPowerToolsException(errorString);
        }

        public virtual void Warning(string s)
        {
            throw new OpenXmlPowerToolsException(s);
        }
    }


    public class FatalError : Exception
    {
        public FatalError(string m) : base(m) { }
    }

    public class CssToken
    {
        public int m_tokenKind;
        public int m_tokenPositionInBytes;
        public int m_tokenPositionInCharacters;
        public int m_tokenColumn;
        public int m_tokenLine;
        public string m_tokenValue;
        public CssToken m_nextToken;
    }

    public class CssBuffer
    {
        public const int EOF = char.MaxValue + 1;
        const int MIN_BUFFER_LENGTH = 1024;
        const int MAX_BUFFER_LENGTH = MIN_BUFFER_LENGTH * 64;
        byte[] m_inputBuffer;
        int m_bufferStart;
        int m_bufferLength;
        int m_inputStreamLength;
        int m_currentPositionInBuffer;
        Stream m_inputStream;
        bool m_isUserStream;

        public CssBuffer(Stream s, bool isUserStream)
        {
            m_inputStream = s; this.m_isUserStream = isUserStream;

            if (m_inputStream.CanSeek)
            {
                m_inputStreamLength = (int)m_inputStream.Length;
                m_bufferLength = Math.Min(m_inputStreamLength, MAX_BUFFER_LENGTH);
                m_bufferStart = Int32.MaxValue;
            }
            else
            {
                m_inputStreamLength = m_bufferLength = m_bufferStart = 0;
            }

            m_inputBuffer = new byte[(m_bufferLength > 0) ? m_bufferLength : MIN_BUFFER_LENGTH];
            if (m_inputStreamLength > 0)
                Pos = 0;
            else
                m_currentPositionInBuffer = 0;
            if (m_bufferLength == m_inputStreamLength && m_inputStream.CanSeek)
                Close();
        }

        protected CssBuffer(CssBuffer b)
        {
            m_inputBuffer = b.m_inputBuffer;
            m_bufferStart = b.m_bufferStart;
            m_bufferLength = b.m_bufferLength;
            m_inputStreamLength = b.m_inputStreamLength;
            m_currentPositionInBuffer = b.m_currentPositionInBuffer;
            m_inputStream = b.m_inputStream;
            b.m_inputStream = null;
            m_isUserStream = b.m_isUserStream;
        }

        ~CssBuffer() { Close(); }

        protected void Close()
        {
            if (!m_isUserStream && m_inputStream != null)
            {
                m_inputStream.Close();
                m_inputStream = null;
            }
        }

        public virtual int Read()
        {
            if (m_currentPositionInBuffer < m_bufferLength)
            {
                return m_inputBuffer[m_currentPositionInBuffer++];
            }
            else if (Pos < m_inputStreamLength)
            {
                Pos = Pos;
                return m_inputBuffer[m_currentPositionInBuffer++];
            }
            else if (m_inputStream != null && !m_inputStream.CanSeek && ReadNextStreamChunk() > 0)
            {
                return m_inputBuffer[m_currentPositionInBuffer++];
            }
            else
            {
                return EOF;
            }
        }

        public int Peek()
        {
            int curPos = Pos;
            int ch = Read();
            Pos = curPos;
            return ch;
        }

        public string GetString(int beg, int end)
        {
            int len = 0;
            char[] buf = new char[end - beg];
            int oldPos = Pos;
            Pos = beg;
            while (Pos < end)
                buf[len++] = (char)Read();
            Pos = oldPos;
            return new String(buf, 0, len);
        }

        public int Pos
        {
            get { return m_currentPositionInBuffer + m_bufferStart; }
            set
            {
                if (value >= m_inputStreamLength && m_inputStream != null && !m_inputStream.CanSeek)
                {
                    while (value >= m_inputStreamLength && ReadNextStreamChunk() > 0) ;
                }

                if (value < 0 || value > m_inputStreamLength)
                {
                    throw new FatalError("buffer out of bounds access, position: " + value);
                }

                if (value >= m_bufferStart && value < m_bufferStart + m_bufferLength)
                {
                    m_currentPositionInBuffer = value - m_bufferStart;
                }
                else if (m_inputStream != null)
                {
                    m_inputStream.Seek(value, SeekOrigin.Begin);
                    m_bufferLength = m_inputStream.Read(m_inputBuffer, 0, m_inputBuffer.Length);
                    m_bufferStart = value; m_currentPositionInBuffer = 0;
                }
                else
                {
                    m_currentPositionInBuffer = m_inputStreamLength - m_bufferStart;
                }
            }
        }

        private int ReadNextStreamChunk()
        {
            int free = m_inputBuffer.Length - m_bufferLength;
            if (free == 0)
            {
                byte[] newBuf = new byte[m_bufferLength * 2];
                Array.Copy(m_inputBuffer, newBuf, m_bufferLength);
                m_inputBuffer = newBuf;
                free = m_bufferLength;
            }
            int read = m_inputStream.Read(m_inputBuffer, m_bufferLength, free);
            if (read > 0)
            {
                m_inputStreamLength = m_bufferLength = (m_bufferLength + read);
                return read;
            }
            return 0;
        }
    }

    public class UTF8Buffer : CssBuffer
    {
        public UTF8Buffer(CssBuffer b) : base(b) { }

        public override int Read()
        {
            int ch;
            do
            {
                ch = base.Read();
            } while ((ch >= 128) && ((ch & 0xC0) != 0xC0) && (ch != EOF));
            if (ch < 128 || ch == EOF)
            {
                // nothing to do
            }
            else if ((ch & 0xF0) == 0xF0)
            {
                int c1 = ch & 0x07;
                ch = base.Read();
                int c2 = ch & 0x3F;
                ch = base.Read();
                int c3 = ch & 0x3F;
                ch = base.Read();
                int c4 = ch & 0x3F;
                ch = (((((c1 << 6) | c2) << 6) | c3) << 6) | c4;
            }
            else if ((ch & 0xE0) == 0xE0)
            {
                int c1 = ch & 0x0F;
                ch = base.Read();
                int c2 = ch & 0x3F;
                ch = base.Read();
                int c3 = ch & 0x3F;
                ch = (((c1 << 6) | c2) << 6) | c3;
            }
            else if ((ch & 0xC0) == 0xC0)
            {
                int c1 = ch & 0x1F;
                ch = base.Read();
                int c2 = ch & 0x3F;
                ch = (c1 << 6) | c2;
            }
            return ch;
        }
    }

    public class Scanner
    {
        const char END_OF_LINE = '\n';
        const int c_eof = 0;
        const int c_maxT = 49;
        const int c_noSym = 49;
        const int c_maxTokenLength = 128;

        public CssBuffer m_scannerBuffer;

        CssToken m_currentToken;
        int m_currentInputCharacter;
        int m_currentCharacterBytePosition;
        int m_unicodeCharacterPosition;
        int m_columnNumberOfCurrentCharacter;
        int m_lineNumberOfCurrentCharacter;
        int m_eolInComment;
        static readonly Hashtable s_start;

        CssToken m_tokensAlreadyPeeked;
        CssToken m_currentPeekToken;

        char[] m_textOfCurrentToken = new char[c_maxTokenLength];
        int m_lengthOfCurrentToken;

        static Scanner()
        {
            s_start = new Hashtable(128);
            for (int i = 65; i <= 84; ++i)
                s_start[i] = 1;
            for (int i = 86; i <= 90; ++i)
                s_start[i] = 1;
            for (int i = 95; i <= 95; ++i)
                s_start[i] = 1;
            for (int i = 97; i <= 122; ++i)
                s_start[i] = 1;
            for (int i = 10; i <= 10; ++i)
                s_start[i] = 2;
            for (int i = 13; i <= 13; ++i)
                s_start[i] = 2;
            for (int i = 48; i <= 57; ++i)
                s_start[i] = 3;
            for (int i = 9; i <= 9; ++i)
                s_start[i] = 4;
            for (int i = 11; i <= 12; ++i)
                s_start[i] = 4;
            for (int i = 32; i <= 32; ++i)
                s_start[i] = 4;
            s_start[60] = 5;
            s_start[45] = 40;
            s_start[39] = 11;
            s_start[34] = 12;
            s_start[40] = 13;
            s_start[41] = 14;
            s_start[64] = 15;
            s_start[44] = 16;
            s_start[123] = 17;
            s_start[59] = 18;
            s_start[125] = 19;
            s_start[43] = 20;
            s_start[62] = 21;
            s_start[126] = 41;
            s_start[42] = 42;
            s_start[35] = 22;
            s_start[46] = 23;
            s_start[91] = 24;
            s_start[61] = 25;
            s_start[124] = 27;
            s_start[36] = 29;
            s_start[94] = 31;
            s_start[93] = 34;
            s_start[58] = 35;
            s_start[33] = 36;
            s_start[47] = 37;
            s_start[85] = 43;
            s_start[37] = 39;
            s_start[CssBuffer.EOF] = -1;

        }

        public Scanner(string fileName)
        {
            try
            {
                Stream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                m_scannerBuffer = new CssBuffer(stream, false);
                Init();
            }
            catch (IOException)
            {
                throw new FatalError("Cannot open file " + fileName);
            }
        }

        public Scanner(Stream s)
        {
            m_scannerBuffer = new CssBuffer(s, true);
            Init();
        }

        void Init()
        {
            m_currentCharacterBytePosition = -1;
            m_lineNumberOfCurrentCharacter = 1;
            m_columnNumberOfCurrentCharacter = 0;
            m_unicodeCharacterPosition = -1;
            m_eolInComment = 0;
            NextCh();
            if (m_currentInputCharacter == 0xEF)
            {
                NextCh();
                int ch1 = m_currentInputCharacter;
                NextCh();
                int ch2 = m_currentInputCharacter;
                if (ch1 != 0xBB || ch2 != 0xBF)
                {
                    throw new FatalError(String.Format("illegal byte order mark: EF {0,2:X} {1,2:X}", ch1, ch2));
                }
                m_scannerBuffer = new UTF8Buffer(m_scannerBuffer);
                m_columnNumberOfCurrentCharacter = 0;
                m_unicodeCharacterPosition = -1;
                NextCh();
            }
            m_currentPeekToken = m_tokensAlreadyPeeked = new CssToken();
        }

        void NextCh()
        {
            if (m_eolInComment > 0)
            {
                m_currentInputCharacter = END_OF_LINE;
                m_eolInComment--;
            }
            else
            {
                m_currentCharacterBytePosition = m_scannerBuffer.Pos;
                m_currentInputCharacter = m_scannerBuffer.Read();
                m_columnNumberOfCurrentCharacter++;
                m_unicodeCharacterPosition++;
                if (m_currentInputCharacter == '\r' && m_scannerBuffer.Peek() != '\n')
                    m_currentInputCharacter = END_OF_LINE;
                if (m_currentInputCharacter == END_OF_LINE)
                {
                    m_lineNumberOfCurrentCharacter++; m_columnNumberOfCurrentCharacter = 0;
                }
            }

        }

        void AddCh()
        {
            if (m_lengthOfCurrentToken >= m_textOfCurrentToken.Length)
            {
                char[] newBuf = new char[2 * m_textOfCurrentToken.Length];
                Array.Copy(m_textOfCurrentToken, 0, newBuf, 0, m_textOfCurrentToken.Length);
                m_textOfCurrentToken = newBuf;
            }
            if (m_currentInputCharacter != CssBuffer.EOF)
            {
                m_textOfCurrentToken[m_lengthOfCurrentToken++] = (char)m_currentInputCharacter;
                NextCh();
            }
        }

        bool Comment0()
        {
            int level = 1, pos0 = m_currentCharacterBytePosition, line0 = m_lineNumberOfCurrentCharacter, col0 = m_columnNumberOfCurrentCharacter, charPos0 = m_unicodeCharacterPosition;
            NextCh();
            if (m_currentInputCharacter == '*')
            {
                NextCh();
                for (;;)
                {
                    if (m_currentInputCharacter == '*')
                    {
                        NextCh();
                        if (m_currentInputCharacter == '/')
                        {
                            level--;
                            if (level == 0)
                            {
                                m_eolInComment = m_lineNumberOfCurrentCharacter - line0;
                                NextCh();
                                return true;
                            }
                            NextCh();
                        }
                    }
                    else if (m_currentInputCharacter == CssBuffer.EOF)
                        return false;
                    else
                        NextCh();
                }
            }
            else
            {
                m_scannerBuffer.Pos = pos0;
                NextCh();
                m_lineNumberOfCurrentCharacter = line0;
                m_columnNumberOfCurrentCharacter = col0;
                m_unicodeCharacterPosition = charPos0;
            }
            return false;
        }


        void CheckLiteral()
        {
            switch (m_currentToken.m_tokenValue)
            {
                case "url":
                    m_currentToken.m_tokenKind = 9;
                    break;
                case "all":
                    m_currentToken.m_tokenKind = 12;
                    break;
                case "aural":
                    m_currentToken.m_tokenKind = 13;
                    break;
                case "braille":
                    m_currentToken.m_tokenKind = 14;
                    break;
                case "embossed":
                    m_currentToken.m_tokenKind = 15;
                    break;
                case "handheld":
                    m_currentToken.m_tokenKind = 16;
                    break;
                case "print":
                    m_currentToken.m_tokenKind = 17;
                    break;
                case "projection":
                    m_currentToken.m_tokenKind = 18;
                    break;
                case "screen":
                    m_currentToken.m_tokenKind = 19;
                    break;
                case "tty":
                    m_currentToken.m_tokenKind = 20;
                    break;
                case "tv":
                    m_currentToken.m_tokenKind = 21;
                    break;
                case "n":
                    m_currentToken.m_tokenKind = 22;
                    break;
                case "important":
                    m_currentToken.m_tokenKind = 45;
                    break;
                default:
                    break;
            }
        }

        CssToken NextToken()
        {
            while (m_currentInputCharacter == ' ' || m_currentInputCharacter == 10 || m_currentInputCharacter == 13)
                NextCh();
            if (m_currentInputCharacter == '/' && Comment0())
                return NextToken();
            int recKind = c_noSym;
            int recEnd = m_currentCharacterBytePosition;
            m_currentToken = new CssToken();
            m_currentToken.m_tokenPositionInBytes = m_currentCharacterBytePosition;
            m_currentToken.m_tokenColumn = m_columnNumberOfCurrentCharacter;
            m_currentToken.m_tokenLine = m_lineNumberOfCurrentCharacter;
            m_currentToken.m_tokenPositionInCharacters = m_unicodeCharacterPosition;
            int state;
            if (s_start.ContainsKey(m_currentInputCharacter))
            {
                state = (int)s_start[m_currentInputCharacter];
            }
            else {
                state = 0;
            }
            m_lengthOfCurrentToken = 0;
            AddCh();

            switch (state)
            {
                case -1: {
                    m_currentToken.m_tokenKind = c_eof;
                    break;
                }
                case 0:
                    {
                        if (recKind != c_noSym)
                        {
                            m_lengthOfCurrentToken = recEnd - m_currentToken.m_tokenPositionInBytes;
                            SetScannerBehindT();
                        }
                        m_currentToken.m_tokenKind = recKind;
                        break;
                    }
                case 1:
                    recEnd = m_currentCharacterBytePosition; recKind = 1;
                    if (m_currentInputCharacter == '-' || m_currentInputCharacter >= '0' && m_currentInputCharacter <= '9' || m_currentInputCharacter >= 'A' && m_currentInputCharacter <= 'Z' || m_currentInputCharacter == '_' || m_currentInputCharacter >= 'a' && m_currentInputCharacter <= 'z')
                    {
                        AddCh();
                        goto case 1;
                    }
                    else
                    {
                        m_currentToken.m_tokenKind = 1; m_currentToken.m_tokenValue = new String(m_textOfCurrentToken, 0, m_lengthOfCurrentToken);
                        CheckLiteral();
                        return m_currentToken;
                    }
                case 2:
                    {
                        m_currentToken.m_tokenKind = 2;
                        break;
                    }
                case 3:
                    {
                        m_currentToken.m_tokenKind = 3;
                        break;
                    }
                case 4:
                    {
                        m_currentToken.m_tokenKind = 4;
                        break;
                    }
                case 5:
                    if (m_currentInputCharacter == '!')
                    {
                        AddCh();
                        goto case 6;
                    }
                    else
                    {
                        goto case 0;
                    }
                case 6:
                    if (m_currentInputCharacter == '-')
                    {
                        AddCh();
                        goto case 7;
                    }
                    else {
                        goto case 0;
                    }
                case 7:
                    if (m_currentInputCharacter == '-')
                    {
                        AddCh();
                        goto case 8;
                    }
                    else
                    {
                        goto case 0;
                    }
                case 8:
                    {
                        m_currentToken.m_tokenKind = 5;
                        break;
                    }
                case 9:
                    if (m_currentInputCharacter == '>')
                    {
                        AddCh();
                        goto case 10;
                    }
                    else
                    {
                        goto case 0;
                    }
                case 10:
                    {
                        m_currentToken.m_tokenKind = 6;
                        break;
                    }
                case 11:
                    {
                        m_currentToken.m_tokenKind = 7;
                        break;
                    }
                case 12:
                    {
                        m_currentToken.m_tokenKind = 8;
                        break;
                    }
                case 13:
                    {
                        m_currentToken.m_tokenKind = 10;
                        break;
                    }
                case 14:
                    {
                        m_currentToken.m_tokenKind = 11;
                        break;
                    }
                case 15:
                    {
                        m_currentToken.m_tokenKind = 23;
                        break;
                    }
                case 16:
                    {
                        m_currentToken.m_tokenKind = 25;
                        break;
                    }
                case 17:
                    {
                        m_currentToken.m_tokenKind = 26;
                        break;
                    }
                case 18:
                    {
                        m_currentToken.m_tokenKind = 27;
                        break;
                    }
                case 19:
                    {
                        m_currentToken.m_tokenKind = 28;
                        break;
                    }
                case 20:
                    {
                        m_currentToken.m_tokenKind = 29;
                        break;
                    }
                case 21:
                    {
                        m_currentToken.m_tokenKind = 30;
                        break;
                    }
                case 22:
                    {
                        m_currentToken.m_tokenKind = 33;
                        break;
                    }
                case 23:
                    {
                        m_currentToken.m_tokenKind = 34;
                        break;
                    }
                case 24:
                    {
                        m_currentToken.m_tokenKind = 35;
                        break;
                    }
                case 25:
                    {
                        m_currentToken.m_tokenKind = 36;
                        break;
                    }
                case 26:
                    {
                        m_currentToken.m_tokenKind = 37;
                        break;
                    }
                case 27:
                    if (m_currentInputCharacter == '=')
                    {
                        AddCh();
                        goto case 28;
                    }
                    else {
                        goto case 0;
                    }
                case 28:
                    {
                        m_currentToken.m_tokenKind = 38;
                        break;
                    }
                case 29:
                    if (m_currentInputCharacter == '=')
                    {
                        AddCh();
                        goto case 30;
                    }
                    else {
                        goto case 0;
                    }
                case 30:
                    {
                        m_currentToken.m_tokenKind = 39;
                        break;
                    }
                case 31:
                    if (m_currentInputCharacter == '=')
                    {
                        AddCh();
                        goto case 32;
                    }
                    else
                    {
                        goto case 0;
                    }
                case 32:
                    {
                        m_currentToken.m_tokenKind = 40;
                        break;
                    }
                case 33:
                    {
                        m_currentToken.m_tokenKind = 41;
                        break;
                    }
                case 34:
                    {
                        m_currentToken.m_tokenKind = 42;
                        break;
                    }
                case 35:
                    {
                        m_currentToken.m_tokenKind = 43;
                        break;
                    }
                case 36:
                    {
                        m_currentToken.m_tokenKind = 44;
                        break;
                    }
                case 37:
                    {
                        m_currentToken.m_tokenKind = 46;
                        break;
                    }
                case 38:
                    {
                        m_currentToken.m_tokenKind = 47;
                        break;
                    }
                case 39:
                    {
                        m_currentToken.m_tokenKind = 48;
                        break;
                    }
                case 40:
                    recEnd = m_currentCharacterBytePosition;
                    recKind = 24;
                    if (m_currentInputCharacter == '-')
                    {
                        AddCh();
                        goto case 9;
                    }
                    else
                    {
                        m_currentToken.m_tokenKind = 24;
                        break;
                    }
                case 41:
                    recEnd = m_currentCharacterBytePosition;
                    recKind = 31;
                    if (m_currentInputCharacter == '=')
                    {
                        AddCh();
                        goto case 26;
                    }
                    else
                    {
                        m_currentToken.m_tokenKind = 31;
                        break;
                    }
                case 42:
                    recEnd = m_currentCharacterBytePosition;
                    recKind = 32;
                    if (m_currentInputCharacter == '=')
                    {
                        AddCh();
                        goto case 33;
                    }
                    else
                    {
                        m_currentToken.m_tokenKind = 32;
                        break;
                    }
                case 43:
                    recEnd = m_currentCharacterBytePosition; recKind = 1;
                    if (m_currentInputCharacter == '-' || m_currentInputCharacter >= '0' && m_currentInputCharacter <= '9' || m_currentInputCharacter >= 'A' && m_currentInputCharacter <= 'Z' || m_currentInputCharacter == '_' || m_currentInputCharacter >= 'a' && m_currentInputCharacter <= 'z')
                    {
                        AddCh();
                        goto case 1;
                    }
                    else if (m_currentInputCharacter == 92)
                    {
                        AddCh();
                        goto case 38;
                    }
                    else
                    {
                        m_currentToken.m_tokenKind = 1;
                        m_currentToken.m_tokenValue = new String(m_textOfCurrentToken, 0, m_lengthOfCurrentToken);
                        CheckLiteral();
                        return m_currentToken;
                    }

            }
            m_currentToken.m_tokenValue = new String(m_textOfCurrentToken, 0, m_lengthOfCurrentToken);
            return m_currentToken;
        }

        private void SetScannerBehindT()
        {
            m_scannerBuffer.Pos = m_currentToken.m_tokenPositionInBytes;
            NextCh();
            m_lineNumberOfCurrentCharacter = m_currentToken.m_tokenLine; m_columnNumberOfCurrentCharacter = m_currentToken.m_tokenColumn; m_unicodeCharacterPosition = m_currentToken.m_tokenPositionInCharacters;
            for (int i = 0; i < m_lengthOfCurrentToken; i++) NextCh();
        }

        public CssToken Scan()
        {
            if (m_tokensAlreadyPeeked.m_nextToken == null)
            {
                return NextToken();
            }
            else
            {
                m_currentPeekToken = m_tokensAlreadyPeeked = m_tokensAlreadyPeeked.m_nextToken;
                return m_tokensAlreadyPeeked;
            }
        }

        public CssToken Peek()
        {
            do
            {
                if (m_currentPeekToken.m_nextToken == null)
                {
                    m_currentPeekToken.m_nextToken = NextToken();
                }
                m_currentPeekToken = m_currentPeekToken.m_nextToken;
            } while (m_currentPeekToken.m_tokenKind > c_maxT);

            return m_currentPeekToken;
        }

        public void ResetPeek()
        {
            m_currentPeekToken = m_tokensAlreadyPeeked;
        }
    }

}
