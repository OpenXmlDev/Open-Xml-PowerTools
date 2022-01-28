using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Codeuctivity
{
    internal class CssApplier
    {
        private static readonly List<PropertyInfo> PropertyInfoList = new List<PropertyInfo>()
        {
            // color
            // Value:          <color> | inherit
            // Initial:        depends on UA
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "color" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "black", Type = CssTermType.String } } },
                ComputedValue = (element, assignedValue, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = GetWmlColorFromExpression(assignedValue), Type = CssTermType.String } } },
            },

            // direction
            // Value:          ltr | rtl | inherit
            // Initial:        ltr
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "direction" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "ltr", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // line-height
            // Value:          normal | <number> | <length> | <percentage> | <inherit>
            // Initial:        normal
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    refer to the font size of the element itself
            // Computed value: for <length> and <percentage> the absolute value, otherwise as specified.
            new PropertyInfo
            {
                Names = new[] { "line-height" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element, "font-size", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // visibility
            // Value:          visible | hidden | collapse | inherit
            // Initial:        visible
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "visibility" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "visible", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // list-style-type
            // Value:          disc | circle | square | decimal | decimal-leading-zero |
            //                 lower-roman | upper-roman | lower-greek | lower-latin |
            //                 upper-latin | armenian | georgian | lower-alpha | upper-alpha |
            //                 none | inherit
            // Initial:        disc
            // Applies to:     elements with display: list-item
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "list-style-type" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "list-item") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "disc", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // list-style-image
            // Value:          <uri> | none | inherit
            // Initial:        none
            // Applies to:     elements with ’display: list-item’
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: absolute URI or ’none’
            new PropertyInfo
            {
                Names = new[] { "list-style-image" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "list-item") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // list-style-position
            // Value:          inside | outside | inherit
            // Initial:        outside
            // Applies to:     elements with ’display: list-item’
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "list-style-position" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "list-item") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // font-family
            // Value:          [[ <family-name> | <generic-family> ] [, <family-name>|
            //                 <generic-family>]* ] | inherit
            // Initial:        depends on user agent
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "font-family" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = settings.MinorLatinFont, Type = CssTermType.String } } },
                ComputedValue = (element, assignedValue, settings) => assignedValue,
            },

            // font-style
            // Value:          normal | italic | oblique | inherit
            // Initial:        normal
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "font-style" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // font-variant
            // Value:          normal | small-caps | inherit
            // Initial:        normal
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "font-variant" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // font-weight
            // Value:          normal | bold | bolder | lighter | 100 | 200 | 300 | 400 | 500 |
            //                 600 | 700 | 800 | 900 | inherit
            // Initial:        normal
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: see text
            new PropertyInfo
            {
                Names = new[] { "font-weight" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // font-size
            // Value:          <absolute-size> | <relative-size> | <length> | <percentage> |
            //                 inherit
            // Initial:        medium
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    refer to inherited font size
            // Computed value: absolute length
            new PropertyInfo
            {
                Names = new[] { "font-size" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = settings.DefaultFontSize.ToString(CultureInfo.InvariantCulture), Type = CssTermType.String, Unit = CssUnit.PT } } },
                ComputedValue = (element, assignedValue, settings) => ComputeAbsoluteFontSize(element, assignedValue, settings),
            },

            // text-indent
            // Value:          <length> | <percentage> | inherit
            // Initial:        0
            // Applies to:     block containers
            // Inherited:      yes
            // Percentages:    refer to width of containing block
            // Computed value: the percentage as specified or the absolute length
            new PropertyInfo
            {
                Names = new[] { "text-indent" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "block") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT, } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "width", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // text-align
            // Value:          left | right | center | justify | inherit
            // Initial:        a nameless value that acts as ’left’ if ’direction’ is ’ltr’, ’right’ if
            //                 ’direction’ is ’rtl’
            // Applies to:     block containers
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: the initial value or as spec
            new PropertyInfo
            {
                Names = new[] { "text-align" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "block") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "left", Type = CssTermType.String, } } },  // todo should be based on the direction property
                ComputedValue = null,
            },

            // text-decoration
            // Value:          none | [ underline || overline || line-through || blink ] | inherit
            // Initial:        none
            // Applies to:     all elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "text-decoration" },
                Inherits = true,   // todo need to read css 16.3.1 in full detail to understand how this is implemented.
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String, } } },
                ComputedValue = null,
            },

            // letter-spacing
            // Value:          normal | <length> | inherit
            // Initial:        normal
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: ’normal’ or absolute length

            // word-spacing
            // Value:          normal | <length> | inherit
            // Initial:        normal
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: for ’normal’ the value 0; otherwise the absolute length
            new PropertyInfo
            {
                Names = new[] { "letter-spacing", "word-spacing" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String, } } },
                ComputedValue = (element, assignedValue, settings) => ComputeAbsoluteLength(element, assignedValue, settings, null),
            },

            // white-space
            // Value:          normal | pre | nowrap | pre-wrap | pre-line | inherit
            // Initial:        normal
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "white-space" },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String, } } },
                ComputedValue = null,
            },

            // caption-side
            // Value:          top | bottom | inherit
            // Initial:        top
            // Applies to:     'table-caption' elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "caption-side" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "table-caption") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "top", Type = CssTermType.String, } } },
                ComputedValue = null,
            },

            // border-collapse
            // Value:          collapse | separate | inherit
            // Initial:        separate
            // Applies to:     ’table’ and ’inline-table’ elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "border-collapse" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "table" || display.ToString() == "inline-table") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "separate", Type = CssTermType.String, } } },
                ComputedValue = null,
            },

            // border-spacing
            // Value:          <length> <length>? | inherit
            // Initial:        0
            // Applies to:     ’table’ and ’inline-table’ elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: two absolute lengths
            new PropertyInfo
            {
                Names = new[] { "border-spacing" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "table" || display.ToString() == "inline-table") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT, } } },
                ComputedValue = (element, assignedValue, settings) => ComputeAbsoluteLength(element, assignedValue, settings, null),  // todo need to handle two lengths here
            },

            // empty-cells
            // Value:          show | hide | inherit
            // Initial:        show
            // Applies to:     'table-cell' elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "empty-cells" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "table" || display.ToString() == "table-cell") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "show", } } },
                ComputedValue = null,
            },

            // margin-top, margin-bottom
            // Value:          <margin-width> | inherit
            // Initial:        0
            // Applies to:     all elements except elements with table display types other than table-caption, table, and inline-table
            //                 all elements except th, td, tr
            // Inherited:      no
            // Percentages:    refer to width of containing block
            // Computed value: the percentage as specified or the absolute length
            new PropertyInfo
            {
                Names = new[] { "margin-top", "margin-bottom", },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "table-caption" || display.ToString() == "table" || display.ToString() == "inline-table") { return false; } return true;
                },
                InitialValue = (element, settings) =>
                    {
                        if (settings.DefaultBlockContentMargin != null)
                        {
                            if (settings.DefaultBlockContentMargin == "auto") { return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "auto", Type = CssTermType.String } } }; } else if (settings.DefaultBlockContentMargin.ToLower().EndsWith("pt"))
                            {
                                var s1 = settings.DefaultBlockContentMargin.Substring(0, settings.DefaultBlockContentMargin.Length - 2);
                                if (double.TryParse(s1, NumberStyles.Float, CultureInfo.InvariantCulture, out var d1))
                                {
                                    return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = d1.ToString(CultureInfo.InvariantCulture), Type = CssTermType.Number, Unit = CssUnit.PT } } };
                                }
                            }
                            throw new OpenXmlPowerToolsException("invalid setting");
                        }
                        return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT } } };
                    },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "width", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // margin-right, margin-left
            // Value:          <margin-width> | inherit
            // Initial:        0
            // Applies to:     all elements except elements with table display types other than table-caption, table, and inline-table
            //                 all elements except th, td, tr
            // Inherited:      no
            // Percentages:    refer to width of containing block
            // Computed value: the percentage as specified or the absolute length
            new PropertyInfo
            {
                Names = new[] { "margin-right", "margin-left", },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "table-caption" || display.ToString() == "table" || display.ToString() == "inline-table") { return false; } return true;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT, } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "width", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // padding-top, padding-right, padding-bottom, padding-left
            // Value:          <padding-width> | inherit
            // Initial:        0
            // Applies to:     all elements except table-row-group, table-header-group,
            //                 table-footer-group, table-row, table-column-group and table-column
            //                 all elements except tr
            // Inherited:      no
            // Percentages:    refer to width of containing block
            // Computed value: the percentage as specified or the absolute length
            new PropertyInfo
            {
                Names = new[] { "padding-top", "padding-right", "padding-bottom", "padding-left" },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    var dv = display.ToString();
                    if (dv == "table-row-group" || dv == "table-header-group" || dv == "table-footer-group" || dv == "table-row" ||
                        dv == "table-column-group" || dv == "table-column") { return false; } return true;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT, } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "width", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // border-top-width, border-right-width, border-bottom-width, border-left-width
            // Value:          <border-width> | inherit
            // Initial:        medium
            // Applies to:     all elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: absolute length; '0' if the border style is 'none' or 'hidden'
            new PropertyInfo
            {
                Names = new[] { "border-top-width", "border-right-width", "border-bottom-width", "border-left-width", },
                Inherits = false,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT, } } },
                ComputedValue = (element, assignedValue, settings) =>
                    {
                        var assignedValueStr = assignedValue.ToString();
                        if (assignedValueStr == "thin") { return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0.75", Type = CssTermType.Number, Unit = CssUnit.PT, } } }; } if (assignedValueStr == "medium") { return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "3.0", Type = CssTermType.Number, Unit = CssUnit.PT, } } }; } if (assignedValueStr == "thick") { return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "4.5", Type = CssTermType.Number, Unit = CssUnit.PT, } } }; } return ComputeAbsoluteLength(element, assignedValue, settings, null);
                    },
            },

            // border-top-style, border-right-style, border-bottom-style, border-left-style
            // Value:          <border-style> | inherit
            // Initial:        none
            // Applies to:     all elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as specified
            new PropertyInfo
            {
                Names = new[] { "border-top-style", "border-right-style", "border-bottom-style", "border-left-style", },
                Inherits = false,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // display
            // Value:          inline | block | list-item | inline-block | table | inline-table |
            //                 table-row-group | table-header-group | table-footer-group |
            //                 table-row | table-column-group | table-column | table-cell |
            //                 table-caption | none | inherit
            // Initial:        inline
            // Applies to:     all elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: see text
            new PropertyInfo
            {
                Names = new[] { "display", },
                Inherits = false,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inline", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // position
            // Value:          static | relative | absolute | fixed | inherit
            // Initial:        static
            // Applies to:     all elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as specified
            new PropertyInfo
            {
                Names = new[] { "position", },
                Inherits = false,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "static", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // float
            // Value:          left | right | none | inherit
            // Initial:        none
            // Applies to:     all, but see 9.7 p. 153
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as specified
            new PropertyInfo
            {
                Names = new[] { "float", },
                Inherits = false,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // unicode-bidi
            // Value:          normal | embed | bidi-override | inherit
            // Initial:        normal
            // Applies to:     all elements, but see prose
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "unicode-bidi", },
                Inherits = false,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // background-color
            // Value:          <color> | transparent | inherit
            // Initial:        transparent
            // Applies to:     all elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "background-color", },
                Inherits = false,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "transparent", Type = CssTermType.String } } },
                ComputedValue = (element, assignedValue, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = GetWmlColorFromExpression(assignedValue), Type = CssTermType.String } } },
            },

            // text-transform
            // Value:          capitalize | uppercase | lowercase | none | inherit
            // Initial:        none
            // Applies to:     all elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "text-transform", },
                Inherits = true,
                Includes = (e, settings) => true,
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String } } },
                ComputedValue = null,
            },

            // table-layout
            // Value:          auto | fixed | inherit
            // Initial:        auto
            // Applies to:     ’table’ and ’inline-table’ elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "table-layout" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "table" || display.ToString() == "inline-table") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "auto", Type = CssTermType.String, } } },
                ComputedValue = null,
            },

            // empty-cells
            // Value:          show | hide | inherit
            // Initial:        show
            // Applies to:     'table-cell' elements
            // Inherited:      yes
            // Percentages:    N/A
            // Computed value: as spec
            new PropertyInfo
            {
                Names = new[] { "border-spacing" },
                Inherits = true,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    if (display.ToString() == "table-cell") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "show", Type = CssTermType.String, } } },
                ComputedValue = null,
            },

            // border-top-color, border-right-color, border-bottom-color, border-left-color
            // Value:          <color> | transparent | inherit
            // Initial:        the value of the color property
            // Applies to:     all elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: when taken from the ’color’ property, the computed value of
            //                 ’color’; otherwise, as specified
            new PropertyInfo
            {
                Names = new[] { "border-top-color", "border-right-color", "border-bottom-color", "border-left-color", },
                Inherits = false,
                Includes = (e, settings) => true,
                InitialValue = (e, settings) => {
                    var display = GetComputedPropertyValue(null, e, "color", settings);
                    return display;
                },
                ComputedValue = (element, assignedValue, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = GetWmlColorFromExpression(assignedValue), Type = CssTermType.String } } },
            },

            // width
            // Value:          <length> | <percentage> | auto | inherit
            // Initial:        auto
            // Applies to:     all elements but non-replaced in-line elements, table rows, and row groups
            // Inherited:      no
            // Percentages:    refer to width of containing block
            // Computed value: the percentage or 'auto' as specified or the absolute length
            new PropertyInfo
            {
                Names = new[] { "width" },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    if (e.Name == XhtmlNoNamespace.img) { return true; } var display = GetComputedPropertyValue(null, e, "display", settings);
                    var dv = display.ToString();
                    if (dv == "inline" || dv == "table-row" || dv == "table-row-group") { return false; } return true;
                },
                InitialValue = (element, settings) =>
                {
                    if (element.Parent == null)
                    {
                        var pageWidth = (double?)settings.SectPr.Elements(W.pgSz).Attributes(W._w).FirstOrDefault();
                        if (pageWidth == null) { pageWidth = 12240; } var leftMargin = (double?)settings.SectPr.Elements(W.pgMar).Attributes(W.left).FirstOrDefault();
                        if (leftMargin == null) { leftMargin = 1440; } var rightMargin = (double?)settings.SectPr.Elements(W.pgMar).Attributes(W.left).FirstOrDefault();
                        if (rightMargin == null) { rightMargin = 1440; } var width = (double)(pageWidth - leftMargin - rightMargin) / 20;
                        return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = width.ToString(CultureInfo.InvariantCulture), Type = CssTermType.String, Unit = CssUnit.PT, } } };
                    }
                    return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "auto", Type = CssTermType.String, } } };
                },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    if (element.Name != XhtmlNoNamespace.caption &&
                        element.Name != XhtmlNoNamespace.td &&
                        element.Name != XhtmlNoNamespace.th &&
                        element.Name != XhtmlNoNamespace.tr &&
                        element.Name != XhtmlNoNamespace.table &&
                        assignedValue.IsAuto)
                    {
                        var pi = PropertyInfoList.FirstOrDefault(p => p.Names.Contains("width"));
                        var display = GetComputedPropertyValue(pi, element, "display", settings).ToString();
                        if (display != "inline")
                        {
                            var parentPropertyValue = GetComputedPropertyValue(pi, element.Parent, "width", settings);
                            return parentPropertyValue;
                        }
                    }
                    CssExpression valueForPercentage = null;
                    var elementToQuery = element.Parent;
                    while (elementToQuery != null)
                    {
                        valueForPercentage = GetComputedPropertyValue(null, elementToQuery, "width", settings);
                        if (valueForPercentage.IsAuto)
                        {
                            elementToQuery = elementToQuery.Parent;
                            continue;
                        }
                        break;
                    }

                    return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // min-width
            // Value:          <length> | <percentage> | inherit
            // Initial:        0
            // Applies to:     all elements but non-replaced in-line elements, table rows, and row groups
            // Inherited:      no
            // Percentages:    refer to width of containing block
            // Computed value: the percentage as spec or the absolute length
            new PropertyInfo
            {
                Names = new[] { "min-width" },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    var dv = display.ToString();
                    if (dv == "inline" || dv == "table-row" || dv == "table-row-group") { return false; } return true;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT, } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "width", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // max-width
            // Value:          <length> | <percentage> | none | inherit
            // Initial:        none
            // Applies to:     all elements but non-replaced in-line elements, table rows, and row groups
            // Inherited:      no
            // Percentages:    refer to width of containing block
            // Computed value: the percentage as spec or the absolute length
            new PropertyInfo
            {
                Names = new[] { "max-width" },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    var dv = display.ToString();
                    if (dv == "inline" || dv == "table-row" || dv == "table-row-group") { return false; } return true;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String, } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "width", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // height
            // Value:          <length> | <percentage> | auto | inherit
            // Initial:        auto
            // Applies to:     all elements but non-replaced in-line elements, table columns, and column groups
            // Inherited:      no
            // Percentages:    see prose
            // Computed value: the percentage as spec or the absolute length
            new PropertyInfo
            {
                Names = new[] { "height" },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    if (e.Name == XhtmlNoNamespace.img) { return true; } var display = GetComputedPropertyValue(null, e, "display", settings);
                    var dv = display.ToString();
                    if (dv == "inline" || dv == "table-row" || dv == "table-row-group") { return false; } return true;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "auto", Type = CssTermType.String, } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "height", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // min-height
            // Value:          <length> | <percentage> | inherit
            // Initial:        0
            // Applies to:     all elements but non-replaced in-line elements, table columns, and column groups
            // Inherited:      no
            // Percentages:    see prose
            // Computed value: the percentage as spec or the absolute length
            new PropertyInfo
            {
                Names = new[] { "min-height" },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    var dv = display.ToString();
                    if (dv == "inline" || dv == "table-column" || dv == "table-column-group") { return false; } return true;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "0", Type = CssTermType.Number, Unit = CssUnit.PT, } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "height", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // max-height
            // Value:          <length> | <percentage> | none | inherit
            // Initial:        none
            // Applies to:     all elements but non-replaced in-line elements, table columns, and column groups
            // Inherited:      no
            // Percentages:    refer to height of containing block
            // Computed value: the percentage as spec or the absolute length
            new PropertyInfo
            {
                Names = new[] { "max-height" },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    var dv = display.ToString();
                    if (dv == "inline" || dv == "table-column" || dv == "table-column-group") { return false; } return true;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String, } } },
                ComputedValue = (element, assignedValue, settings) =>
                {
                    CssExpression valueForPercentage = null;
                    if (element.Parent != null) { valueForPercentage = GetComputedPropertyValue(null, element.Parent, "height", settings); } return ComputeAbsoluteLength(element, assignedValue, settings, valueForPercentage);
                },
            },

            // vertical-align
            // Value:          baseline | sub | super | top | text-top | middle | bottom | text-bottom |
            //                 <percentage> | <length> | inherit
            // Initial:        baseline
            // Applies to:     inline-level and 'table-cell' elements
            // Inherited:      no
            // Percentages:    refer to the line height of the element itself
            // Computed value: for <length> and <percentage> the absolute length, otherwise as specified.
            new PropertyInfo
            {
                Names = new[] { "vertical-align" },
                Inherits = false,
                Includes = (e, settings) =>
                {
                    var display = GetComputedPropertyValue(null, e, "display", settings);
                    var dv = display.ToString();
                    if (dv == "inline" || dv == "table-cell") { return true; } return false;
                },
                InitialValue = (element, settings) => new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "baseline", Type = CssTermType.String, } } },
                ComputedValue = (element, assignedValue, settings) => assignedValue,  // todo fix
            },

            // positioned elements are not supported
            //
            // top
            // Value:          <length> | <percentage> | auto | inherit
            // Initial:        auto
            // Applies to:     positioned elements
            // Inherited:      no
            // Percentages:    refer to height of containing block
            // Computed value: if specified as a length, the corresponding absolute length; if
            //                 specified as a percentage, the specified value; otherwise, ’auto’.
            //
            // right
            // Value:          <length> | <percentage> | auto | inherit
            // Initial:        auto
            // Applies to:     positioned elements
            // Inherited:      no
            // Percentages:    refer to width of containing block
            // Computed value: if specified as a length, the corresponding absolute length; if
            //                 specified as a percentage, the specified value; otherwise, ’auto’.
            //
            // bottom
            // Value:          <length> | <percentage> | auto | inherit
            // Initial:        auto
            // Applies to:     positioned elements
            // Inherited:      no
            // Percentages:    refer to height of containing block
            // Computed value: if specified as a length, the corresponding absolute length; if
            //                 specified as a percentage, the specified value; otherwise, ’auto’.
            //
            // left
            // Value:          <length> | <percentage> | auto | inherit
            // Initial:        auto
            // Applies to:     positioned elements
            // Inherited:      no
            // Percentages:    refer to width of containing block
            // Computed value: if specified as a length, the corresponding absolute length; if
            //                 specified as a percentage, the specified value; otherwise, ’auto’.

            // floated elements are not supported
            //
            // clear
            // Value:          none | left | right | both | inherit
            // Initial:        none
            // Applies to:     block-level elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as specified
            //
            // z-index
            // Value:          auto | integer | inherit
            // Initial:        auto
            // Applies to:     positioned elements
            // Inherited:      no
            // Percentages:    N/A
            // Computed value: as spec
        };

        /*
         * 1. Process user-agent default style sheet
         * 2. Process user-supplied style sheet
         * 3. process author-supplied style sheet
         * 4. process STYLE element
         * 5. process style attribute on all elements
         * 6. expand all shorthand properties - the new properties have the same sort key as shorthand prop
         * 7. Add initial values for all properties that don't have values
        */

        public static void ApplyAllCss(
            string defaultCss,
            string authorCss,
            string userCss,
            XElement newXHtml,
            HtmlToWmlConverterSettings settings,
            out CssDocument defaultCssDoc,
            out CssDocument authorCssDoc,
            out CssDocument userCssDoc,
            string annotatedHtmlDumpFileName)
        {
            var propertySequence = 1;

            var defaultCssParser = new CssParser();
            defaultCssDoc = defaultCssParser.ParseText(defaultCss);
            ApplyCssDocument(
                defaultCssDoc,
                newXHtml,
                Property.HighOrderPriority.UserAgentNormal,
                Property.HighOrderPriority.UserAgentHigh,
                ref propertySequence);

            var userCssParser = new CssParser();
            userCssDoc = userCssParser.ParseText(userCss);
            ApplyCssDocument(
                userCssDoc,
                newXHtml,
                Property.HighOrderPriority.UserHigh,
                Property.HighOrderPriority.UserNormal,
                ref propertySequence);

            var authorCssParser = new CssParser();
            authorCssDoc = authorCssParser.ParseText(authorCss);
            ApplyCssDocument(
                authorCssDoc,
                newXHtml,
                Property.HighOrderPriority.AuthorNormal,
                Property.HighOrderPriority.AuthorHigh,
                ref propertySequence);

            //string s = DumpCss(authorCssDoc);
            //File.WriteAllText("CssTreeDump.txt", s);
            //Environment.Exit(0);

            // If processing style element, do it here.

            ApplyStyleAttributes(newXHtml, ref propertySequence);

            ExpandShorthandProperties(newXHtml, settings);

            SetAllValues(newXHtml, settings);

            if (annotatedHtmlDumpFileName != null)
            {
                var sb = new StringBuilder();
                WriteXHtmlWithAnnotations(newXHtml, sb);
                File.WriteAllText(annotatedHtmlDumpFileName, sb.ToString());
            }
        }

        private static void SetAllValues(XElement xHtml, HtmlToWmlConverterSettings settings)
        {
            foreach (var element in xHtml.DescendantsAndSelf())
            {
                foreach (var propertyInfo in PropertyInfoList)
                {
                    if (propertyInfo.Includes(element, settings))
                    {
                        foreach (var name in propertyInfo.Names)
                        {
                            GetComputedPropertyValue(propertyInfo, element, name, settings);
                        }
                    }
                }
            }
        }

        public static CssExpression GetComputedPropertyValue(PropertyInfo propertyInfo, XElement element, string propertyName,
            HtmlToWmlConverterSettings settings)
        {
            // if (property is already computed)
            //   return the computed value
            var computedValues = element.Annotation<Dictionary<string, CssExpression>>();
            if (computedValues == null)
            {
                computedValues = new Dictionary<string, CssExpression>();
                element.AddAnnotation(computedValues);
            }
            if (computedValues.ContainsKey(propertyName))
            {
                var r = computedValues[propertyName];
                return r;
            }

            // if property is not set or property is set to inherited value, then get inherited or initialized value.
            var pName = propertyName.ToLower();
            if (propertyInfo == null)
            {
                propertyInfo = PropertyInfoList.FirstOrDefault(pi => pi.Names.Contains(pName));
                if (propertyInfo == null)
                {
                    throw new OpenXmlPowerToolsException("all possible properties should be in the list");
                }
            }
            var propList = element.Annotation<Dictionary<string, Property>>();
            if (propList == null)
            {
                var computedValue = GetInheritedOrInitializedValue(computedValues, propertyInfo, element, propertyName, false, settings);
                return computedValue;
            }
            if (!propList.ContainsKey(pName))
            {
                var computedValue = GetInheritedOrInitializedValue(computedValues, propertyInfo, element, propertyName, false, settings);
                return computedValue;
            }
            var prop = propList[pName];
            var propStr = prop.Expression.ToString();
            if (propStr == "inherited" || propStr == "auto")
            {
                var computedValue = GetInheritedOrInitializedValue(computedValues, propertyInfo, element, propertyName, true, settings);
                return computedValue;
            }
            // if property is set, then compute the value, return the computed value
            CssExpression computedValue2;
            if (propertyInfo.ComputedValue == null)
            {
                computedValue2 = prop.Expression;
            }
            else
            {
                computedValue2 = propertyInfo.ComputedValue(element, prop.Expression, settings);
            }
            computedValues.Add(propertyName, computedValue2);
            return computedValue2;
        }

        //if (property is inherited (either because it was set to inherit, or because it is an inherited property))
        //  if (parent exists)
        //    call GetComputedValue on parent
        //    return the computed value
        //else
        //  call GetInitialValue for property
        //  compute value
        //  set the computed value
        //  return the computed value
        public static CssExpression GetInheritedOrInitializedValue(Dictionary<string, CssExpression> computedValues, PropertyInfo propertyInfo, XElement element, string propertyName, bool valueIsInherit, HtmlToWmlConverterSettings settings)
        {
            if ((propertyInfo.Inherits || valueIsInherit) && element.Parent != null && propertyInfo.Includes(element.Parent, settings))
            {
                var parentPropertyValue = GetComputedPropertyValue(propertyInfo, element.Parent, propertyName, settings);
                computedValues.Add(propertyName, parentPropertyValue);
                return parentPropertyValue;
            }
            var initialPropertyValue = propertyInfo.InitialValue(element, settings);
            CssExpression computedValue;
            if (propertyInfo.ComputedValue == null)
            {
                computedValue = initialPropertyValue;
            }
            else
            {
                computedValue = propertyInfo.ComputedValue(element, initialPropertyValue, settings);
            }

            computedValues.Add(propertyName, computedValue);
            return computedValue;
        }

        private static void ApplyCssDocument(
            CssDocument cssDoc,
            XElement xHtml,
            Property.HighOrderPriority notImportantHighOrderSort,
            Property.HighOrderPriority importantHighOrderSort,
            ref int propertySequence)
        {
            foreach (var ruleSet in cssDoc.RuleSets)
            {
                foreach (var selector in ruleSet.Selectors)
                {
                    ApplySelector(selector, ruleSet, xHtml, notImportantHighOrderSort,
                        importantHighOrderSort, ref propertySequence);
                }
            }
        }

        private static CssExpression ComputeAbsoluteLength(XElement element, CssExpression assignedValue, HtmlToWmlConverterSettings settings,
            CssExpression lengthForPercentage)
        {
            if (assignedValue.Terms.Count != 1)
            {
                throw new OpenXmlPowerToolsException("Should not have a unit with more than one term");
            }

            var value = assignedValue.Terms.First().Value;
            var negative = assignedValue.Terms.First().Sign == '-';

            if (value == "thin")
            {
                var newExpr1 = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = ".3", Type = CssTermType.Number, Unit = CssUnit.PT, } } };
                return newExpr1;
            }
            if (value == "medium")
            {
                var newExpr2 = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "1.20", Type = CssTermType.Number, Unit = CssUnit.PT, } } };
                return newExpr2;
            }
            if (value == "thick")
            {
                var newExpr3 = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "1.80", Type = CssTermType.Number, Unit = CssUnit.PT, } } };
                return newExpr3;
            }
            if (value == "auto" || value == "normal" || value == "none")
            {
                return assignedValue;
            }

            var unit = assignedValue.Terms.First().Unit;
            if (unit == CssUnit.PT || unit == null)
            {
                return assignedValue;
            }

            if (unit == CssUnit.Percent && lengthForPercentage == null)
            {
                return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "auto", Type = CssTermType.String } } };
            }

            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var decValue))
            {
                throw new OpenXmlPowerToolsException("value did not parse");
            }

            if (negative)
            {
                decValue = -decValue;
            }

            double? newPtSize = null;
            if (unit == CssUnit.Percent)
            {
                if (!double.TryParse(lengthForPercentage.Terms.First().Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var ptSize))
                {
                    throw new OpenXmlPowerToolsException("did not return a double?");
                }

                newPtSize = ptSize * decValue / 100d;
            }
            else if (unit == CssUnit.EM || unit == CssUnit.EX)
            {
                var fontSize = GetComputedPropertyValue(null, element, "font-size", settings);
                if (!double.TryParse(fontSize.Terms.First().Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var decFontSize))
                {
                    throw new OpenXmlPowerToolsException("Internal error");
                }

                newPtSize = unit == CssUnit.EM ? decFontSize * decValue : decFontSize * decValue / 2;
            }
            else
            {
                if (unit == null && decValue == 0d)
                {
                    newPtSize = 0d;
                }

                if (unit == CssUnit.IN)
                {
                    newPtSize = decValue * 72.0d;
                }

                if (unit == CssUnit.CM)
                {
                    newPtSize = decValue / 2.54d * 72.0d;
                }

                if (unit == CssUnit.MM)
                {
                    newPtSize = decValue / 25.4d * 72.0d;
                }

                if (unit == CssUnit.PC)
                {
                    newPtSize = decValue * 12d;
                }

                if (unit == CssUnit.PX)
                {
                    newPtSize = decValue * 0.75d;
                }
            }
            if (!newPtSize.HasValue)
            {
                throw new OpenXmlPowerToolsException("Internal error: should not have reached this exception");
            }

            var newExpr = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = newPtSize.Value.ToString(CultureInfo.InvariantCulture), Type = CssTermType.Number, Unit = CssUnit.PT, } } };
            return newExpr;
        }

        private static CssExpression ComputeAbsoluteFontSize(XElement element, CssExpression assignedValue, HtmlToWmlConverterSettings settings)
        {
            if (assignedValue.Terms.Count != 1)
            {
                throw new OpenXmlPowerToolsException("Should not have a unit with more than one term, I think");
            }

            var value = assignedValue.Terms.First().Value;
            var unit = assignedValue.Terms.First().Unit;
            if (unit == CssUnit.PT)
            {
                return assignedValue;
            }

            if (FontSizeMap.ContainsKey(value))
            {
                return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = FontSizeMap[value].ToString(CultureInfo.InvariantCulture), Type = CssTermType.Number, Unit = CssUnit.PT, } } };
            }

            // todo what should the calculation be for computing larger / smaller?
            if (value == "larger" || value == "smaller")
            {
                var parentFontSize = GetComputedPropertyValue(null, element.Parent, "font-size", settings);
                if (!double.TryParse(parentFontSize.Terms.First().Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var ptSize))
                {
                    throw new OpenXmlPowerToolsException("did not return a double?");
                }

                double newPtSize2 = 0;
                if (value == "larger")
                {
                    if (ptSize < 10)
                    {
                        newPtSize2 = 10d;
                    }

                    if (ptSize == 10 || ptSize == 11)
                    {
                        newPtSize2 = 12d;
                    }

                    if (ptSize == 12)
                    {
                        newPtSize2 = 13.5d;
                    }

                    if (ptSize >= 13 && ptSize <= 15)
                    {
                        newPtSize2 = 18d;
                    }

                    if (ptSize >= 16 && ptSize <= 20)
                    {
                        newPtSize2 = 24d;
                    }

                    if (ptSize >= 21)
                    {
                        newPtSize2 = 36d;
                    }
                }
                if (value == "smaller")
                {
                    if (ptSize <= 11)
                    {
                        newPtSize2 = 7.5d;
                    }

                    if (ptSize == 12)
                    {
                        newPtSize2 = 10d;
                    }

                    if (ptSize >= 13 && ptSize <= 15)
                    {
                        newPtSize2 = 12d;
                    }

                    if (ptSize >= 16 && ptSize <= 20)
                    {
                        newPtSize2 = 13.5d;
                    }

                    if (ptSize >= 21 && ptSize <= 29)
                    {
                        newPtSize2 = 18d;
                    }

                    if (ptSize >= 30)
                    {
                        newPtSize2 = 24d;
                    }
                }
                return new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = newPtSize2.ToString(CultureInfo.InvariantCulture), Type = CssTermType.Number, Unit = CssUnit.PT, } } };
            }
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var decValue))
            {
                throw new OpenXmlPowerToolsException("em value did not parse");
            }

            double? newPtSize = null;
            if (unit == CssUnit.EM || unit == CssUnit.EX || unit == CssUnit.Percent)
            {
                var parentFontSize = GetComputedPropertyValue(null, element.Parent, "font-size", settings);
                if (!double.TryParse(parentFontSize.Terms.First().Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var ptSize))
                {
                    throw new OpenXmlPowerToolsException("did not return a double?");
                }

                if (unit == CssUnit.EM)
                {
                    newPtSize = ptSize * decValue;
                }

                if (unit == CssUnit.EX)
                {
                    newPtSize = ptSize / 2 * decValue;
                }

                if (unit == CssUnit.Percent)
                {
                    newPtSize = ptSize * decValue / 100d;
                }
            }
            else
            {
                if (unit == CssUnit.IN)
                {
                    newPtSize = decValue * 72.0d;
                }

                if (unit == CssUnit.CM)
                {
                    newPtSize = decValue / 2.54d * 72.0d;
                }

                if (unit == CssUnit.MM)
                {
                    newPtSize = decValue / 25.4d * 72.0d;
                }

                if (unit == CssUnit.PC)
                {
                    newPtSize = decValue * 12d;
                }

                if (unit == CssUnit.PX)
                {
                    newPtSize = decValue * 0.75d;
                }
            }
            if (!newPtSize.HasValue)
            {
                throw new OpenXmlPowerToolsException("Internal error: should not have reached this exception");
            }

            var newExpr = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = newPtSize.Value.ToString(CultureInfo.InvariantCulture), Type = CssTermType.Number, Unit = CssUnit.PT, } } };
            return newExpr;
        }

        private static readonly Dictionary<string, double> FontSizeMap = new Dictionary<string, double>()
        {
            { "xx-small", 7.5d },
            { "x-small", 10d },
            { "small", 12d },
            { "medium", 13.5d },
            { "large", 18d },
            { "x-large", 24d },
            { "xx-large", 36d },
        };

        private static void ApplySelector(
            CssSelector selector,
            CssRuleSet ruleSet,
            XElement xHtml,
            Property.HighOrderPriority notImportantHighOrderSort,
            Property.HighOrderPriority importantHighOrderSort,
            ref int propertySequence)
        {
            foreach (var element in xHtml.DescendantsAndSelf())
            {
                if (DoesSelectorMatch(selector, element))
                {
                    foreach (var declaration in ruleSet.Declarations)
                    {
                        var prop = new Property()
                        {
                            Name = declaration.Name.ToLower(),
                            Expression = declaration.Expression,
                            HighOrderSort = declaration.Important ? importantHighOrderSort : notImportantHighOrderSort,
                            IdAttributesInSelector = CountIdAttributesInSelector(selector),
                            AttributesInSelector = CountAttributesInSelector(selector),
                            ElementNamesInSelector = CountElementNamesInSelector(selector),
                            SequenceNumber = propertySequence++,
                        };
                        AddPropertyToElement(element, prop);
                    }
                }
            }
        }

        private static bool DoesSelectorMatch(
            CssSelector selector,
            XElement element)
        {
            var currentSimpleSelector = selector.SimpleSelectors.Count - 1;
            var currentElement = element;
            while (true)
            {
                if (!DoesSimpleSelectorMatch(selector.SimpleSelectors[currentSimpleSelector], currentElement))
                {
                    return false;
                }

                if (currentSimpleSelector == 0)
                {
                    return true;
                }

                if (selector.SimpleSelectors[currentSimpleSelector].Combinator == CssCombinator.ChildOf)
                {
                    currentElement = currentElement.Parent;
                    if (currentElement == null)
                    {
                        return false;
                    }

                    currentSimpleSelector--;
                    continue;
                }
                if (selector.SimpleSelectors[currentSimpleSelector].Combinator == CssCombinator.PrecededImmediatelyBy)
                {
                    currentElement = currentElement.ElementsBeforeSelf().Reverse().FirstOrDefault();
                    if (currentElement == null)
                    {
                        return false;
                    }

                    currentSimpleSelector--;
                    continue;
                }
                if (selector.SimpleSelectors[currentSimpleSelector].Combinator == null)
                {
                    var continueOuter = false;
                    foreach (var ancestor in element.Ancestors())
                    {
                        if (DoesSimpleSelectorMatch(selector.SimpleSelectors[currentSimpleSelector - 1], ancestor))
                        {
                            currentElement = ancestor;
                            currentSimpleSelector--;
                            continueOuter = true;
                            break;
                        }
                    }
                    if (continueOuter)
                    {
                        continue;
                    }

                    return false;
                }
            }
        }

        private static bool DoesSimpleSelectorMatch(
            CssSimpleSelector simpleSelector,
            XElement element)
        {
            var elemantNameMatch = true;
            var classNameMatch = true;
            var childSimpleSelectorMatch = true;
            var idMatch = true;
            var attributeMatch = true;

            if (simpleSelector.Pseudo != null)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(simpleSelector.ElementName) && simpleSelector.ElementName != "*")
            {
                elemantNameMatch = element.Name.ToString() == simpleSelector.ElementName.ToString();
            }

            if (elemantNameMatch)
            {
                if (simpleSelector.Class != null && simpleSelector.Class != "")
                {
                    classNameMatch = ClassesOf(element).Contains(simpleSelector.Class);
                }

                if (classNameMatch)
                {
                    if (simpleSelector.Child != null)
                    {
                        childSimpleSelectorMatch = DoesSimpleSelectorMatch(simpleSelector.Child, element);
                    }

                    if (childSimpleSelectorMatch)
                    {
                        if (simpleSelector.ID != null && simpleSelector.ID != "")
                        {
                            var id = (string)element.Attribute("ID");
                            if (id == null)
                            {
                                id = (string)element.Attribute("id");
                            }

                            idMatch = simpleSelector.ID == id;
                        }
                        if (idMatch)
                        {
                            if (simpleSelector.Attribute != null)
                            {
                                attributeMatch = DoesAttributeMatch(simpleSelector.Attribute, element);
                            }
                        }
                    }
                }
            }
            var result =
                elemantNameMatch &&
                classNameMatch &&
                childSimpleSelectorMatch &&
                idMatch &&
                attributeMatch;
            return result;
        }

        private static bool DoesAttributeMatch(CssAttribute attribute, XElement element)
        {
            var attName = attribute.Operand.ToLower();
            var attValue = (string)element.Attribute(attName);
            if (attValue == null)
            {
                return false;
            }

            if (attribute.Operator == null)
            {
                return true;
            }

            var value = attribute.Value;
            switch (attribute.Operator)
            {
                case CssAttributeOperator.Equals:
                    return attValue == value;

                case CssAttributeOperator.BeginsWith:
                    return attValue.StartsWith(value);

                case CssAttributeOperator.Contains:
                    return attValue.Contains(value);

                case CssAttributeOperator.EndsWith:
                    return attValue.EndsWith(value);

                case CssAttributeOperator.InList:
                    return attValue.Split(' ').Contains(value);

                case CssAttributeOperator.Hyphenated:
                    return attValue.Split('-')[0] == value;

                default:
                    return false;
            }
        }

        private static int CountIdAttributesInSimpleSelector(CssSimpleSelector simpleSelector)
        {
            var count = simpleSelector.ID != null ? 1 : 0 +
                (simpleSelector.Child != null ? CountIdAttributesInSimpleSelector(simpleSelector.Child) : 0);
            return count;
        }

        private static int CountIdAttributesInSelector(CssSelector selector)
        {
            var count = selector.SimpleSelectors.Select(ss => CountIdAttributesInSimpleSelector(ss)).Sum();
            return count;
        }

        private static int CountAttributesInSimpleSelector(CssSimpleSelector simpleSelector)
        {
            var count = (simpleSelector.Attribute != null ? 1 : 0) +
                (simpleSelector.Class != null && simpleSelector.Class != "" ? 1 : 0) +
                (simpleSelector.Child != null ? CountAttributesInSimpleSelector(simpleSelector.Child) : 0);
            return count;
        }

        private static int CountAttributesInSelector(CssSelector selector)
        {
            var count = selector.SimpleSelectors.Select(ss => CountAttributesInSimpleSelector(ss)).Sum();
            return count;
        }

        private static int CountElementNamesInSimpleSelector(CssSimpleSelector simpleSelector)
        {
            var count = simpleSelector.ElementName != null &&
                    simpleSelector.ElementName != "" &&
                    simpleSelector.ElementName != "*"
                    ? 1 : 0 +
                (simpleSelector.Child != null ? CountElementNamesInSimpleSelector(simpleSelector.Child) : 0);
            return count;
        }

        private static int CountElementNamesInSelector(CssSelector selector)
        {
            var count = selector.SimpleSelectors.Select(ss => CountElementNamesInSimpleSelector(ss)).Sum();
            return count;
        }

        private static void AddPropertyToElement(
            XElement element,
            Property property)
        {
            var propList = element.Annotation<Dictionary<string, Property>>();
            if (propList == null)
            {
                propList = new Dictionary<string, Property>();
                element.AddAnnotation(propList);
            }
            if (!propList.ContainsKey(property.Name))
            {
                propList.Add(property.Name, property);
            }
            else
            {
                var current = propList[property.Name];
                if (((IComparable<Property>)property).CompareTo(current) == 1)
                {
                    propList[property.Name] = property;
                }
            }
        }

        private static void AddPropertyToDictionary(
            Dictionary<string, Property> propList,
            Property property)
        {
            if (!propList.ContainsKey(property.Name))
            {
                propList.Add(property.Name, property);
            }
            else
            {
                var current = propList[property.Name];
                if (((IComparable<Property>)property).CompareTo(current) == 1)
                {
                    propList[property.Name] = property;
                }
            }
        }

        private static string[] ClassesOf(XElement element)
        {
            var classesString = (string)element.Attribute("class");
            if (classesString == null)
            {
                return Array.Empty<string>();
            }

            return classesString.Split(' ');
        }

        private static void ApplyDeclarationsToElement(
            CssRuleSet ruleSet,
            XElement element,
            Property.HighOrderPriority notImportantHighOrderSort,
            Property.HighOrderPriority importantHighOrderSort,
            ref int propertySequence)
        {
            foreach (var declaration in ruleSet.Declarations)
            {
                var prop = new Property()
                {
                    Name = declaration.Name.ToLower(),
                    Expression = declaration.Expression,
                    HighOrderSort = declaration.Important ? importantHighOrderSort : notImportantHighOrderSort,
                    IdAttributesInSelector = 0,
                    AttributesInSelector = 0,
                    ElementNamesInSelector = 0,
                    SequenceNumber = propertySequence++,
                };
                AddPropertyToElement(element, prop);
            }
        }

        private static void ApplyCssToElement(
            CssDocument cssDoc,
            XElement element,
            Property.HighOrderPriority notImportantHighOrderSort,
            Property.HighOrderPriority importantHighOrderSort,
            ref int propertySequence)
        {
            foreach (var ruleSet in cssDoc.RuleSets)
            {
                ApplyDeclarationsToElement(ruleSet, element, notImportantHighOrderSort, importantHighOrderSort, ref propertySequence);
            }
        }

        private static void ApplyStyleAttributes(XElement xHtml, ref int propertySequence)
        {
            foreach (var element in xHtml.DescendantsAndSelf())
            {
                var styleAtt = element.Attribute(XhtmlNoNamespace.style);
                if (styleAtt != null)
                {
                    var style = (string)styleAtt;
                    var cssString = element.Name + "{" + style + "}";
                    cssString = cssString.Replace('\"', '\'');
                    var cssParser = new CssParser();
                    var cssDoc = cssParser.ParseText(cssString);
                    ApplyCssToElement(
                        cssDoc,
                        element,
                        Property.HighOrderPriority.StyleAttributeNormal,
                        Property.HighOrderPriority.StyleAttributeHigh,
                        ref propertySequence);
                }
                var dirAtt = element.Attribute(XhtmlNoNamespace.dir);
                if (dirAtt != null)
                {
                    var dir = dirAtt.Value.ToLower();
                    var prop = new Property()
                    {
                        Name = "direction",
                        Expression = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = dir, Type = CssTermType.String } } },
                        HighOrderSort = Property.HighOrderPriority.HtmlAttribute,
                        IdAttributesInSelector = 0,
                        AttributesInSelector = 0,
                        ElementNamesInSelector = 0,
                        SequenceNumber = propertySequence++,
                    };
                    AddPropertyToElement(element, prop);
                }
            }
        }

        private enum CssDataType
        {
            BorderWidth,
            BorderStyle,
            Color,
            ListStyleType,
            ListStylePosition,
            ListStyleImage,
            BackgroundColor,
            BackgroundImage,
            BackgroundRepeat,
            BackgroundAttachment,
            BackgroundPosition,
            FontStyle,
            FontVarient,
            FontWeight,
            FontSize,
            LineHeight,
            FontFamily,
            Length,
        };

        private class ShorthandPropertiesInfo
        {
            public string Name;
            public string Pattern;
        }

        private static readonly ShorthandPropertiesInfo[] ShorthandProperties = new[]
        {
            new ShorthandPropertiesInfo
            {
                Name = "margin",
                Pattern = "margin-{0}",
            },
            new ShorthandPropertiesInfo
            {
                Name = "padding",
                Pattern = "padding-{0}",
            },
            new ShorthandPropertiesInfo
            {
                Name = "border-width",
                Pattern = "border-{0}-width",
            },
            new ShorthandPropertiesInfo
            {
                Name = "border-color",
                Pattern = "border-{0}-color",
            },
            new ShorthandPropertiesInfo
            {
                Name = "border-style",
                Pattern = "border-{0}-style",
            },
        };

        private static void ExpandShorthandProperties(XElement xHtml, HtmlToWmlConverterSettings settings)
        {
            foreach (var element in xHtml.DescendantsAndSelf())
            {
                ExpandShorthandPropertiesForElement(element);
            }
        }

        private static void ExpandShorthandPropertiesForElement(XElement element)
        {
            var propertyList = element.Annotation<Dictionary<string, Property>>();
            if (propertyList == null)
            {
                propertyList = new Dictionary<string, Property>();
                element.AddAnnotation(propertyList);
            }
            foreach (var kvp in propertyList.ToList())
            {
                var p = kvp.Value;
                if (p.Name == "border")
                {
                    CssExpression borderColor;
                    CssExpression borderWidth;
                    CssExpression borderStyle;
                    if (p.Expression.Terms.Count == 1 && p.Expression.Terms.First().Value == "inherit")
                    {
                        borderColor = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        borderWidth = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        borderStyle = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                    }
                    else
                    {
                        borderColor = null;
                        borderWidth = null;
                        borderStyle = null;
                        foreach (var term in p.Expression.Terms)
                        {
                            var dataType = GetDatatypeFromBorderTerm(term);
                            switch (dataType)
                            {
                                case CssDataType.Color:
                                    borderColor = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.BorderWidth:
                                    borderWidth = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.BorderStyle:
                                    borderStyle = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;
                            }
                        }
                    }
                    foreach (var side in new[] { "top", "left", "bottom", "right" })
                    {
                        if (borderWidth != null)
                        {
                            var bwp = new Property
                            {
                                Name = "border-" + side + "-width",
                                Expression = borderWidth,
                                HighOrderSort = p.HighOrderSort,
                                IdAttributesInSelector = p.IdAttributesInSelector,
                                ElementNamesInSelector = p.ElementNamesInSelector,
                                AttributesInSelector = p.AttributesInSelector,
                                SequenceNumber = p.SequenceNumber,
                            };
                            AddPropertyToDictionary(propertyList, bwp);
                        }
                        if (borderStyle != null)
                        {
                            var bsp = new Property
                            {
                                Name = "border-" + side + "-style",
                                Expression = borderStyle,
                                HighOrderSort = p.HighOrderSort,
                                IdAttributesInSelector = p.IdAttributesInSelector,
                                ElementNamesInSelector = p.ElementNamesInSelector,
                                AttributesInSelector = p.AttributesInSelector,
                                SequenceNumber = p.SequenceNumber,
                            };
                            AddPropertyToDictionary(propertyList, bsp);
                        }
                        if (borderColor != null)
                        {
                            var bc = new Property
                            {
                                Name = "border-" + side + "-color",
                                Expression = borderColor,
                                HighOrderSort = p.HighOrderSort,
                                IdAttributesInSelector = p.IdAttributesInSelector,
                                ElementNamesInSelector = p.ElementNamesInSelector,
                                AttributesInSelector = p.AttributesInSelector,
                                SequenceNumber = p.SequenceNumber,
                            };
                            AddPropertyToDictionary(propertyList, bc);
                        }
                    }
                    continue;
                }
                if (p.Name == "border-top" || p.Name == "border-right" || p.Name == "border-bottom" || p.Name == "border-left")
                {
                    CssExpression borderColor;
                    CssExpression borderWidth;
                    CssExpression borderStyle;
                    if (p.Expression.Terms.Count == 1 && p.Expression.Terms.First().Value == "inherit")
                    {
                        borderColor = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        borderWidth = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        borderStyle = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                    }
                    else
                    {
                        borderColor = null;
                        borderWidth = null;
                        borderStyle = null;
                        foreach (var term in p.Expression.Terms)
                        {
                            var dataType = GetDatatypeFromBorderTerm(term);
                            switch (dataType)
                            {
                                case CssDataType.Color:
                                    borderColor = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.BorderWidth:
                                    borderWidth = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.BorderStyle:
                                    borderStyle = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;
                            }
                        }
                    }
                    if (borderWidth != null)
                    {
                        var bwp = new Property
                        {
                            Name = p.Name + "-width",
                            Expression = borderWidth,
                            HighOrderSort = p.HighOrderSort,
                            IdAttributesInSelector = p.IdAttributesInSelector,
                            ElementNamesInSelector = p.ElementNamesInSelector,
                            AttributesInSelector = p.AttributesInSelector,
                            SequenceNumber = p.SequenceNumber,
                        };
                        AddPropertyToDictionary(propertyList, bwp);
                    }
                    if (borderStyle != null)
                    {
                        var bsp = new Property
                        {
                            Name = p.Name + "-style",
                            Expression = borderStyle,
                            HighOrderSort = p.HighOrderSort,
                            IdAttributesInSelector = p.IdAttributesInSelector,
                            ElementNamesInSelector = p.ElementNamesInSelector,
                            AttributesInSelector = p.AttributesInSelector,
                            SequenceNumber = p.SequenceNumber,
                        };
                        AddPropertyToDictionary(propertyList, bsp);
                    }
                    if (borderColor != null)
                    {
                        var bc = new Property
                        {
                            Name = p.Name + "-color",
                            Expression = borderColor,
                            HighOrderSort = p.HighOrderSort,
                            IdAttributesInSelector = p.IdAttributesInSelector,
                            ElementNamesInSelector = p.ElementNamesInSelector,
                            AttributesInSelector = p.AttributesInSelector,
                            SequenceNumber = p.SequenceNumber,
                        };
                        AddPropertyToDictionary(propertyList, bc);
                    }
                    continue;
                }

                if (p.Name == "list-style")
                {
                    CssExpression listStyleType;
                    CssExpression listStylePosition;
                    CssExpression listStyleImage;
                    if (p.Expression.Terms.Count == 1 && p.Expression.Terms.First().Value == "inherit")
                    {
                        listStyleType = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        listStylePosition = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        listStyleImage = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                    }
                    else
                    {
                        listStyleType = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "disc", Type = CssTermType.String } } };
                        listStylePosition = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "outside", Type = CssTermType.String } } };
                        listStyleImage = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String } } };
                        foreach (var term in p.Expression.Terms)
                        {
                            var dataType = GetDatatypeFromListStyleTerm(term);
                            switch (dataType)
                            {
                                case CssDataType.ListStyleType:
                                    listStyleType = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.ListStylePosition:
                                    listStylePosition = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.ListStyleImage:
                                    listStyleImage = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;
                            }
                        }
                    }
                    var lst = new Property
                    {
                        Name = "list-style-type",
                        Expression = listStyleType,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, lst);
                    var lsp = new Property
                    {
                        Name = "list-style-position",
                        Expression = listStylePosition,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, lsp);
                    var lsi = new Property
                    {
                        Name = "list-style-image",
                        Expression = listStyleImage,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, lsi);
                    continue;
                }

                if (p.Name == "background")
                {
                    CssExpression backgroundColor;
                    CssExpression backgroundImage;
                    CssExpression backgroundRepeat;
                    CssExpression backgroundAttachment;
                    CssExpression backgroundPosition;
                    if (p.Expression.Terms.Count == 1 && p.Expression.Terms.First().Value == "inherit")
                    {
                        backgroundColor = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        backgroundImage = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        backgroundRepeat = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        backgroundAttachment = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        backgroundPosition = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                    }
                    else
                    {
                        backgroundColor = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "transparent", Type = CssTermType.String } } };
                        backgroundImage = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "none", Type = CssTermType.String } } };
                        backgroundRepeat = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "repeat", Type = CssTermType.String } } };
                        backgroundAttachment = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "scroll", Type = CssTermType.String } } };
                        backgroundPosition = new CssExpression
                        {
                            Terms = new List<CssTerm> {
                            new CssTerm {
                                Value = "0",
                                Unit = CssUnit.Percent,
                                Type = CssTermType.Number },
                            new CssTerm {
                                Value = "0",
                                Unit = CssUnit.Percent,
                                Type = CssTermType.Number },
                        }
                        };
                        var backgroundPositionList = new List<CssTerm>();
                        foreach (var term in p.Expression.Terms)
                        {
                            var dataType = GetDatatypeFromBackgroundTerm(term);
                            switch (dataType)
                            {
                                case CssDataType.BackgroundColor:
                                    backgroundColor = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.BackgroundImage:
                                    backgroundImage = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.BackgroundRepeat:
                                    backgroundRepeat = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.BackgroundAttachment:
                                    backgroundAttachment = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.BackgroundPosition:
                                    backgroundPositionList.Add(term);
                                    break;
                            }
                        }
                        if (backgroundPositionList.Count == 1)
                        {
                            backgroundPosition = new CssExpression
                            {
                                Terms = new List<CssTerm> {
                                backgroundPositionList.First(),
                                new CssTerm {
                                    Value = "center",
                                    Type = CssTermType.String
                                },
                            }
                            };
                        }
                        if (backgroundPositionList.Count == 2)
                        {
                            backgroundPosition = new CssExpression
                            {
                                Terms = backgroundPositionList
                            };
                        }
                    }
                    var bc = new Property
                    {
                        Name = "background-color",
                        Expression = backgroundColor,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, bc);
                    var bgi = new Property
                    {
                        Name = "background-image",
                        Expression = backgroundImage,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, bgi);
                    var bgr = new Property
                    {
                        Name = "background-repeat",
                        Expression = backgroundRepeat,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, bgr);
                    var bga = new Property
                    {
                        Name = "background-attachment",
                        Expression = backgroundAttachment,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, bga);
                    var bgp = new Property
                    {
                        Name = "background-position",
                        Expression = backgroundPosition,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, bgp);
                    continue;
                }

                if (p.Name == "font")
                {
                    CssExpression fontStyle;
                    CssExpression fontVarient;
                    CssExpression fontWeight;
                    CssExpression fontSize;
                    CssExpression lineHeight;
                    CssExpression fontFamily;
                    if (p.Expression.Terms.Count == 1 && p.Expression.Terms.First().Value == "inherit")
                    {
                        fontStyle = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        fontVarient = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        fontWeight = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        fontSize = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        lineHeight = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                        fontFamily = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "inherit", Type = CssTermType.String } } };
                    }
                    else
                    {
                        fontStyle = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } };
                        fontVarient = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } };
                        fontWeight = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } };
                        fontSize = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "medium", Type = CssTermType.String } } };
                        lineHeight = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "normal", Type = CssTermType.String } } };
                        fontFamily = new CssExpression { Terms = new List<CssTerm> { new CssTerm { Value = "serif", Type = CssTermType.String } } };
                        var fontFamilyList = new List<CssTerm>();
                        foreach (var term in p.Expression.Terms)
                        {
                            var dataType = GetDatatypeFromFontTerm(term);
                            switch (dataType)
                            {
                                case CssDataType.FontStyle:
                                    fontStyle = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.FontVarient:
                                    fontVarient = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.FontWeight:
                                    fontWeight = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.FontSize:
                                    fontSize = new CssExpression { Terms = new List<CssTerm> { term } };
                                    break;

                                case CssDataType.Length:
                                    if (term.SeparatorChar == "/")
                                    {
                                        lineHeight = new CssExpression { Terms = new List<CssTerm> { term } };
                                    }
                                    else
                                    {
                                        fontSize = new CssExpression { Terms = new List<CssTerm> { term } };
                                    }

                                    break;

                                case CssDataType.FontFamily:
                                    fontFamilyList.Add(term);
                                    break;
                            }
                        }
                        if (fontFamilyList.Count > 0)
                        {
                            fontFamily = new CssExpression { Terms = fontFamilyList };
                        }
                    }
                    var fs = new Property
                    {
                        Name = "font-style",
                        Expression = fontStyle,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, fs);
                    var fv = new Property
                    {
                        Name = "font-varient",
                        Expression = fontVarient,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, fv);
                    var fw = new Property
                    {
                        Name = "font-weight",
                        Expression = fontWeight,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, fw);
                    var fsz = new Property
                    {
                        Name = "font-size",
                        Expression = fontSize,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, fsz);
                    var lh = new Property
                    {
                        Name = "line-height",
                        Expression = lineHeight,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, lh);
                    var ff = new Property
                    {
                        Name = "font-family",
                        Expression = fontFamily,
                        HighOrderSort = p.HighOrderSort,
                        IdAttributesInSelector = p.IdAttributesInSelector,
                        ElementNamesInSelector = p.ElementNamesInSelector,
                        AttributesInSelector = p.AttributesInSelector,
                        SequenceNumber = p.SequenceNumber,
                    };
                    AddPropertyToDictionary(propertyList, ff);
                    continue;
                }

                foreach (var shPr in ShorthandProperties)
                {
                    if (p.Name == shPr.Name)
                    {
                        switch (p.Expression.Terms.Count)
                        {
                            case 1:
                                foreach (var direction in new[] { "top", "right", "bottom", "left" })
                                {
                                    var ep = new Property()
                                    {
                                        Name = string.Format(shPr.Pattern, direction),
                                        Expression = new CssExpression { Terms = new List<CssTerm> { p.Expression.Terms.First() } },
                                        HighOrderSort = p.HighOrderSort,
                                        IdAttributesInSelector = p.IdAttributesInSelector,
                                        AttributesInSelector = p.AttributesInSelector,
                                        ElementNamesInSelector = p.ElementNamesInSelector,
                                        SequenceNumber = p.SequenceNumber,
                                    };
                                    AddPropertyToDictionary(propertyList, ep);
                                }
                                break;

                            case 2:
                                foreach (var direction in new[] { "top", "bottom" })
                                {
                                    var ep = new Property()
                                    {
                                        Name = string.Format(shPr.Pattern, direction),
                                        Expression = new CssExpression { Terms = new List<CssTerm> { p.Expression.Terms.First() } },
                                        HighOrderSort = p.HighOrderSort,
                                        IdAttributesInSelector = p.IdAttributesInSelector,
                                        AttributesInSelector = p.AttributesInSelector,
                                        ElementNamesInSelector = p.ElementNamesInSelector,
                                        SequenceNumber = p.SequenceNumber,
                                    };
                                    AddPropertyToDictionary(propertyList, ep);
                                }
                                foreach (var direction in new[] { "left", "right" })
                                {
                                    var ep = new Property()
                                    {
                                        Name = string.Format(shPr.Pattern, direction),
                                        Expression = new CssExpression { Terms = new List<CssTerm> { p.Expression.Terms.Skip(1).First() } },
                                        HighOrderSort = p.HighOrderSort,
                                        IdAttributesInSelector = p.IdAttributesInSelector,
                                        AttributesInSelector = p.AttributesInSelector,
                                        ElementNamesInSelector = p.ElementNamesInSelector,
                                        SequenceNumber = p.SequenceNumber,
                                    };
                                    AddPropertyToDictionary(propertyList, ep);
                                }
                                break;

                            case 3:
                                var ep3 = new Property()
                                {
                                    Name = string.Format(shPr.Pattern, "top"),
                                    Expression = new CssExpression { Terms = new List<CssTerm> { p.Expression.Terms.First() } },
                                    HighOrderSort = p.HighOrderSort,
                                    IdAttributesInSelector = p.IdAttributesInSelector,
                                    AttributesInSelector = p.AttributesInSelector,
                                    ElementNamesInSelector = p.ElementNamesInSelector,
                                    SequenceNumber = p.SequenceNumber,
                                };
                                AddPropertyToDictionary(propertyList, ep3);
                                foreach (var direction in new[] { "left", "right" })
                                {
                                    var ep2 = new Property()
                                    {
                                        Name = string.Format(shPr.Pattern, direction),
                                        Expression = new CssExpression { Terms = new List<CssTerm> { p.Expression.Terms.Skip(1).First() } },
                                        HighOrderSort = p.HighOrderSort,
                                        IdAttributesInSelector = p.IdAttributesInSelector,
                                        AttributesInSelector = p.AttributesInSelector,
                                        ElementNamesInSelector = p.ElementNamesInSelector,
                                        SequenceNumber = p.SequenceNumber,
                                    };
                                    AddPropertyToDictionary(propertyList, ep2);
                                }
                                var ep4 = new Property()
                                {
                                    Name = string.Format(shPr.Pattern, "bottom"),
                                    Expression = new CssExpression { Terms = new List<CssTerm> { p.Expression.Terms.Skip(2).First() } },
                                    HighOrderSort = p.HighOrderSort,
                                    IdAttributesInSelector = p.IdAttributesInSelector,
                                    AttributesInSelector = p.AttributesInSelector,
                                    ElementNamesInSelector = p.ElementNamesInSelector,
                                    SequenceNumber = p.SequenceNumber,
                                };
                                AddPropertyToDictionary(propertyList, ep4);
                                break;

                            case 4:
                                var skip = 0;
                                foreach (var direction in new[] { "top", "right", "bottom", "left" })
                                {
                                    var ep = new Property()
                                    {
                                        Name = string.Format(shPr.Pattern, direction),
                                        Expression = new CssExpression { Terms = new List<CssTerm> { p.Expression.Terms.Skip(skip++).First() } },
                                        HighOrderSort = p.HighOrderSort,
                                        IdAttributesInSelector = p.IdAttributesInSelector,
                                        AttributesInSelector = p.AttributesInSelector,
                                        ElementNamesInSelector = p.ElementNamesInSelector,
                                        SequenceNumber = p.SequenceNumber,
                                    };
                                    AddPropertyToDictionary(propertyList, ep);
                                }
                                break;
                        }
                    }
                }
            }
        }

        private static readonly string[] BackgroundRepeatValues = new[]
        {
            "repeat",
            "repeat-x",
            "repeat-y",
            "no-repeat",
        };

        private static readonly string[] BackgroundAttachmentValues = new[]
        {
            "scroll",
            "fixed",
        };

        private static readonly string[] BackgroundPositionValues = new[]
        {
            "left",
            "right",
            "top",
            "bottom",
            "center",
        };

        private static CssDataType GetDatatypeFromBackgroundTerm(CssTerm term)
        {
            if (term.IsColor)
            {
                return CssDataType.BackgroundColor;
            }

            if (BackgroundRepeatValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.BackgroundRepeat;
            }

            if (BackgroundAttachmentValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.BackgroundAttachment;
            }

            if (term.Function != null)
            {
                return CssDataType.BackgroundImage;
            }

            if (term.Unit == CssUnit.CM ||
                term.Unit == CssUnit.EM ||
                term.Unit == CssUnit.IN ||
                term.Unit == CssUnit.MM ||
                term.Unit == CssUnit.PT ||
                term.Unit == CssUnit.PX ||
                term.Unit == CssUnit.Percent)
            {
                return CssDataType.BackgroundPosition;
            }

            if (BackgroundPositionValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.BackgroundPosition;
            }

            return CssDataType.BackgroundPosition;
        }

        private static readonly string[] ListStylePositionValues = new[]
        {
            "inside",
            "outside",
        };

        private static readonly string[] BorderStyleValues = new[]
        {
            "none",
            "hidden",
            "dotted",
            "dashed",
            "solid",
            "double",
            "groove",
            "ridge",
            "inset",
            "outset",
        };

        private static CssDataType GetDatatypeFromBorderTerm(CssTerm term)
        {
            if (term.IsColor)
            {
                return CssDataType.Color;
            }
            if (BorderStyleValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.BorderStyle;
            }

            return CssDataType.BorderWidth;
        }

        private static readonly string[] ListStyleTypeValues = new[]
        {
            "disc",
            "circle",
            "square",
            "decimal",
            "decimal-leading-zero",
            "lower-roman",
            "upper-roman",
            "lower-greek",
            "lower-latin",
            "upper-latin",
            "armenian",
            "georgian",
            "lower-alpha",
            "upper-alpha",
        };

        private static CssDataType GetDatatypeFromListStyleTerm(CssTerm term)
        {
            if (ListStyleTypeValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.ListStyleType;
            }

            if (ListStylePositionValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.ListStylePosition;
            }

            return CssDataType.ListStyleImage;
        }

        private static readonly string[] FontStyleValues = new[]
        {
            "italic",
            "oblique",
        };

        private static readonly string[] FontVarientValues = new[]
        {
            "small-caps",
        };

        private static readonly string[] FontWeightValues = new[]
        {
            "bold",
            "bolder",
            "lighter",
            "100",
            "200",
            "300",
            "400",
            "500",
            "600",
            "700",
            "800",
            "900",
        };

        private static CssDataType GetDatatypeFromFontTerm(CssTerm term)
        {
            if (FontStyleValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.FontStyle;
            }

            if (FontVarientValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.FontVarient;
            }

            if (FontWeightValues.Contains(term.Value.ToLower()))
            {
                return CssDataType.FontWeight;
            }

            if (FontSizeMap.ContainsKey(term.Value.ToLower()))
            {
                return CssDataType.FontSize;
            }

            if (term.Unit == CssUnit.CM ||
                term.Unit == CssUnit.EM ||
                term.Unit == CssUnit.IN ||
                term.Unit == CssUnit.MM ||
                term.Unit == CssUnit.PT ||
                term.Unit == CssUnit.PX ||
                term.Unit == CssUnit.Percent)
            {
                return CssDataType.Length;
            }

            return CssDataType.FontFamily;
        }

        public class PropertyInfo
        {
            public string[] Names;
            public bool Inherits;
            public Func<XElement, HtmlToWmlConverterSettings, bool> Includes;
            public Func<XElement, HtmlToWmlConverterSettings, CssExpression> InitialValue;
            public Func<XElement, CssExpression, HtmlToWmlConverterSettings, CssExpression> ComputedValue;
        }

        private static void WriteXHtmlWithAnnotations(XElement element, StringBuilder sb)
        {
            var depth = element.Ancestors().Count() * 2;
            var dummyElement = new XElement(element.Name, element.Attributes());
            sb.Append(string.Format("{0}{1}", "".PadRight(depth), dummyElement) + Environment.NewLine);
            var propList = element.Annotation<Dictionary<string, Property>>();
            if (propList != null)
            {
                sb.Append("".PadRight(depth + 2) + "Properties from Stylesheets" + Environment.NewLine);
                sb.Append("".PadRight(depth + 2) + "===========================" + Environment.NewLine);
                foreach (var kvp in propList.OrderBy(p => p.Key).ThenBy(p => p.Value))
                {
                    var prop = kvp.Value;
                    var propString = string.Format("{0} High:{1} Id:{2} Att:{3} Ell:{4} Seq:{5}",
                        (prop.Name + ":" + prop.Expression + " ").PadRight(50 - depth + 2, '.'), (int)prop.HighOrderSort, prop.IdAttributesInSelector, prop.AttributesInSelector,
                        prop.ElementNamesInSelector, prop.SequenceNumber);
                    sb.Append(string.Format("{0}{1}", "".PadRight(depth + 2), propString) + Environment.NewLine);
                }
                sb.Append(Environment.NewLine);
            }
            var computedProperties = element.Annotation<Dictionary<string, CssExpression>>();
            if (computedProperties != null)
            {
                sb.Append("".PadRight(depth + 2) + "Computed Properties" + Environment.NewLine);
                sb.Append("".PadRight(depth + 2) + "===================" + Environment.NewLine);
                foreach (var prop in computedProperties.OrderBy(cp => cp.Key))
                {
                    var propString = prop.Key + ":" + prop.Value;
                    sb.Append(string.Format("{0}{1}", "".PadRight(depth + 2), propString) + Environment.NewLine);
                }
                sb.Append(Environment.NewLine);
            }
            foreach (var child in element.Elements())
            {
                WriteXHtmlWithAnnotations(child, sb);
            }
        }

        public static string DumpCss(CssDocument css)
        {
            var sb = new StringBuilder();
            var indent = 0;

            Pr(sb, indent, "CSS Tree Dump");
            Pr(sb, indent, "=============");

            Pr(sb, indent, "Directives count: {0}", css.Directives.Count);
            Pr(sb, indent, "RuleSet count: {0}", css.RuleSets.Count);
            foreach (var rs in css.RuleSets)
            {
                DumpRuleSet(sb, indent, rs);
            }

            Pr(sb, indent, "");
            return sb.ToString();
        }

        private static void DumpFunction(StringBuilder sb, int indent, CssFunction f)
        {
            Pr(sb, indent, "Function: {0}", f);
            if (f != null)
            {
                indent++;
                Pr(sb, indent, "Name: {0}", f.Name);
                DumpExpression(sb, indent, f.Expression);
            }
        }

        private static void DumpAttribute(StringBuilder sb, int indent, CssAttribute a)
        {
            Pr(sb, indent, "Attribute: {0}", a);
            if (a != null)
            {
                indent++;
                Pr(sb, indent, "Operand: {0}", a.Operand);
                Pr(sb, indent, "Operator: {0}", a.Operator);
                Pr(sb, indent, "OperatorString: {0}", a.CssOperatorString);
                Pr(sb, indent, "Value: {0}", a.Value);
            }
        }

        private static void DumpSimpleSelector(StringBuilder sb, int indent, CssSimpleSelector s)
        {
            indent++;
            Pr(sb, indent, "SimpleSelector: {0}", s);
            if (s != null)
            {
                indent++;
                DumpAttribute(sb, indent, s.Attribute);
                Pr(sb, indent, "Child: {0}", s.Child);
                DumpSimpleSelector(sb, indent, s.Child);
                Pr(sb, indent, "Class: {0}", s.Class);
                Pr(sb, indent, "Combinator: {0}", s.Combinator);
                Pr(sb, indent, "CombinatorString: {0}", s.CombinatorString);
                Pr(sb, indent, "ElementName: >{0}<", s.ElementName);
                DumpFunction(sb, indent, s.Function);
                Pr(sb, indent, "ID: {0}", s.ID);
                Pr(sb, indent, "Pseudo: {0}", s.Pseudo);
                indent--;
            }
        }

        private static void DumpSelectors(StringBuilder sb, int indent, CssSelector s)
        {
            indent++;
            Pr(sb, indent, "SimpleSelectors count: {0}", s.SimpleSelectors.Count);
            foreach (var ss in s.SimpleSelectors)
            {
                DumpSimpleSelector(sb, indent, ss);
            }
        }

        private static void DumpTerm(StringBuilder sb, int indent, CssTerm t)
        {
            Pr(sb, indent, "Term >{0}<", t.ToString());
            indent++;
            DumpFunction(sb, indent, t.Function);
            Pr(sb, indent, "IsColor: {0}", t.IsColor);
            Pr(sb, indent, "Separator: {0}", t.Separator);
            Pr(sb, indent, "SeparatorChar: {0}", t.SeparatorChar);
            Pr(sb, indent, "Sign: {0}", t.Sign);
            Pr(sb, indent, "SignChar: {0}", t.SignChar);
            Pr(sb, indent, "Type: {0}", t.Type);
            Pr(sb, indent, "Unit: {0}", t.Unit);
            Pr(sb, indent, "UnitString: {0}", t.UnitString);
            Pr(sb, indent, "Value: {0}", t.Value);
        }

        private static void DumpExpression(StringBuilder sb, int indent, CssExpression e)
        {
            Pr(sb, indent, "Expression >{0}<", e.ToString());
            indent++;
            Pr(sb, indent, "Terms count: {0}", e.Terms.Count);
            foreach (var t in e.Terms)
            {
                DumpTerm(sb, indent, t);
            }
        }

        private static void DumpDeclarations(StringBuilder sb, int indent, CssDeclaration d)
        {
            indent++;
            Pr(sb, indent, "Declaration >{0}<", d.ToString());
            indent++;
            Pr(sb, indent, "Name: {0}", d.Name);
            DumpExpression(sb, indent, d.Expression);
            Pr(sb, indent, "Important: {0}", d.Important);
            indent--;
        }

        private static void DumpRuleSet(StringBuilder sb, int indent, CssRuleSet rs)
        {
            indent++;
            Pr(sb, indent, "RuleSet");
            indent++;
            Pr(sb, indent, "Selectors count: {0}", rs.Selectors.Count);
            foreach (var s in rs.Selectors)
            {
                DumpSelectors(sb, indent, s);
            }

            Pr(sb, indent, "Declarations count: {0}", rs.Declarations.Count);
            foreach (var d in rs.Declarations)
            {
                DumpDeclarations(sb, indent, d);
            }

            indent--;
        }

        private static void Pr(StringBuilder sb, int indent, string format, object o)
        {
            if (o == null)
            {
                return;
            }

            var text = string.Format(format, o);
            var sb2 = new StringBuilder("".PadRight(indent * 2) + text);
            sb.Append(sb2);
            sb.Append(Environment.NewLine);
        }

        private static void Pr(StringBuilder sb, int indent, string text)
        {
            var sb2 = new StringBuilder("".PadRight(indent * 2) + text);
            sb.Append(sb2);
            sb.Append(Environment.NewLine);
        }

        public class Property : IComparable<Property>
        {
            public string Name { get; set; }
            public CssExpression Expression { get; set; }
            public HighOrderPriority HighOrderSort { get; set; }
            public int IdAttributesInSelector { get; set; }
            public int AttributesInSelector { get; set; }
            public int ElementNamesInSelector { get; set; }
            public int SequenceNumber { get; set; }

            public enum HighOrderPriority
            {
                InitialValue = 0,
                Inherited = 1,
                UserAgentNormal = 2,
                UserAgentHigh = 3,
                UserNormal = 4,
                AuthorNormal = 5,
                HtmlAttribute = 6,
                StyleAttributeNormal = 7,
                StyleAttributeHigh = 8,
                AuthorHigh = 9,
                UserHigh = 10,
            };

            int IComparable<Property>.CompareTo(Property other)
            {
                // if this is less than other, return -1
                // if this is greater than other, return 1

                var gt = 1;
                var lt = -1;
                if (HighOrderSort < other.HighOrderSort)
                {
                    return lt;
                }

                if (HighOrderSort > other.HighOrderSort)
                {
                    return gt;
                }

                if (IdAttributesInSelector < other.IdAttributesInSelector)
                {
                    return lt;
                }

                if (IdAttributesInSelector > other.IdAttributesInSelector)
                {
                    return gt;
                }

                if (AttributesInSelector < other.AttributesInSelector)
                {
                    return lt;
                }

                if (AttributesInSelector > other.AttributesInSelector)
                {
                    return gt;
                }

                if (ElementNamesInSelector < other.ElementNamesInSelector)
                {
                    return lt;
                }

                if (ElementNamesInSelector > other.ElementNamesInSelector)
                {
                    return gt;
                }

                return SequenceNumber.CompareTo(other.SequenceNumber);
            }
        }

        private static readonly Dictionary<string, string> ColorMap = new Dictionary<string, string>()
        {
            { "maroon", "800000" },
            { "red", "FF0000" },
            { "orange", "FFA500" },
            { "yellow", "FFFF00" },
            { "olive", "808000" },
            { "purple", "800080" },
            { "fuchsia", "FF00FF" },
            { "white", "FFFFFF" },
            { "lime", "00FF00" },
            { "green", "008000" },
            { "navy", "000080" },
            { "blue", "0000FF" },
            { "mediumblue", "0000CD" },
            { "aqua", "00FFFF" },
            { "teal", "008080" },
            { "black", "000000" },
            { "silver", "C0C0C0" },
            { "gray", "808080" },
            { "darkgray", "A9A9A9" },
            { "beige", "F5F5DC" },
            { "windowtext", "000000" },
        };

        public static string GetWmlColorFromExpression(CssExpression color)
        {
            // todo have to handle all forms of colors here
            if (color.Terms.Count == 1)
            {
                var term = color.Terms.First();
                if (term.Type == CssTermType.Function && term.Function.Name.ToUpper() == "RGB" && term.Function.Expression.Terms.Count == 3)
                {
                    var lt = term.Function.Expression.Terms;
                    if (lt.First().Unit == CssUnit.Percent)
                    {
                        var v1 = lt.First().Value;
                        var v2 = lt.ElementAt(1).Value;
                        var v3 = lt.ElementAt(2).Value;
                        var colorInHex = string.Format("{0:x2}{1:x2}{2:x2}", (int)(float.Parse(v1) / 100.0 * 255),
                            (int)(float.Parse(v2) / 100.0 * 255), (int)(float.Parse(v3) / 100.0 * 255));
                        return colorInHex;
                    }
                    else
                    {
                        var v1 = lt.First().Value;
                        var v2 = lt.ElementAt(1).Value;
                        var v3 = lt.ElementAt(2).Value;
                        var colorInHex = string.Format("{0:x2}{1:x2}{2:x2}", int.Parse(v1), int.Parse(v2), int.Parse(v3));
                        return colorInHex;
                    }
                }
                var value = term.Value;
                if (value.Substring(0, 1) == "#" && value.Length == 4)
                {
                    var e = ConvertSingleDigit(value.Substring(1, 1)) +
                        ConvertSingleDigit(value.Substring(2, 1)) +
                        ConvertSingleDigit(value.Substring(3, 1));
                    return e;
                }
                if (value.Substring(0, 1) == "#")
                {
                    return value.Substring(1);
                }

                if (ColorMap.ContainsKey(value))
                {
                    return ColorMap[value];
                }

                return value;
            }
            return "000000";
        }

        private static string ConvertSingleDigit(string singleDigit)
        {
            return singleDigit + singleDigit;
        }
    }
}