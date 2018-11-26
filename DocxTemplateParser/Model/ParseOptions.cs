using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxTemplateParser.Model
{
    public class ParseOptions
    {
        public Variable Variable { get; set; } = new Variable { Start = "<<", End = ">>" };
        public Variable LoopVariable { get; set; } = new Variable { Start = "<<<", End = ">>>" };
        public string RowNumberVariable = "RowNumber";
    }
    public class Variable
    {
        public string Start { get; set; }
        public string End { get; set; }
        public string Regex { get { return $@"\{Start}(.*?){End}"; } }

    }
}
