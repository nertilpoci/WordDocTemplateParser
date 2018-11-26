using DocxTemplateParser.Extensions;
using DocxTemplateParser.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;
using Font = System.Drawing.Font;

namespace DocxTemplateParser
{
    public class TemplateParser
    {

        public void GenerateInvoice(Stream input,Object data, ParseOptions options=null )
        {
            if (options == null) options = new ParseOptions();
            using (DocX document = DocX.Load(input))
            {
                //do loop replace first to avoid variable confusion

                //get a list of all loop variables
                var loopsToReplace = document.FindUniqueByPattern(options.LoopVariable.Regex, System.Text.RegularExpressions.RegexOptions.Multiline);

                foreach (var loop in loopsToReplace)
                {
                    var loopListName = loop.Replace(options.LoopVariable.Start, string.Empty).Replace(options.LoopVariable.End, string.Empty);
                    var loopData = GetPropValue(data,loopListName) as IEnumerable<Object>;
                    if (loopData != null)
                    {
                        var table = document.Tables.Where(z => z.Rows.Any(r => r.Cells.Any(c => c.Xml.Value.Contains(loop)))).FirstOrDefault();
                        if (table != null)
                        {
                            var row = table.Rows.FirstOrDefault(r => r.Cells.Any(c => c.Xml.Value.Contains(loop)));
                            var index = table.Rows.FindIndex((r => r.Cells.Any(c => c.Xml.Value.Contains(loop))));
                            row.ReplaceText(loop, "");
                            if (table.Rows.Count == 1) table.InsertRow(); // insert row doesn't work when table has zero rows. We add 1 row and remove the original template row  if there is only 1 row in total
                            table.RemoveRow(index);
                            int rowNumber = 1;
                            foreach (var rowData in loopData)
                            {
                                var newRow = table.InsertRow(row, index);
                                foreach (var cell in newRow.Cells)
                                {
                                    var cellValue = cell.Xml.Value;
                                    var variables = Regex.Matches(cellValue, options.Variable.Regex, RegexOptions.Multiline);
                                    foreach (var variable in variables)
                                    {
                                        var variableName = variable.ToString().Replace(options.Variable.Start, string.Empty).Replace(options.Variable.End, string.Empty);
                                        var cellData = variableName == options.RowNumberVariable ? rowNumber : rowData.GetPropValue(variableName);
                                        cell.ReplaceText(variable.ToString(), cellData?.ToString() ?? "");
                                    }
                                }
                                rowNumber++;
                                index++;
                            }
                        }
                    }

                }


                //replace variables
                var variablesToReplace = document.FindUniqueByPattern(options.Variable.Regex, System.Text.RegularExpressions.RegexOptions.Multiline);
                foreach (var variable in variablesToReplace)
                {
                    var variableName = variable.Replace(options.Variable.Start, string.Empty).Replace(options.Variable.End, string.Empty);
                    var cellData = GetPropValue(data,variableName);
                    document.ReplaceText(variable.ToString(), cellData?.ToString() ?? "");
                }
                //end of variable replace
                document.Save();
                document.Dispose();
            }
        }
        public object GetPropValue(object obj, String propName, bool markNotFoundVariables = false)
        {
            string[] nameParts = propName.Split('.');
            if (nameParts.Length == 1)
            {
                var prop = obj.GetType().GetProperty(propName);
                if (prop == null) return markNotFoundVariables ? $"##Invalid Varaible({propName})##" : "";
                return prop.GetValue(obj, null);

            }

            foreach (String part in nameParts)
            {
                if (obj == null) return obj;

                if (obj is ILazy) ((ILazy)obj).Load();

                Type type = obj.GetType();
                PropertyInfo info = type.GetProperty(part);
                if (info == null) return markNotFoundVariables ? $"##Invalid Varaible({propName})##" : "";

                obj = info.GetValue(obj, null);
            }

            return obj;
        }



    }
}
