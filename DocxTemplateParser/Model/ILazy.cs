using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxTemplateParser.Model
{

    public interface ILazy
    {
        bool IsLoaded { get; }

        void Load();
    }
}
