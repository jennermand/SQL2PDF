using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQL2PDFReport
{
    public enum Style
    {
        normal, bold, italic
    }


    public class Font
    {
        public int Size { get; set; }
        public Style Style { get; set; }
        public string Name { get; set; }
    }
}
