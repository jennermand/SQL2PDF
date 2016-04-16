using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SQL2PDFReport
{
    [Serializable]    
    [XmlInclude(typeof(Section))]
    [XmlInclude(typeof(Page))]
    [XmlInclude(typeof(Table))]
    [XmlInclude(typeof(Header))]    
    public class Sections : IEnumerable<Section>
    {
        List<Section> _Items;

        [XmlElement(Type = typeof(Section))]
        [XmlElement(Type = typeof(Page))]
        [XmlElement(Type = typeof(Table))]
        [XmlElement(Type = typeof(Header))]
        public List<Section> Items
        {
            get { return _Items; }
            set { _Items = value; }
        }

        public Sections()
        {
            _Items = new List<Section>();
        }

        public void Add(Section section)
        {
            _Items.Add(section);
        }
       
        public IEnumerator<Section> GetEnumerator()
        {
            return _Items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
