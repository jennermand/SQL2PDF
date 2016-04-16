using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SQL2PDFReport
{
    public enum FieldType
    {
        Field, CellField, PhraseField, Text, MultiField
    }

    [Serializable]
    [XmlInclude(typeof(CellFields))]
    [XmlInclude(typeof(PhraseField))]
    [XmlInclude(typeof(TextPhraseField))]   
    public class Fields 
    {
        [XmlAttribute("Key")]
        public string Key { get; set; }
        //public bool Print { get; set; }
       
        protected FieldType _field;

        [XmlIgnore]        
        public FieldType Field
        {
            get { return _field; }
        }

        public Fields()
        {
            _field = FieldType.Field;
        }
    }

    [Serializable]
    [XmlType("CellFields")]
    public class CellFields : Fields
    {
        [XmlAttribute("Width")]
        public int Width { get; set; }

        [XmlAttribute("Header")]
        public string Header { get; set; }

        public CellFields()
        {
            Width = 1;
            _field = FieldType.CellField;
        }
    }

    [Serializable]
    public class PhraseField : Fields
    {
        public PhraseField()
        {
            _field = FieldType.PhraseField;
        }
    }

    [Serializable]
    public class ParagraphField : Fields
    {
        public List<TextPhraseField> Items { get; set; }

        public ParagraphField()
        {
            _field = FieldType.MultiField;
            Items = new List<TextPhraseField>();
            Key = "[paragraph]";
        }
    }

    [Serializable]
    public class TextPhraseField : Fields
    {
        [XmlAttribute("Text")]
        public string Text { get; set; }

        public TextPhraseField()
        {
            _field = FieldType.Text;
            Key = "TEXT";
        }
    }

}
