﻿using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Runtime.Serialization;

namespace SQL2PDFReport
{
    [Serializable]    
    [XmlInclude(typeof(Page))]
    [XmlInclude(typeof(Table))]
    [XmlInclude(typeof(Header))]
    [XmlInclude(typeof(CellFields))]
    [XmlInclude(typeof(PhraseField))]
    [XmlInclude(typeof(TextPhraseField))]
    [XmlInclude(typeof(ParagraphField))]
    public abstract class Section
    {
        [XmlElement(Type = typeof(Section))]
        [XmlElement(Type = typeof(Page))]
        [XmlElement(Type = typeof(Table))]
        [XmlElement(Type = typeof(Header))]
        public List<Section> Sections {get;set;} 
        
        [XmlElement(Type = typeof(Fields))]
        [XmlElement(Type = typeof(CellFields))]
        [XmlElement(Type = typeof(PhraseField))]
        [XmlElement(Type = typeof(TextPhraseField))]
        [XmlElement(typeof(ParagraphField))]
        public  List<Fields> DisplayFields { get; set; }

        public Section()
        {
            Sections = new List<Section>();
            DisplayFields = new List<Fields>();
        }

        public void Add(Section section)
        {
            Sections.Add(section);
        }

        public bool show(string key)
        {
            foreach (var k in DisplayFields)
            {
                if (k.Field == FieldType.Text)                    
                    return true;
                if (k.Key == key )
                    return true;
            }
            return false;
        }
        
        virtual public IEnumerable<Fields> Fields(IEnumerable<Dictionary<string, object>> data)
        {
            

            foreach (var d in data)
            {
                foreach (var k in d.Keys)
                {
                    if (show(k))
                        yield return new Fields();//{ Text= d[k].ToString()};
                }
            }

            foreach (var s in Sections)
            {
                foreach (var s2 in s.Fields(data))
                {
                    yield return s2;
                }
            }
        }
        
        virtual public IEnumerable<Paragraph> Paragraphs(IEnumerable<Dictionary<string, object>> data)
        {
            foreach (var d in data)
            {
                foreach (var k in d.Keys)
                {
                    if (show(k))
                    {
                        yield return new Paragraph(d[k].ToString());
                    }

                    if (isMultiFields(k))
                    {
                        var par = new Paragraph();
                        foreach (var p in getMultifields(k))
                        {
                            par.Add(new Paragraph(d[k].ToString()));
                        }
                        yield return par;
                    }
                }
            }            
        }

        protected IEnumerable<Paragraph> getMultiParagraphs(Fields field, IEnumerable<Dictionary<string, object>> data)
        {
            if (field.Field== FieldType.MultiField)
            {
                foreach (var d in data)
                {
                    var par = new Paragraph();
                    foreach (var f in (field as ParagraphField).Items)
                    {

                        if (d.ContainsKey(f.Key))
                            par.Add(new Paragraph(d[f.Key].ToString()));
                        else if (!string.IsNullOrEmpty(f.Text))
                            par.Add(new Paragraph(f.Text));
                    }
                    yield return par;
                }
            }
        }

        protected IEnumerable<Fields> getMultifields(string key)
        {

            foreach (var x in DisplayFields)
            {
                if (x.Field == FieldType.MultiField && (x is ParagraphField) && x.Key==key)
                {
                    foreach (var field in (x as ParagraphField).Items)
                        yield return field;
                }
            }
        }

        protected bool isMultiFields(string key)
        {
            foreach (var x in DisplayFields)
            {
                if (x.Field == FieldType.MultiField && x.Key==key)
                    return true;
            }
            return false;
        }

        protected bool isTextPhrase(string key)
        {
            foreach (var x in DisplayFields)
            {
                if (x.Field== FieldType.Text)
                    return true;
            }
            return false;
        }

        protected bool isPhrase(string key)
        {
            foreach (var x in DisplayFields)
            {
                if (x.Key==key && x is PhraseField && x.Field== FieldType.PhraseField)
                    return true;
            }
            return false;
        }
        
        protected TextPhraseField getField(string key)
        {
            foreach (var x in DisplayFields)
            {
                if (x.Key == key && x is TextPhraseField)
                    return x as TextPhraseField;
            }
            return null;
        }
    }

    [Serializable]
    public class Page : Section
    {        
        [XmlAttribute("GroupBy")]
        public string GroupBy { get; set; }

        [XmlAttribute("SplitDoc")]        
        public bool SplitDoc { get; set; }

        [XmlAttribute("DistinctBy")]
        public string DistinctBy { get; set; }

        [XmlIgnore]
        public string CurrentGroup { get; set; }

        [XmlIgnore]
        public bool DocumentDone { get; set; }

        bool ContainCells()
        {
            foreach (var f in DisplayFields)
            {
                if (f is CellFields)
                    return true;
            }
            
            return false;
        }
        
        public override IEnumerable<Fields> Fields(IEnumerable<Dictionary<string, object>> data)
        {
            var newList = data.GroupByTagname(GroupBy);


            foreach (var x in newList)
            {
                
                    foreach (var y in x.FirstOrDefault().Keys)
                    {
                        yield return new Fields();// { Text = x.FirstOrDefault()[y].ToString(), Print = show(y) };
                    }
                }
            
            foreach (var x in newList)
            {
                foreach (var y in base.Fields(x))
                {
                    yield return y;
                }
            }

        }
        public override IEnumerable<Paragraph> Paragraphs(IEnumerable<Dictionary<string, object>> data)
        {
            IEnumerable<IGrouping<object, Dictionary<string, object>>> newList = null;
            if (!string.IsNullOrEmpty(DistinctBy))
                newList = data.GroupByTagname(GroupBy, DistinctBy);
            else 
                newList = data.GroupByTagname(GroupBy);

            foreach (var x in newList)
            {
                if (SplitDoc)
                {
                    CurrentGroup = x.Key.ToString();
                    DocumentDone = false;
                }
                Paragraph result = new Paragraph();
                List<Dictionary<string, object>> ifTable = new List<Dictionary<string, object>>();

                if (ContainCells())
                {
                    IEnumerable<Dictionary<string, object>> liste = from i in x.FirstOrDefault().Keys
                                                                    where show(i)
                                                                    select x.FirstOrDefault();
                    foreach (var p in this.getTable(liste.Distinct()))
                        result.Add(p);
                        //ifTable.Add(p);
                }
                else
                {
                    foreach (var y in this.DisplayFields)// x.FirstOrDefault().Keys)
                    {
                        //if (show(y.Key))
                        //{
                        if (y is PhraseField)//  isPhrase(y.Key))
                            result.Add(new Phrase(x.FirstOrDefault()[y.Key].ToString()));
                        else if (y is TextPhraseField)// isTextPhrase(y.Key))
                            result.Add(new Phrase(getField(y.Key).Text));
                        else if (isMultiFields(y.Key))
                        {
                            IEnumerable<Dictionary<string, object>> liste = from i in x.FirstOrDefault().Keys
                                                                            where show(i)
                                                                            select x.FirstOrDefault();
                            foreach (var tempP in getMultiParagraphs(y, liste.Distinct()))
                                result.Add(tempP);
                        }
                        else
                            result.Add(new Paragraph(x.FirstOrDefault()[y.Key].ToString()));
                        //}
                    }
                }
                foreach (var s in Sections)
                {
                    foreach (var p in s.Paragraphs(x))
                    {
                        if (result == null)
                            result = new Paragraph();
                        result.Add(p);
                    }
                }


                //if (ContainCells() && ifTable.Count() > 0)
                //{
                //    foreach (var p in this.getTable(ifTable))
                //        result.Add(p);
                //}
                if (SplitDoc)
                    DocumentDone = true;
                if (result != null)
                    yield return result;
            }
        }
    }

    [Serializable]
    public class Table : Section
    {
        public override IEnumerable<Paragraph> Paragraphs(IEnumerable<Dictionary<string, object>> data)
        {            
            foreach (var p in this.getTable(data))
                yield return p;
        }
    }

    [Serializable]
    public class Header : Section
    {

    }

    [Serializable]
    public static class SectionExtensions
    {
       public static IEnumerable<Paragraph> getTable (this Section section, IEnumerable<Dictionary<string, object>> data)
       {
           Paragraph result = new Paragraph();
           PdfPTable table = new PdfPTable(section.DisplayFields.Count());
           table.SplitLate = false;
           List<int> widths = new List<int>();
           foreach (var x in section.DisplayFields)
           {
               if (x is CellFields)
                   widths.Add((x as CellFields).Width);
               else
                   widths.Add(1);
           }

           table.WidthPercentage = 100;
           table.LockedWidth = false;
           table.SetWidths(widths.ToArray());
           foreach (var d in section.DisplayFields)
           {
               string tekst = d.Key;
               if (d is CellFields)
                   tekst = (d as CellFields).Header;

               var pc = new PdfPCell(new Phrase(tekst));
               pc.NoWrap = true;
               table.AddCell(pc);
           }

           foreach (var l in data)
           {
               foreach (var d in l.Keys)
               {
                   if (section.show(d))
                   {
                       var pc = new PdfPCell(new Phrase(l[d].ToString()));
                       pc.NoWrap = true;
                       table.AddCell(pc);
                   }
               }
           }
           result.Add(table);
           yield return result;
       }
    }
}
