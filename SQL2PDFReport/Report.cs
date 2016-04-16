using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SQL2PDFReport
{

    public enum Orientation
    {
        Portrait, Landscape
    }

    [Serializable]
    public class Report
    {
        SqlConnection _con;
        Sections _sections;
        string _connectionString;
        string _SqlCommand;
        SqlCommand _com;
        List<Dictionary<string, object>> _dataSet;
        PdfPageHelper e;


        #region properties
        public string SqlCommand
        {
            get {
                if (!string.IsNullOrEmpty(_SqlCommand))
                    return _SqlCommand;
                else if (_com != null)
                    return _com.CommandText;
                else return null;
            }
            set { 
                _SqlCommand = value;
            }
        }
        public string ConnectionString
        {
            get {
                if (!string.IsNullOrEmpty(_connectionString))
                    return _connectionString;
                else if (_con != null)
                    return _con.ConnectionString;
                else return null;
            }
            set
            {
                _connectionString = value;
            }
        }
        public Orientation Orientation { get; set; }
        public Header Header { get; set; }
        public Footer Footer { get; set; }
        public Font DefaultFont { get; set; }
        [XmlElement(Type = typeof(Section))]
        [XmlElement(Type = typeof(Page))]
        [XmlElement(Type = typeof(Table))]
        [XmlElement(Type = typeof(Header))]
        public Sections Sections
        {
            get { return _sections; }
            set { _sections = value; }
        }

        [XmlIgnore]
        public SqlConnection Connection
        {
            get { return _con; }
            set { _con = value; }
        }

        [XmlIgnore]
        public SqlCommand Command
        {
            get { return _com; }
            set { _com = value; }
        }
        #endregion

        public Report()
        {
            _sections = new Sections();
           
        }

        #region Methods
        void LoadDataSet()
        {
            if (_dataSet == null && EnableSQL())
            {
                _dataSet = new List<Dictionary<string, object>>();
                _con.Open();
                SqlDataReader dr = _com.ExecuteReader();
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        Dictionary<string, object> row = new Dictionary<string, object>();
                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            string header = dr.GetName(i);
                            row.Add(header, dr[header]);
                        }
                        _dataSet.Add(row);
                    }
                }
                _con.Close();
            }
        }

        public IEnumerable<string> DataSetKeys()
        {
            LoadDataSet();
            foreach (var k in _dataSet[0].Keys)
                yield return k;
        }

        public IEnumerable<Dictionary<string, object>> DataSet()
        {
            LoadDataSet();
            foreach (var k in _dataSet)
                yield return k;
        }
                
        private Document getNewDocument(string filename)
        {
            var document = new Document(PageSize.A4, 25, 25, 25, 25);
            e = new PdfPageHelper();
            if (this.Header != null)            
                e.ImageHeader = Image.GetInstance( Header.ImagePath);
            if (this.Footer != null)
                e.ImageFooter = Image.GetInstance(Footer.ImagePath);
            if (this.Orientation== SQL2PDFReport.Orientation.Landscape)
                document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
            // Create a new PdfWriter object, specifying the output stream
            var output = new FileStream(filename, FileMode.Create);
            var writer = PdfWriter.GetInstance(document, output);
            writer.PageEvent = e;
            // Open the Document for writing
            return document;
        }

        private bool EnableSQL()
        {
            bool result = true;
            try
            {
                if (_con == null && string.IsNullOrEmpty(_connectionString) == false)
                    _con = new SqlConnection(_connectionString);
                if (_con == null)
                    return false;
                if (_com == null && string.IsNullOrEmpty(_SqlCommand) == false)
                    _com = new SqlCommand(_SqlCommand, _con);
                if (_com == null)
                    return false;
            }
            catch
            {
                result = false;
            }
            return result;
        }

        private iTextSharp.text.Font _font()
        {
            if (DefaultFont == null)
                return new Font().Parse();
            else return DefaultFont.Parse();
        }

        public void Generate(string filename)
        {
            LoadDataSet();
            if (_sections.Count() == 0)
                return;

            Document document = null;
            bool split = false;
            string path = "";
            if (_sections.First() is Page && (_sections.First() as Page).SplitDoc)
                split = true;
            if (split)
            {
                path = filename.Replace(".pdf", "");
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
            }
            else
            {
                document = getNewDocument(filename);
                document.Open();
            }
            //... Step 3: Add elements to the document! ...
            foreach (var s in _sections)
            {
                foreach (var p in s.Paragraphs(_dataSet, _font()))
                {
                    if ((s is Page) && (s as Page).SplitDoc && document == null)
                    {
                        document = getNewDocument(path + "\\" + (s as Page).CurrentGroup + ".pdf");
                        document.Open();
                    }

                    document.Add(p);
                    if (s is Page && (s as Page).SplitDoc==false)
                    document.NewPage();
                    if (s is Page && (s as Page).DocumentDone)
                    {
                        document.Close();
                        document = null;
                    }
                }
                
            }
            // Close the Document - this saves the document contents to the output stream
            if (!split)
                document.Close();
        }
        #endregion
    }

    public static class extensions
    {
        public static IEnumerable<IGrouping<object, Dictionary<string, object>>> GroupByTagname(this IEnumerable<Dictionary<string, object>> list, string field)
        {
            IEnumerable<IGrouping<object, Dictionary<string, object>>> result =
                from i in list
                group i by i[field];
            return result;                        
        }

        public static IEnumerable<IGrouping<object, Dictionary<string, object>>> GroupByTagname(this IEnumerable<Dictionary<string, object>> list, string field, string dist)
        {
            IEnumerable<IGrouping<object, Dictionary<string, object>>> result =
                from i in list.DistinctBy(u=> u[dist])
                group i by i[field];
            return result;
        }

        public static IEnumerable<T> DistinctBy<T, TKey>(this IEnumerable<T> items, Func<T, TKey> property)
        {
            return items.GroupBy(property).Select(x => x.First());
        }

        public static iTextSharp.text.Font Parse(this Font font)
        {
            if (string.IsNullOrEmpty(font.Name))
                font.Name = "Arial";
            if (font.Size == 0)
                font.Size = 12;

            var style = iTextSharp.text.Font.NORMAL;
            switch (font.Style)
            {
                case Style.bold:
                    style = iTextSharp.text.Font.BOLD;
                    break;
                case Style.italic:
                    style =iTextSharp.text.Font.ITALIC;
                    break;
                case Style.normal:
                    style = iTextSharp.text.Font.NORMAL;
                    break;
            }
            iTextSharp.text.Font _font = FontFactory.GetFont(font.Name, font.Size, style);
            //document.add(new Paragraph(font.Name, fontbold));
            string fontsfolder = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts);
            FontFactory.Register(fontsfolder+ @"\"+font.Name+".ttf", "my_font");
           // Font myBoldFont = FontFactory.getFont("my_bold_font");
            //BaseFont bf = _font.BaseFont;

            //BaseFont bfTimes = FontFactory.GetFont("my_font").BaseFont;
            //iTextSharp.text.Font times = new iTextSharp.text.Font(bf, font.Size, style, Color.BLACK);
            iTextSharp.text.Font times = new iTextSharp.text.Font(FontFactory.GetFont("my_font"));
            times.Size = font.Size;
            times.SetStyle(style);
            return times;
        }

        public static iTextSharp.text.Font Parse(this Font font, iTextSharp.text.Font relative)
        {
            var tempFont = new Font();
            if (string.IsNullOrEmpty(font.Name))
                tempFont.Name = relative.Familyname;
            else
                tempFont.Name = font.Name;

            if (font.Size == 0)
                tempFont.Size = Convert.ToInt32(relative.Size);
            else
                tempFont.Size = font.Size;

            if (font.Style != tempFont.Style)
                tempFont.Style = font.Style;

            return tempFont.Parse();
               
        }
    }
}
