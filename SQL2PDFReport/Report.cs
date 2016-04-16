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
    [Serializable]
    public class Report
    {
        SqlConnection _con;
        Sections _sections;
        string _connectionString;
        string _SqlCommand;
        SqlCommand _com;
        List<Dictionary<string, object>> _dataSet;

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
                _com = new SqlCommand(value);
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
                _con = new SqlConnection(value);
            }
        }


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
            if (_dataSet == null)
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
            var document = new Document(PageSize.A4, 50, 50, 25, 25);

            // Create a new PdfWriter object, specifying the output stream
            var output = new FileStream(filename, FileMode.Create);
            var writer = PdfWriter.GetInstance(document, output);

            // Open the Document for writing
            return document;
        }

        public void Generate(string filename, IEnumerable<Section> sections)
        {
            if (sections.Count() == 0)
                return;

            Document document = null;
            bool split = false;
            string path = "";
            if (sections.First() is Page && (sections.First() as Page).SplitDoc)
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
            foreach (var s in sections)
            {
                foreach (var p in s.Paragraphs(_dataSet))
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
    }
}
