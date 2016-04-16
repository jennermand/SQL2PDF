using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace SQL2PDFReport
{
    
    public class ReportGenerator
    {
        Report _report;
        
        public Report Report
        {
            get { return _report; }
            set { _report = value; }
        }
               
        public ReportGenerator(Report report)
        {
            _report = report;
            
        }

        public void AddSection(Section section)
        {
            _report.Sections.Add(section);
        }

        public void Create(string filename)
        {
            _report.Generate(filename, _report.Sections);
        }

        public void SaveProjekt(string filename)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Report));
            using (TextWriter writer = new StreamWriter(filename))
            {
                serializer.Serialize(writer, _report);
            }
        }
    }   
}
