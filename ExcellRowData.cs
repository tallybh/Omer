using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace c_sharp_quickstart
{
    public class ExcellRowData
    {
        public string TierId { get; set; }

        public string ProgramId { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }

        public string License_Number { get; set; }
        public DateTime ExpirationDate { get; set; }

        public string PhotoPath { get; set; }

    }
}
