using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Code_Report
{
    [Serializable]
    class Codes
    {
        private string _Type = "";
        public string Number { get; set; } = "";
        public string Link { get; set; } = "";
        public string WebType { get; set; } = "";
        public string ProductCategory { get; set; } = "";
        public string Description { get; set; } = "";
        public string ProductsListed { get; set; } = "";
        public string LatestCode { get; set; } = "";
        public string IssueDate { get; set; } = "";
        public string ExpirationDate { get; set; } = "";
        public Codes()
        {

        }
        ~Codes()
        {
            this.Number = string.Empty;
            this.Link = string.Empty;
            this.WebType = string.Empty;
            this.ProductCategory = string.Empty;
            this.Description = string.Empty;
            this.ProductsListed = string.Empty;
            this.LatestCode = string.Empty;
            this.IssueDate = string.Empty;
            this.ExpirationDate = string.Empty;
            GC.Collect();
        }

        public void setType(string type)
        {
            _Type = type;
        }
        public string getType()
        {
            return _Type;
        }
    }
}
