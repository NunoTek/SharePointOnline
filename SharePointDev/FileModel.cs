using System;
using System.Collections.Generic;
using System.Text;

namespace SharePointDev
{
    public class FileModel
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public decimal Size { get; set; }
        public string WebUrl { get; set; }

        public DateTime LastModifiedDateTime { get; set; }
        public DateTime CreatedDateTime { get; set; }
    }
}
