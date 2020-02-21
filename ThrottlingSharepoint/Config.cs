using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThrottlingSharepoint
{
    class Config
    {
        public string crmUrl;
        public string crmUserName;
        public string sharepointUrl;
        public string sharepointUserEmailID;
        public string entityLogicalName;
        public string numberOfFoldersPerDL;
        public string isThereAnySubsite;
        public string subsiteUrl;
        public string newDLPrefix;
        public string newDLSuffixNumber;
        public string descriptionOfNewDL;
        public string deleteFilesAfterCopying;
        public string maxRecordsToBeMoved;
    }
}
