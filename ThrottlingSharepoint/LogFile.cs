using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThrottlingSharepoint
{
    public class Failure
    {
        public EntityReference entityRef;
        public Guid SDLId;
        public string errorMessage;
    }

    public class Success
    {
        public Guid EntityId;
        public Guid SDLId;
        public string LogicalName;
        public string Name;
    }
}
