using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace c_sharp_quickstart
{
    public static class Generator
    {
        public static void Generate(ExcellRowData data)
        {
            var channel = GrpcConnection.GrpcConnection.ConnectWithPassKitServer();
            QuickstartLoyalty.Membership buildLoyalty = new QuickstartLoyalty.Membership(data);
            buildLoyalty.Quickstart(channel);
        }
    }
}
