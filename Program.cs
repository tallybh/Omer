
using c_sharp_quickstart;
using ClosedXML.Excel;

var ids = File.ReadAllText("Data/Ids.txt").Split();
var programId = ids[1];
var tierId = ids[3];

using (var workbook = new XLWorkbook("Data/members.xlsx"))
{
    var worksheet = workbook.Worksheet(1);
    var rows = worksheet.RangeUsed().RowsUsed();
    int i = 0;
    foreach (var row in rows)
    {
        i++;
        if (i ==1)
        {
            continue;
        }
        try
        {
            ExcellRowData data = new ExcellRowData();
            data.TierId = tierId;
            data.ProgramId = programId;

            var cells = row.Cells().ToList();
            data.Name = cells[0].Value.ToString();
            data.Id = cells[1].Value.ToString();
            data.License_Number = cells[2].Value.ToString();
            var times = cells[3].Value.ToString().Split(' ')[0].Split('/');
            data.ExpirationDate = new DateTime(Convert.ToInt32(times[2]), Convert.ToInt32(times[1]), Convert.ToInt32(times[0]), 0, 0, 0, DateTimeKind.Utc);
            data.PhotoPath = cells[4].Value.ToString();

            Generator.Generate(data);
        }

        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
      
    }
}
    // Creates channel for connecting to PassKit server
    //var channel = GrpcConnection.GrpcConnection.ConnectWithPassKitServer();

// Loyalty Quickstart
//QuickstartLoyalty.Membership buildLoyalty = new QuickstartLoyalty.Membership();
//buildLoyalty.Quickstart(channel);

// Coupons Quickstart
//QuickstartCoupons.Coupons buildCoupons = new QuickstartCoupons.Coupons();
//buildCoupons.Quickstart(channel);

// Flight Quickstart
//QuickstartFlightTickets.FlightTickets buildFlights = new QuickstartFlightTickets.FlightTickets();
//buildFlights.QuickStart(channel);

// Event Tickets Quickstart
//QuickstartEventickets.EventTicket buildEventTickets = new QuickstartEventickets.EventTicket();
//buildEventTickets.QuickStart(channel);
