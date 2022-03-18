using AshfordSync.Entities;
using AshfordSync.Interfaces;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace AshfordSync.Service
{
    class ReadInventoryService : IReadInventoryService
    {
        private readonly ILogger<ReadInventoryService> _logger;

        public ReadInventoryService(ILogger<ReadInventoryService> logger)
        {
            _logger = logger;
        }
        public async Task ReadInventoryAsync(int supplierId, string fileName)
        {

            var inventorylists = new List<InventoryEntry>();

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);
            Console.WriteLine(jsonParamModel);
            Console.WriteLine(fileName);
            var excelFile = new FileInfo(".\\Inbox\\" + fileName);
            _logger.LogInformation("Opening Inventory File");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var epPackage = new ExcelPackage(excelFile))
            {


                var wscount = epPackage.Workbook.Worksheets.Count();
                Console.WriteLine("WS Count " + wscount);
                _logger.LogInformation("WS Count " + wscount);
                int initialRow = jsonParamModel.initialRow;
                int invItemNameColumn = jsonParamModel.invItemNameColumn;
                int invQuantityColumn = jsonParamModel.invQuantityColumn;
                Console.WriteLine("Param:" + initialRow);
                _logger.LogInformation("Param:" + initialRow);
                string itemName = "";
                int quantity = 0;


                for (int ws = 0; ws < wscount; ws++)
                {
                    var worksheet = epPackage.Workbook.Worksheets[ws];
                    int colCount = worksheet.Dimension.End.Column;
                    int rowCount = worksheet.Dimension.End.Row;
                    for (int row = initialRow; row <= rowCount; row = row + 1)
                    {
                        try
                        {
                            itemName = SkuTranslate(worksheet.Cells[row, invItemNameColumn].Value.ToString());
                            quantity = int.Parse(worksheet.Cells[row, invQuantityColumn].Value.ToString());
                            Console.WriteLine("Item: " + itemName + " Quantity: " + quantity);
                            _logger.LogInformation("Item: " + itemName + " Quantity: " + quantity);
                            InventoryEntry entry = new InventoryEntry();
                            entry.item_name = itemName;
                            entry.quantity = quantity;
                            if (quantity > 0)
                            {
                                inventorylists.Add(entry);
                            }

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                            _logger.LogError(ex.Message);
                            continue;
                        }
                    }
                }

                Console.WriteLine("ListCount:" + inventorylists.Count());
                _logger.LogInformation("ListCount:" + inventorylists.Count());

                var modelJson = JsonSerializer.Serialize(inventorylists, options);
                Console.WriteLine(modelJson);
                _logger.LogInformation(modelJson);
                Uri u = new Uri(jsonParamModel.tcouri + "/upload/inventory?SupplierID=" + supplierId);
                HttpClient httpClient = new HttpClient();
                HttpContent c = new StringContent(modelJson, System.Text.Encoding.UTF8, "application/json");
                var result = await httpClient.PostAsync(u, c);
                if (result.IsSuccessStatusCode)
                {
                    Console.WriteLine("Inventory Update accepted");
                    _logger.LogInformation("Inventory Update accepted");
                }
                else
                {
                    Console.WriteLine("Inventory Updated failed");
                    _logger.LogInformation("Inventory Updated failed");
                }

            }
        }

        public string SkuTranslate(string sku)
        {

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);

            return (sku.Replace(" ", jsonParamModel.spaceReplacement));
        }

        public async Task ReadRMAAsync(int supplierId, string fileName)
        {

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);
            Console.WriteLine(jsonParamModel);
            _logger.LogInformation(jsonParamModel.ToString());
            Console.WriteLine(fileName);
            _logger.LogInformation(fileName);
            var excelFile = new FileInfo(".\\Inbox\\" + fileName);
            var cultureInfo = new CultureInfo("en-US");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var epPackage = new ExcelPackage(excelFile))
            {
                var wscount = epPackage.Workbook.Worksheets.Count();
                Console.WriteLine("WS Count " + wscount);
                _logger.LogInformation("WS Count " + wscount);
                int initialRow = jsonParamModel.initialRow;
                Console.WriteLine("Param:" + initialRow);
                _logger.LogInformation("Param:" + initialRow);
                
                var worksheet = epPackage.Workbook.Worksheets[0]; 
                int colCount = worksheet.Dimension.End.Column;  
                int rowCount = worksheet.Dimension.End.Row;   
                for (int row = initialRow; row <= rowCount; row = row + 1)
                {
                    try
                    {
                        RMA rma = new RMA();
                        Line rmaEntry = new Line();

                        if (jsonParamModel.rmaHeaderIdColumn > 0)
                        {
                            rma.rmaHeaderId = int.Parse(worksheet.Cells[row, jsonParamModel.rmaHeaderIdColumn].Value.ToString());
                        }
                        else
                        {
                            rma.rmaHeaderId = int.Parse(worksheet.Cells[row, jsonParamModel.rmaOrderNumberColumn].Value.ToString());
                        }

                        rma.rmaOrderNumber = worksheet.Cells[row, jsonParamModel.rmaOrderNumberColumn].Value.ToString();

                        rma.rmaDate = DateTime.Parse(worksheet.Cells[row, jsonParamModel.rmaDateColumn].Value.ToString(), cultureInfo);

                        var worksheetDtl = epPackage.Workbook.Worksheets[1]; 
                        int colCountDtl = worksheetDtl.Dimension.End.Column;  
                        int rowCountDtl = worksheetDtl.Dimension.End.Row;     
                        for (int rowDtl = initialRow; rowDtl <= rowCountDtl; rowDtl = rowDtl + 1)
                        {
                            try
                            {

                                if (worksheetDtl.Cells[row, jsonParamModel.rmaSourceOrderNumberColumn].Value.ToString().Equals(rma.rmaOrderNumber))
                                {

                                    rmaEntry.sourceOrderNumber = worksheetDtl.Cells[row, jsonParamModel.rmaSourceOrderNumberColumn].Value.ToString();

                                    rmaEntry.itemNumber = SkuTranslate(worksheetDtl.Cells[row, jsonParamModel.rmaItemNumberColumn].Value.ToString());

                                    OrderEntry entry = GetOrderQuantity(rmaEntry.itemNumber, rmaEntry.sourceOrderNumber);

                                    rmaEntry.quantity = int.Parse(worksheetDtl.Cells[row, jsonParamModel.rmaQuantityColumn].Value.ToString());


                                    if (jsonParamModel.rmaLineIdColumn > 0)
                                        rmaEntry.rmaLineId = int.Parse(worksheetDtl.Cells[row, jsonParamModel.rmaLineIdColumn].Value.ToString());

                                    if (jsonParamModel.rmaSourceLineNumberColumn > 0)
                                    {
                                        rmaEntry.sourceLineNumber = int.Parse(worksheetDtl.Cells[row, jsonParamModel.rmaSourceLineNumberColumn].Value.ToString());
                                    }
                                    else
                                    {
                                        rmaEntry.sourceLineNumber = entry.lineNumber;
                                    }

                                    if (jsonParamModel.rmaReasonColumn > 0)
                                        rmaEntry.reason = worksheetDtl.Cells[row, jsonParamModel.rmaReasonColumn].Value.ToString();

                                    if (jsonParamModel.rmaRestockCodeColumn > 0)
                                        rmaEntry.restockCode = worksheetDtl.Cells[row, jsonParamModel.rmaRestockCodeColumn].Value.ToString();


                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError(ex.Message);
                                Console.WriteLine(ex);
                                continue;
                            }
                        }


                        List<Line> rmaList = new List<Line>();
                        rmaList.Add(rmaEntry);
                        rma.lines = rmaList;

                        var modelJson = JsonSerializer.Serialize(rma, options);
                        Console.WriteLine("modelJson: " + modelJson);
                        _logger.LogInformation("modelJson: " + modelJson);
                        Uri u = new Uri(jsonParamModel.tcouri + "/upload/returns?SupplierID=" + supplierId);
                        HttpClient httpClient = new HttpClient();
                        HttpContent c = new StringContent(modelJson, System.Text.Encoding.UTF8, "application/json");
                        var result = await httpClient.PostAsync(u, c);
                        if (result.IsSuccessStatusCode)
                        {
                            Console.WriteLine("Return Update accepted");
                            _logger.LogInformation("Return Update accepted:" + modelJson);
                        }
                        else
                        {
                            Console.WriteLine("Return Update failed");
                            _logger.LogInformation("Return Update accepted:" + modelJson);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        _logger.LogError(ex.ToString());
                        continue;
                    }
                }
                
            }
        }

        public OrderEntry GetOrderQuantity(string sku, string orderNumber)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);

            Console.WriteLine("jsonParamModel.connectionString:" + jsonParamModel.connectionString);
            _logger.LogInformation("jsonParamModel.connectionString:" + jsonParamModel.connectionString);

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

            builder.ConnectionString = jsonParamModel.connectionString;

            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                connection.Open();
                string SQLstr = String.Format("SELECT a.ItemLookupCode,a.QtyOrdered,a.SimpleProdLineNo FROM InvictaAUX.dbo.eCommerceOrderEntry a WHERE a.OrderNumber = '{0}' and a.ItemLookUpCode = '{1}' and a.SimpleProdLineNo>0 ", orderNumber, sku);
                using (SqlCommand cmd = new SqlCommand(SQLstr, connection))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            return new OrderEntry()
                            {
                                itemNumber = reader.GetString(0),
                                orderedQuantity = reader.GetInt16(1),
                                lineNumber = reader.GetInt16(2)
                            };

                        }
                    }
                }

            }
            return new OrderEntry();

        }

        public async Task ReadShipConfirmAsync(int supplierId, string fileName)
        {

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);
            Console.WriteLine(jsonParamModel);
            _logger.LogInformation(jsonParamModel.ToString());
            Console.WriteLine(fileName);
            _logger.LogInformation(fileName);
            var excelFile = new FileInfo(".\\Inbox\\" + fileName);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var cultureInfo = new CultureInfo("en-US");
            using (var epPackage = new ExcelPackage(excelFile))
            {
                var wscount = epPackage.Workbook.Worksheets.Count();
                Console.WriteLine("WS Count " + wscount);
                _logger.LogInformation("WS Count " + wscount);
                int initialRow = jsonParamModel.initialRow;
                Console.WriteLine("Param:" + initialRow);
                _logger.LogInformation("Param:" + initialRow);
                for (int ws = 0; ws < wscount; ws++)
                {
                    var worksheet = epPackage.Workbook.Worksheets[ws];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count
                    for (int row = initialRow; row <= rowCount; row = row + 1)
                    {
                        try
                        {
                            ShipConfirm sc = new ShipConfirm();
                            Detail det = new Detail();
                            sc.orderNumber = worksheet.Cells[row, jsonParamModel.scOrderNumberColumn].Value.ToString();
                            sc.customerNumber = worksheet.Cells[row, jsonParamModel.scCustomerNumberColumn].Value.ToString();
                            sc.orderDate = DateTime.Parse(worksheet.Cells[row, jsonParamModel.scOrderDateColumn].Value.ToString(), cultureInfo);


                            det.itemNumber = SkuTranslate(worksheet.Cells[row, jsonParamModel.scItemNumberColumn].Value.ToString());

                            OrderEntry entry = GetOrderQuantity(det.itemNumber, sc.orderNumber);

                            if (jsonParamModel.scLineNumberColumn > 0)
                            {
                                det.lineNumber = int.Parse(worksheet.Cells[row, jsonParamModel.scLineNumberColumn].Value.ToString());
                            }
                            else
                            {
                                det.lineNumber = entry.lineNumber;
                            }

                            if (jsonParamModel.scOrderedQuantityColumn > 0)
                            {
                                det.orderedQuantity = int.Parse(worksheet.Cells[row, jsonParamModel.scOrderedQuantityColumn].Value.ToString());
                            }
                            else
                            {
                                det.orderedQuantity = entry.orderedQuantity;
                            }

                            det.shippedQuantity = int.Parse(worksheet.Cells[row, jsonParamModel.scShippedQuantityColumn].Value.ToString());

                            if (jsonParamModel.scCancelledQuantityColumn > 0)
                                det.canceledQuantity = int.Parse(worksheet.Cells[row, jsonParamModel.scCancelledQuantityColumn].Value.ToString());

                            if (jsonParamModel.scShippedDateColumn > 0)
                            {
                                det.shippedDate = DateTime.Parse(worksheet.Cells[row, jsonParamModel.scShippedDateColumn].Value.ToString(), cultureInfo);
                            }
                            else
                            {
                                det.shippedDate = DateTime.Now;
                            }

                            if (jsonParamModel.scCarrierColumn > 0)
                            {
                                det.carrier = worksheet.Cells[row, jsonParamModel.scCarrierColumn].Value.ToString();
                            }
                            else
                            {
                                det.carrier = jsonParamModel.defaultcarrier;
                            }

                            det.trackingNumber = worksheet.Cells[row, jsonParamModel.scTrackingNumberColumn].Value.ToString();

                            if (jsonParamModel.scPrePaidRetunLabelUsedColumn > 0)
                            {
                                det.prePaidReturnLabelUsed = (worksheet.Cells[row, jsonParamModel.scPrePaidRetunLabelUsedColumn].Value.ToString().Equals("Y")) ? true : false;
                            }
                            else
                            {
                                det.prePaidReturnLabelUsed = false;
                            }

                            if (jsonParamModel.scPrePaidReturnLabelCostColumn > 0)
                                det.prePaidReturnLabelCost = Decimal.Parse(worksheet.Cells[row, jsonParamModel.scPrePaidReturnLabelCostColumn].Value.ToString());

                            List<Detail> details = new List<Detail>();
                            details.Add(det);
                            sc.details = details;

                            var modelJson = JsonSerializer.Serialize(sc, options);
                            Console.WriteLine("modelJson: " + modelJson);
                            _logger.LogInformation("modelJson: " + modelJson);
                            Uri u = new Uri(jsonParamModel.tcouri + "/upload/shipping?SupplierID=" + supplierId);
                            HttpClient httpClient = new HttpClient();
                            HttpContent c = new StringContent(modelJson, System.Text.Encoding.UTF8, "application/json");
                            var result = await httpClient.PostAsync(u, c);
                            if (result.IsSuccessStatusCode)
                            {
                                Console.WriteLine("Shipping Update accepted");
                                _logger.LogInformation("Shipping Update accepted:" + modelJson);
                            }
                            else
                            {
                                Console.WriteLine("Shipping Updated failed");
                                _logger.LogInformation("Shipping Update failed:" + modelJson);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                            _logger.LogInformation(ex.ToString());
                            continue;
                        }
                    }
                }
            }
        }

        public async Task ReadJsonShipConfirmAsync(int supplierId, string fileName)
        {

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);
            Console.WriteLine(jsonParamModel);
            _logger.LogInformation(jsonParamModel.ToString());
            Console.WriteLine(fileName);
            _logger.LogInformation(fileName);
            var excelFile = new FileInfo(".\\Inbox\\" + fileName);
            var cultureInfo = new CultureInfo("en-US");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var epPackage = new ExcelPackage(excelFile))
            {
                var wscount = epPackage.Workbook.Worksheets.Count();
                Console.WriteLine("WS Count " + wscount);
                _logger.LogInformation("WS Count " + wscount);
                int initialRow = jsonParamModel.initialRow;
                Console.WriteLine("Param:" + initialRow);
                _logger.LogInformation("Param:" + initialRow);
                for (int ws = 0; ws < wscount; ws++)
                {
                    var worksheet = epPackage.Workbook.Worksheets[ws];
                    int colCount = worksheet.Dimension.End.Column;  
                    int rowCount = worksheet.Dimension.End.Row;     
                    for (int row = initialRow; row <= rowCount; row = row + 1)
                    {
                        try
                        {
                            var jsonentry = worksheet.Cells[row, 2].Value.ToString();
                            Uri u = new Uri(jsonParamModel.tcouri + "/upload/shipping?SupplierID=" + supplierId);
                            HttpClient httpClient = new HttpClient();
                            HttpContent c = new StringContent(jsonentry, Encoding.UTF8, "application/json");
                            var result = await httpClient.PostAsync(u, c);
                            if (result.IsSuccessStatusCode)
                            {
                                Console.WriteLine("Shipping Update accepted");
                                _logger.LogInformation("Shipping Update accepted");
                            }
                            else
                            {
                                Console.WriteLine("Shipping Updated failed");
                                _logger.LogInformation("Shipping Updated failed");
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }
        }
    }
}
