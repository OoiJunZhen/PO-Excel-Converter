using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using OfficeOpenXml;
using PO_Excel.Model;

namespace PO_Excel
{
    public class ExcelHelper : IDisposable
    {
        private readonly ExcelPackage? package;
        private readonly ExcelWorksheet MODetailsSheet;
        private readonly ExcelWorksheet PRDetailsSheet;
        private readonly ExcelWorksheet PODetailsSheet;
        private bool disposed = false;
        private readonly Dictionary<string, int> moColumnIndices;
        private readonly Dictionary<string, int> prColumnIndices;
        private readonly Dictionary<string, int> poColumnIndices;

        private static readonly HashSet<int> previouslyReturnedRows_loadData = ([]);

        public ExcelHelper(string FilePath1)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            try
            {
                package = new ExcelPackage(new FileInfo(FilePath1));
                MODetailsSheet = package.Workbook.Worksheets["MO Details"];
                PRDetailsSheet = package.Workbook.Worksheets["PR Details"];
                PODetailsSheet = package.Workbook.Worksheets["PO Details"];
                moColumnIndices = GetColumnIndices(MODetailsSheet);
                prColumnIndices = GetColumnIndices(PRDetailsSheet);
                poColumnIndices = GetColumnIndices(PODetailsSheet);
            }
            catch (Exception ex)
            {
                Dispose();
                throw new InvalidOperationException("Worksheet cannot be found.", ex);
            }
        }

        private static Dictionary<string, int> GetColumnIndices(ExcelWorksheet worksheet)
        {
            var columns = new Dictionary<string, int>();
            if (worksheet == null)
                throw new Exception("Worksheet cannot be null");

            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                string header = worksheet.Cells[1, col].Text.Trim();
                if (!string.IsNullOrEmpty(header))
                {
                    columns[header] = col;
                }
            }

            return columns;
        }

        public List<DataList> LoadColumns(string projectCodeFlter, BackgroundWorker? worker, CancellationToken cancellationToken)
        {
            var dataList = new List<DataList>();
            AddDataToList(projectCodeFlter, dataList, worker, cancellationToken);
            previouslyReturnedRows_loadData.Clear();
            return dataList;
        }

        private void AddDataToList(string projectCodeFilter, List<DataList> dataList, BackgroundWorker? worker, CancellationToken cancellationToken)
        {
            var rowErrors = new List<string>();
            int totalRows = CountNonNullProjectCodeRows(projectCodeFilter);
            int rowsProcessed = 0;

            for (int row = 2; row <= PRDetailsSheet.Dimension.Rows; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    string PRNo = PRDetailsSheet.Cells[row, prColumnIndices["PR Number"]].Text;
                    string PRProjectCode = PRDetailsSheet.Cells[row, prColumnIndices["Project Code"]].Text;
                    string PRApprovedOn = PRDetailsSheet.Cells[row, prColumnIndices["PR Approved On"]].Text;
                    string PRMaterialCode = PRDetailsSheet.Cells[row, prColumnIndices["Material Code"]].Text;
                    string PRQty = PRDetailsSheet.Cells[row, prColumnIndices["PR Quantity"]].Text;

                    if (string.IsNullOrEmpty(PRProjectCode) || !PRProjectCode.Contains(projectCodeFilter))
                    {
                        continue;
                    }
                    var matchingPOData = FindMatchingPOData(PRMaterialCode, PRProjectCode, PRQty);

                    dataList.Add(new DataList
                    {
                        PRNo = PRNo,
                        PRProjectCode = PRProjectCode,
                        PRApprovedOn = PRApprovedOn,
                        PRMaterialCode = PRMaterialCode,
                        PRQty = PRQty,
                        POProjectCode = matchingPOData?.POProjectCode ?? string.Empty,
                        PONo = matchingPOData?.PONo ?? string.Empty,
                        Supplier = matchingPOData?.Supplier ?? string.Empty,
                        POMaterialCode = matchingPOData?.POMaterialCode ?? string.Empty,
                        POApprovedOn = matchingPOData?.POApprovedOn ?? string.Empty,
                        POQty = matchingPOData?.POQty ?? string.Empty,
                        ReceivedQty = matchingPOData?.ReceivedQty ?? string.Empty,
                        ETA = matchingPOData?.ETA ?? string.Empty,
                    });

                    rowsProcessed++;

                    ReportProgress_LoadData(worker, rowsProcessed, totalRows);
                }
                catch (Exception ex)
                {
                    rowErrors.Add($"Error processing data {row}: {ex.Message}");
                }
            }
            if (rowErrors.Count > 0)
            {
                throw new Exception(string.Join(Environment.NewLine, rowErrors));
            }
        }

        private (string POProjectCode, string PONo, string Supplier, string POMaterialCode, string POApprovedOn, string POQty, string ReceivedQty, string ETA)? FindMatchingPOData(string PRMaterialCode, string PRProjectCode, string PRQty)
        {
            if (string.IsNullOrEmpty(PRMaterialCode))
                return null;

            for (int row = 2; row <= PODetailsSheet.Dimension.Rows; row++)
            {
                string POMaterialCode = PODetailsSheet.Cells[row, poColumnIndices["Material Code"]].Text;
                string POProjectCode = PODetailsSheet.Cells[row, poColumnIndices["Project Code"]].Text;
                string POQty = PODetailsSheet.Cells[row, poColumnIndices["PO Quantity"]].Text;
                string Supplier = PODetailsSheet.Cells[row, poColumnIndices["Supplier"]].Text;

                if (POMaterialCode.Equals(PRMaterialCode, StringComparison.OrdinalIgnoreCase) && POProjectCode.Equals(PRProjectCode, StringComparison.OrdinalIgnoreCase) &&
                    POQty.Equals(PRQty, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(PRProjectCode))
                {

                    if (previouslyReturnedRows_loadData.Contains(row))
                    {
                        continue;
                    }

                    string PONo = PODetailsSheet.Cells[row, poColumnIndices["PO Number"]].Text;
                    string POApprovedOn = PODetailsSheet.Cells[row, poColumnIndices["PO Approved On"]].Text;
                    string ReceivedQty = PODetailsSheet.Cells[row, poColumnIndices["Received Quantity"]].Text;
                    string ETA = PODetailsSheet.Cells[row, poColumnIndices["Planned Arrival Date"]].Text;

                    previouslyReturnedRows_loadData.Add(row);
                    return (POProjectCode, PONo, Supplier, POMaterialCode, POApprovedOn, POQty, ReceivedQty, ETA);
                }
            }
            return null;
        }

        public void SaveToFile(List<DataList> dataList, string saveFilePath, BackgroundWorker? worker, string projectCodeFilter, CancellationToken cancellationToken)
        {
            List<DataListMO> dataListMO = [];
            List<DataListPO> dataListPO = [];
            using var package = new ExcelPackage(new FileInfo(saveFilePath));

            var worksheetFab = package.Workbook.Worksheets["Fab_Parts"]
                ?? throw new Exception($"Worksheet 'Fab_Parts' not found in the provided Excel file: {saveFilePath}");
            var worksheetStd = package.Workbook.Worksheets["Std_Parts"]
                ?? throw new Exception($"Worksheet 'Std_Parts' not found in the provided Excel file: {saveFilePath}");
            var worksheetEE = package.Workbook.Worksheets["EE_Parts"]
                ?? throw new Exception($"Worksheet 'EE_Parts' not found in the provided Excel file: {saveFilePath}");

            int totalRows = dataList.Count;
            int totalRowsMO = CountNonNullMOProjectCodeRows(projectCodeFilter);
            int totalRowsPO = CountNonNullPOProjectCodeRows(projectCodeFilter);
            int rowsProcessed = 0;
            int rowsProcessedMO = 0;
            int rowsProcessedPO = 0;

            var batchFab = new List<(int, DataList)>();
            var batchStd = new List<(int, DataList)>();
            var batchEE = new List<(int, DataList)>();
            var batchMOFab = new List<(int, DataListMO)>();
            var batchMOStd = new List<(int, DataListMO)>();
            var batchMOEE = new List<(int, DataListMO)>();
            var batchPOFab = new List<(int, DataListPO)>();
            var batchPOStd = new List<(int, DataListPO)>();
            var batchPOEE = new List<(int, DataListPO)>();

            try
            {
                foreach (var data in dataList)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    List<int> matchingRowsFab = FindRowIndices(worksheetFab, data.PRMaterialCode, data.POMaterialCode, data.PRProjectCode, data.POProjectCode);
                    List<int> matchingRowsStd = FindRowIndices(worksheetStd, data.PRMaterialCode, data.POMaterialCode, data.PRProjectCode, data.POProjectCode);
                    List<int> matchingRowsEE = FindRowIndices(worksheetEE, data.PRMaterialCode, data.POMaterialCode, data.PRProjectCode, data.POProjectCode);

                    batchFab.AddRange(matchingRowsFab.Select(row => (row, data)));
                    batchStd.AddRange(matchingRowsStd.Select(row => (row, data)));
                    batchEE.AddRange(matchingRowsEE.Select(row => (row, data)));

                    rowsProcessed++;
                    ReportProgress_InsertData(worker, rowsProcessed, totalRows);
                }

                InsertDataBatch(worksheetFab, batchFab, saveFilePath);
                InsertDataBatch(worksheetStd, batchStd, saveFilePath);
                InsertDataBatch(worksheetEE, batchEE, saveFilePath);

                for (int row = 2; row <= PODetailsSheet.Dimension.Rows; row++)
                {
                    string POProjectCode = PODetailsSheet.Cells[row, poColumnIndices["Project Code"]].Text;
                    string POMaterialCode = PODetailsSheet.Cells[row, poColumnIndices["Material Code"]].Text;
                    string POQty = PODetailsSheet.Cells[row, poColumnIndices["PO Quantity"]].Text;
                    string PONo = PODetailsSheet.Cells[row, poColumnIndices["PO Number"]].Text;
                    string POApprovedOn = PODetailsSheet.Cells[row, poColumnIndices["PO Approved On"]].Text;
                    string ReceivedQty = PODetailsSheet.Cells[row, poColumnIndices["Received Quantity"]].Text;

                    if (!POProjectCode.Contains(projectCodeFilter)) continue;

                    dataListPO.Add(new DataListPO
                    {
                        POProjectCode = POProjectCode,
                        POMaterialCode = POMaterialCode,
                        POQty = POQty,
                        PONo = PONo,
                        POApprovedOn = POApprovedOn,
                        ReceivedQty = ReceivedQty
                    });
                }

                foreach (var dataPO in dataListPO)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (!dataPO.POProjectCode.Contains(projectCodeFilter) || string.IsNullOrEmpty(dataPO.PONo)) continue;

                    List<int> matchingPORowsFab = POFindRowIndices(worksheetFab, dataPO.POProjectCode, dataPO.POMaterialCode, dataPO.POQty);
                    List<int> matchingPORowsStd = POFindRowIndices(worksheetStd, dataPO.POProjectCode, dataPO.POMaterialCode, dataPO.POQty);
                    List<int> matchingPORowsEE = POFindRowIndices(worksheetEE, dataPO.POProjectCode, dataPO.POMaterialCode, dataPO.POQty);

                    batchPOFab.AddRange(matchingPORowsFab.Select(row => (row, dataPO)));
                    batchPOStd.AddRange(matchingPORowsStd.Select(row => (row, dataPO)));
                    batchPOEE.AddRange(matchingPORowsEE.Select(row => (row, dataPO)));

                    rowsProcessedPO++;
                    ReportProgress_InsertPOData(worker, rowsProcessedPO, totalRowsPO);
                }

                InsertPODataBatch(worksheetFab, batchPOFab, saveFilePath);
                InsertPODataBatch(worksheetStd, batchPOStd, saveFilePath);
                InsertPODataBatch(worksheetEE, batchPOEE, saveFilePath);

                for (int row = 2; row <= MODetailsSheet.Dimension.Rows; row++)
                {
                    string MOProjectCode = MODetailsSheet.Cells[row, moColumnIndices["Project Code"]].Text;
                    string MOMaterialCode = MODetailsSheet.Cells[row, moColumnIndices["Material Code"]].Text;
                    string MOQty = MODetailsSheet.Cells[row, moColumnIndices["BOM Quantity"]].Text;
                    string CollectedQuantity = MODetailsSheet.Cells[row, moColumnIndices["Collected Quantity"]].Text;

                    if (!MOProjectCode.Contains(projectCodeFilter)) continue;

                    dataListMO.Add(new DataListMO
                    {
                        MOProjectCode = MOProjectCode,
                        MOMaterialCode = MOMaterialCode,
                        MOQty = MOQty,
                        CollectedQty = CollectedQuantity
                    });
                }

                foreach (var dataMO in dataListMO)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (!dataMO.MOProjectCode.Contains(projectCodeFilter)) continue;

                    List<int> matchingMORowsFab = MOFindRowIndices(worksheetFab, dataMO.MOProjectCode, dataMO.MOMaterialCode, dataMO.CollectedQty);
                    List<int> matchingMORowsStd = MOFindRowIndices(worksheetStd, dataMO.MOProjectCode, dataMO.MOMaterialCode, dataMO.CollectedQty);
                    List<int> matchingMORowsEE = MOFindRowIndices(worksheetEE, dataMO.MOProjectCode, dataMO.MOMaterialCode, dataMO.CollectedQty);

                    batchMOFab.AddRange(matchingMORowsFab.Select(row => (row, dataMO)));
                    batchMOStd.AddRange(matchingMORowsStd.Select(row => (row, dataMO)));
                    batchMOEE.AddRange(matchingMORowsEE.Select(row => (row, dataMO)));

                    rowsProcessedMO++;
                    ReportProgress_InsertStatusData(worker, rowsProcessedMO, totalRowsMO);
                }

                InsertMODataBatch(worksheetFab, batchMOFab, saveFilePath);
                InsertMODataBatch(worksheetStd, batchMOStd, saveFilePath);
                InsertMODataBatch(worksheetEE, batchMOEE, saveFilePath);

                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    using var memoryStream = new MemoryStream();
                    package.SaveAs(memoryStream);
                    File.WriteAllBytes(saveFilePath, memoryStream.ToArray());
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Error saving the Excel file. Please check if the file is open or not.", ex);
                }
            }
            finally
            {
                batchFab.Clear();
                batchStd.Clear();
                batchEE.Clear();
                batchMOFab.Clear();
                batchMOStd.Clear();
                batchMOEE.Clear();
                batchPOFab.Clear();
                batchPOStd.Clear();
                batchPOEE.Clear();
                dataListMO.Clear();
            }
        }

        private static void ReportProgress_LoadData(BackgroundWorker? worker, int rowsProcessed, int totalRows)
        {
            if (worker != null && worker.WorkerReportsProgress)
            {
                int percentage = rowsProcessed * 100 / totalRows;
                worker.ReportProgress(percentage, "Loading data... Please wait.");
            }
        }
        private static void ReportProgress_InsertData(BackgroundWorker? worker, int rowsProcessed, int totalRows)
        {
            if (worker != null && worker.WorkerReportsProgress)
            {
                int basePercentage = rowsProcessed * 100 / totalRows;
                int scaledPercentage = 1 + (basePercentage * 39 / 100);
                worker.ReportProgress(scaledPercentage, "Inserting data... Please wait.");
            }
        }

        private static void ReportProgress_InsertPOData(BackgroundWorker? worker, int rowsProcessed, int totalRows)
        {
            if (worker != null && worker.WorkerReportsProgress)
            {
                int basePercentage = rowsProcessed * 100 / totalRows;
                int scaledPercentage = 41 + (basePercentage * 19 / 100);
                worker.ReportProgress(scaledPercentage, "Inserting data... Please wait.");
            }
        }

        private static void ReportProgress_InsertStatusData(BackgroundWorker? worker, int rowsProcessed, int totalRows)
        {
            if (worker != null && worker.WorkerReportsProgress)
            {
                int basePercentage = rowsProcessed * 100 / totalRows;
                int scaledPercentage = 61 + (basePercentage * 39 / 100);
                worker.ReportProgress(scaledPercentage, "Inserting data... Please wait.");
            }
        }



        private int CountNonNullProjectCodeRows(string projectCodeFilter)
        {
            int count = 0;
            for (int row = 2; row <= PRDetailsSheet.Dimension.Rows; row++)
            {
                string PRProjectCode = PRDetailsSheet.Cells[row, prColumnIndices["Project Code"]].Text;
                if (PRProjectCode.Contains(projectCodeFilter))
                {
                    count++;
                }
            }
            return count;
        }

        private int CountNonNullMOProjectCodeRows(string projectCodeFilter)
        {
            int count = 0;
            for (int row = 2; row <= MODetailsSheet.Dimension.Rows; row++)
            {
                string MOProjectCode = MODetailsSheet.Cells[row, moColumnIndices["Project Code"]].Text;
                if (!string.IsNullOrEmpty(MOProjectCode) && MOProjectCode.Contains(projectCodeFilter))
                {
                    count++;
                }
            }
            return count;
        }

        private int CountNonNullPOProjectCodeRows(string projectCodeFilter)
        {
            int count = 0;
            for (int row = 2; row <= PODetailsSheet.Dimension.Rows; row++)
            {
                string POProjectCode = PODetailsSheet.Cells[row, poColumnIndices["Project Code"]].Text;
                if (!string.IsNullOrEmpty(POProjectCode) && (POProjectCode.Contains(projectCodeFilter)))
                {
                    count++;
                }
            }
            return count;
        }

        private static List<int> FindRowIndices(ExcelWorksheet worksheet, string PRMaterialCode, string POMaterialCode, string PRProjectCode, string POProjectCode)
        {
            List<int> matchingRows = [];
            PRMaterialCode = PRMaterialCode.Trim().Replace(" ", "");
            POMaterialCode = POMaterialCode.Trim().Replace(" ", "");

            for (int row = 10; row <= worksheet.Dimension.End.Row; row++)
            {
                string sheetMaterialCode = worksheet.Cells[row, 4].Text.Trim().Replace(" ", "");
                string sheetProjectCode = worksheet.Cells[row, 3].Text.Trim().Replace(" ", "");
                string obsolete = worksheet.Cells[row, 25].Text.Trim().Replace(" ", "");

                bool materialCodeMatches = (!string.IsNullOrEmpty(PRMaterialCode) && sheetMaterialCode.Equals(PRMaterialCode, StringComparison.OrdinalIgnoreCase)) ||
                                           (!string.IsNullOrEmpty(POMaterialCode) && sheetMaterialCode.Equals(POMaterialCode, StringComparison.OrdinalIgnoreCase));

                bool projectCodeMatches = (!string.IsNullOrEmpty(PRProjectCode) && sheetProjectCode.Equals(PRProjectCode, StringComparison.OrdinalIgnoreCase)) ||
                                          (!string.IsNullOrEmpty(POProjectCode) && sheetProjectCode.Equals(POProjectCode, StringComparison.OrdinalIgnoreCase));

                bool isObsolete = obsolete == "N";

                if (materialCodeMatches && projectCodeMatches && isObsolete)
                {
                    matchingRows.Add(row);
                }
            }
            return matchingRows;
        }

        private static List<int> POFindRowIndices(ExcelWorksheet worksheet, string POProjectCode, string POMaterialCode, string POQty)
        {
            List<int> matchingRows = [];
            POProjectCode = POProjectCode.Trim().Replace(" ", "");
            POMaterialCode = POMaterialCode.Trim().Replace(" ", "");
            POQty = POQty.Trim().Replace(" ", "");

            _ = double.TryParse(POQty, out double poQty);
            string formattedPOQty = poQty.ToString("F1");

            for (int row = 10; row <= worksheet.Dimension.End.Row; row++)
            {
                string sheetProjectCode = worksheet.Cells[row, 3].Text.Trim().Replace(" ", "");
                string sheetMaterialCode = worksheet.Cells[row, 4].Text.Trim().Replace(" ", "");
                string sheetQty = worksheet.Cells[row, 5].Text.Trim().Replace(" ", "");
                string obsolete = worksheet.Cells[row, 25].Text.Trim().Replace(" ", "");

                _ = double.TryParse(sheetQty, out double SheetQty);
                string formattedSheetQty = SheetQty.ToString("F1");

                if (string.IsNullOrEmpty(sheetProjectCode))
                {
                    continue;
                }

                bool materialCodeMatches = (!string.IsNullOrEmpty(POMaterialCode) && sheetMaterialCode.Equals(POMaterialCode, StringComparison.OrdinalIgnoreCase));

                bool projectCodeMatches = (!string.IsNullOrEmpty(POProjectCode) && sheetProjectCode.Equals(POProjectCode, StringComparison.OrdinalIgnoreCase));

                bool qtyMatches = (!string.IsNullOrEmpty(POQty) && formattedSheetQty.Equals(formattedPOQty, StringComparison.OrdinalIgnoreCase));

                bool isObsolete = obsolete == "N";

                if (materialCodeMatches && projectCodeMatches && qtyMatches && isObsolete)
                {
                    matchingRows.Add(row);
                }
            }
            return matchingRows;
        }

        private static List<int> MOFindRowIndices(ExcelWorksheet worksheet, string MOProjectCode, string MOMaterialCode, string CollectedQuantity)
        {
            List<int> matchingRows = [];
            MOProjectCode = MOProjectCode.Trim().Replace(" ", "");
            MOMaterialCode = MOMaterialCode.Trim().Replace(" ", "");

            for (int row = 10; row <= worksheet.Dimension.End.Row; row++)
            {
                if (string.IsNullOrEmpty(CollectedQuantity) )
                {
                    continue;
                }

                string sheetMaterialCode = worksheet.Cells[row, 4].Text.Trim().Replace(" ", "");
                string sheetProjectCode = worksheet.Cells[row, 3].Text.Trim().Replace(" ", "");
                string obsolete = worksheet.Cells[row, 25].Text.Trim().Replace(" ", "");

                bool materialCodeMatches = (!string.IsNullOrEmpty(MOMaterialCode) && sheetMaterialCode.Equals(MOMaterialCode, StringComparison.OrdinalIgnoreCase));

                bool projectCodeMatches = (!string.IsNullOrEmpty(MOProjectCode) && sheetProjectCode.Equals(MOProjectCode, StringComparison.OrdinalIgnoreCase));

                bool isObsolete = obsolete == "N";

                if (materialCodeMatches && projectCodeMatches && isObsolete)
                {
                    matchingRows.Add(row);
                }
            }
            return matchingRows;
        }

        private static void InsertDataBatch(ExcelWorksheet worksheet, List<(int Row, DataList Data)> batchData, string saveFilePath)
        {
            using var package = new ExcelPackage(new FileInfo(saveFilePath));

            var wsFab = package.Workbook.Worksheets["Fab_Parts"];
            var wsStd = package.Workbook.Worksheets["Std_Parts"];
            var wsEE = package.Workbook.Worksheets["EE_Parts"];

            for (int i = 0; i < batchData.Count; i++)
            {
                var (row, data) = batchData[i];
                bool isDuplicate = false;
                bool qtyMatches = false;

                _ = double.TryParse(data.POQty, out double POQty);
                _ = double.TryParse(data.ReceivedQty, out double receivedQty);

                string sheetQtyString = worksheet.Cells[row, 5].Text.Replace(" ", "");
                _ = double.TryParse(sheetQtyString, out double sheetQty);
                _ = double.TryParse(data.PRQty, out double inputPRQty);

                string formattedSheetQty = sheetQty.ToString("F1");
                string formattedInputPRQty = inputPRQty.ToString("F1");

                for (int j = 0; j < i; j++)
                {
                    var (_, prevData) = batchData[j];
                    if (prevData.PRMaterialCode.Equals(data.PRMaterialCode, StringComparison.OrdinalIgnoreCase) &&
                        prevData.PRProjectCode.Equals(data.PRProjectCode, StringComparison.OrdinalIgnoreCase))
                    {
                        isDuplicate = true;

                        qtyMatches = formattedSheetQty.Equals(formattedInputPRQty, StringComparison.OrdinalIgnoreCase);
                        break;
                    }
                }

                if (!isDuplicate || qtyMatches)
                {
                    bool isPartialDelivery = POQty > 0 && POQty - receivedQty > 0;
                    bool isFullDelivery = POQty > 0 && POQty - receivedQty == 0;

                    worksheet.Cells[row, 11].Value = data.PRNo;
                    worksheet.Cells[row, 12].Value = data.PRApprovedOn;
                    worksheet.Cells[row, 13].Value = data.PONo;
                    worksheet.Cells[row, 16].Value = data.POApprovedOn;
                    if (worksheet.Name != "Fab_Parts")
                    {
                        worksheet.Cells[row, 18].Value = data.Supplier;
                        worksheet.Cells[row, 19].Value = data.ETA;
                    }

                    if (isFullDelivery && !string.IsNullOrEmpty(data.POApprovedOn) && !string.IsNullOrEmpty(data.ReceivedQty))
                        worksheet.Cells[row, 17].Value = "FD";
                    else if (isPartialDelivery && !string.IsNullOrEmpty(data.POApprovedOn) && !string.IsNullOrEmpty(data.ReceivedQty) && (receivedQty != 0.0))
                        worksheet.Cells[row, 17].Value = "PD";
                    else if (string.IsNullOrEmpty(data.POApprovedOn) && string.IsNullOrEmpty(data.PRApprovedOn) && string.IsNullOrEmpty(data.ReceivedQty) && !string.IsNullOrEmpty(data.PRNo))
                        worksheet.Cells[row, 17].Value = "PR PFA";
                    else if (string.IsNullOrEmpty(data.POApprovedOn) && !string.IsNullOrEmpty(data.PRApprovedOn) && string.IsNullOrEmpty(data.ReceivedQty) && (receivedQty == 0.0))
                        worksheet.Cells[row, 17].Value = "PO PFA";
                    else if (!string.IsNullOrEmpty(data.POApprovedOn))
                        worksheet.Cells[row, 17].Value = "APPROVED";
                }
            }
        }

        private static void InsertMODataBatch(ExcelWorksheet worksheet, List<(int Row, DataListMO Data)> batchMOData, string saveFilePath)
        {
            using var package = new ExcelPackage(new FileInfo(saveFilePath));

            var wsFab = package.Workbook.Worksheets["Fab_Parts"];
            var wsStd = package.Workbook.Worksheets["Std_Parts"];
            var wsEE = package.Workbook.Worksheets["EE_Parts"];

            foreach (var (row, dataMO) in batchMOData)
            {
                _ = double.TryParse(dataMO.CollectedQty, out double CollectedQty);

                string sheetQtyString = worksheet.Cells[row, 5].Text.Replace(" ", "");
                _ = double.TryParse(sheetQtyString, out double sheetQty);
                _ = double.TryParse(dataMO.MOQty, out double MOQty);

                string formattedSheetQty = sheetQty.ToString("F1");
                string formattedMOQty = MOQty.ToString("F1");
                bool qtyMatches = !string.IsNullOrEmpty(dataMO.CollectedQty) && formattedSheetQty.Equals(formattedMOQty, StringComparison.OrdinalIgnoreCase);

                double qtyFab = worksheet.Name == "Fab_Parts" ? GetQtyFromWorksheet(wsFab, [row]) : 0;
                double qtyStd = worksheet.Name == "Std_Parts" ? GetQtyFromWorksheet(wsStd, [row]) : 0;
                double qtyEE = worksheet.Name == "EE_Parts" ? GetQtyFromWorksheet(wsEE, [row]) : 0;

                bool isPartialTaken = (qtyFab > 0 && qtyFab - CollectedQty > 0) || (qtyStd > 0 && qtyStd - CollectedQty > 0) || (qtyEE > 0 && qtyEE - CollectedQty > 0);
                bool isTaken = (qtyFab > 0 && qtyFab - CollectedQty == 0) || (qtyStd > 0 && qtyStd - CollectedQty == 0) || (qtyEE > 0 && qtyEE - CollectedQty == 0);

                if (qtyMatches && (!CollectedQty.Equals(0)))
                {
                    if (isTaken && !string.IsNullOrEmpty(dataMO.CollectedQty))
                        worksheet.Cells[row, 17].Value = "TAKEN";
                    else if (isPartialTaken && !string.IsNullOrEmpty(dataMO.CollectedQty))
                        worksheet.Cells[row, 17].Value = "PT";
                }
            }
        }

        private static void InsertPODataBatch(ExcelWorksheet worksheet, List<(int Row, DataListPO Data)> batchPOData, string saveFilePath)
        {
            using var package = new ExcelPackage(new FileInfo(saveFilePath));

            var wsFab = package.Workbook.Worksheets["Fab_Parts"];
            var wsStd = package.Workbook.Worksheets["Std_Parts"];
            var wsEE = package.Workbook.Worksheets["EE_Parts"];

            foreach (var (row, dataPO) in batchPOData)
            {
                if (string.IsNullOrEmpty(dataPO.ReceivedQty)) {
                    worksheet.Cells[row, 13].Value = dataPO.PONo;
                    worksheet.Cells[row, 16].Value = dataPO.POApprovedOn;
                    worksheet.Cells[row, 17].Value = "APPROVED";
                }
            }
        }

        private static double GetQtyFromWorksheet(ExcelWorksheet worksheet, List<int> rows)
        {
            double Qty = 0;

            foreach (int row in rows)
            {
                if (double.TryParse(worksheet.Cells[row, 5].Text, out double quantity))
                {
                    Qty = quantity;
                }
            }

            return Qty;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    package?.Dispose();
                }

                disposed = true;
            }
        }
    }
}
