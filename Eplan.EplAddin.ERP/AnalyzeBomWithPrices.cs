using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using Microsoft.VisualBasic.FileIO;

namespace Eplan.EplAddin.ERP
{
    public class AnalyzeBomWithPrices : IEplAction, IEplActionEnable
    {
        private const string CsvPath = @"C:\EplanData\parts_supplier_full.csv";

        public bool Execute(ActionCallingContext ctx)
        {
            Project project = new SelectionSet().GetCurrentProject(true);
            if (project == null)
            {
                MessageBox.Show("Brak aktywnego projektu.", "Analyze BOM");
                return false;
            }

            var finder = new DMObjectsFinder(project);
            var functions = finder.GetFunctions(null).Where(f => f.IsMainFunction);

            var partCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in functions)
            {
                foreach (var ar in f.ArticleReferences)
                {
                    if (string.IsNullOrWhiteSpace(ar.PartNr)) continue;
                    if (!partCounts.ContainsKey(ar.PartNr))
                        partCounts[ar.PartNr] = 0;
                    partCounts[ar.PartNr]++;
                }
            }

            if (partCounts.Count == 0)
            {
                MessageBox.Show("Brak artykułów w funkcjach projektu.", "Analyze BOM");
                return false;
            }

            if (!File.Exists(CsvPath))
            {
                MessageBox.Show($"Nie znaleziono pliku CSV:\n{CsvPath}", "Analyze BOM");
                return false;
            }

            var data = new Dictionary<string, string[]>();
            var headerIndices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            using (var parser = new TextFieldParser(CsvPath))
            {
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                if (!parser.EndOfData)
                {
                    string[] headers = parser.ReadFields();
                    for (int i = 0; i < headers.Length; i++)
                        headerIndices[headers[i].Trim()] = i;
                }

                while (!parser.EndOfData)
                {
                    string[] row = parser.ReadFields();
                    string key = GetValue(row, "Part Number");
                    if (!string.IsNullOrWhiteSpace(key))
                        data[key] = row;
                }
            }

            string GetValue(string[] row, string col)
            {
                if (!headerIndices.TryGetValue(col, out int i)) return "";
                if (i < 0 || i >= row.Length) return "";
                return row[i].Trim();
            }

            var form = new Form
            {
                Text = "Analiza BOM – Dostępność i Ceny",
                Width = 1200,
                Height = 500,
                StartPosition = FormStartPosition.CenterScreen,
                KeyPreview = true
            };
            form.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Enter)
                    form.Close();
            };

            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True },
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            };

            dgv.Columns.Add("PartNr", "Nr artykułu");
            dgv.Columns.Add("Qty", "Zapotrzeb.");
            dgv.Columns.Add("Stock", "Magazyn");
            dgv.Columns.Add("LastPurchase", "Ostatnia cena PLN");
            dgv.Columns.Add("LastSupplier", "Ostatni dostawca");

            for (int i = 1; i <= 4; i++)
            {
                dgv.Columns.Add($"WH{i}_Name", $"H{i} Dostawca");
                dgv.Columns.Add($"WH{i}_FinalPrice", $"H{i} Cena PLN");
                dgv.Columns.Add($"WH{i}_Delivery", $"H{i} Dostawa (dni)");
            }

            var exportList = new List<Tuple<string, int, int, string, double, string>>();

            foreach (var kvp in partCounts)
            {
                string partNr = kvp.Key;
                int qty = kvp.Value;

                if (!data.TryGetValue(partNr, out var row))
                {
                    dgv.Rows.Add(partNr, qty, "?", "Brak w CSV", "", "", "", "", "", "", "", "", "", "");
                    continue;
                }

                int stock = int.TryParse(GetValue(row, "Stock Qty"), out var st) ? st : 0;
                double lastPln = double.TryParse(GetValue(row, "Last Purchase Price PLN"), NumberStyles.Any, CultureInfo.InvariantCulture, out var lp) ? lp : 0;
                string lastSupp = GetValue(row, "Last Supplier");

                double bestPrice = double.MaxValue;
                string bestSupp = "";
                string bestDeliv = "";

                var rowVals = new List<string>
                {
                    partNr,
                    qty.ToString(),
                    stock.ToString(),
                    lastPln.ToString("0.00"),
                    lastSupp
                };

                for (int i = 1; i <= 4; i++)
                {
                    string whName = GetValue(row, $"WH{i}_Name");
                    string whPrice = GetValue(row, $"WH{i}_Price");
                    string whDisc = GetValue(row, $"WH{i}_Discount");
                    string whDeliv = GetValue(row, $"WH{i}_Delivery");

                    if (string.IsNullOrWhiteSpace(whName))
                    {
                        rowVals.AddRange(new[] { "", "", "" });
                        continue;
                    }

                    double price = double.TryParse(whPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out var pr) ? pr : 0;
                    double discount = double.TryParse(whDisc, NumberStyles.Any, CultureInfo.InvariantCulture, out var di) ? di : 0;
                    string delivery = string.IsNullOrWhiteSpace(whDeliv) ? "brak info" : whDeliv;

                    double final = price * (1 - discount);
                    if (final < bestPrice)
                    {
                        bestPrice = final;
                        bestSupp = whName;
                        bestDeliv = delivery;
                    }

                    rowVals.AddRange(new[]
                    {
                        whName,
                        final.ToString("0.00"),
                        delivery
                    });
                }

                dgv.Rows.Add(rowVals.ToArray());
                exportList.Add(Tuple.Create(partNr, qty, stock, bestSupp, bestPrice, bestDeliv));
            }

            var exportBtn = new Button
            {
                Text = "📤 Eksportuj do Excela",
                Dock = DockStyle.Bottom,
                Height = 40
            };
            exportBtn.Click += (s, e) =>
            {
                try
                {
                    string fileName = $"BOM_{Path.GetFileNameWithoutExtension(project.ProjectFullName)}_{DateTime.Now:yyyy-MM-dd HH-mm}.xlsx";
                    string path = Path.Combine(@"C:\EplanData", fileName);

                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("BOM");
                        ws.Cell(1, 1).InsertTable(exportList.Select(x => new
                        {
                            NrArtykulu = x.Item1,
                            Zapotrzebowanie = x.Item2,
                            Magazyn = x.Item3,
                            Dostawca = x.Item4,
                            CenaPLN = x.Item5.ToString("0.00"),
                            Dostawa = x.Item6
                        }));

                        ws.Columns().AdjustToContents();
                        wb.SaveAs(path);
                    }

                    MessageBox.Show($"Zapisano plik:\n{path}", "Eksport zakończony");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Błąd eksportu: " + ex.Message);
                }
            };

            form.Controls.Add(dgv);
            form.Controls.Add(exportBtn);
            form.ShowDialog();
            return true;
        }

        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "AnalyzeBomWithPrices";
            Ordinal = 50;
            return true;
        }

        public void GetActionProperties(ref ActionProperties props) { }

        public bool Enabled(string strActionName, ActionCallingContext actionContext)
        {
            return new SelectionSet().GetCurrentProject(true) != null;
        }
    }
}