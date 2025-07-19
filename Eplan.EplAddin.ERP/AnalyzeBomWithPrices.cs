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
            // 1. pobranie projektu
            Project project = new SelectionSet().GetCurrentProject(true);
            if (project == null)
            {
                MessageBox.Show("Brak aktywnego projektu.", "Analyze BOM");
                return false;
            }

            // 2. odczyt BOM
            var finder = new DMObjectsFinder(project);
            var functions = finder.GetFunctions(null).Where(f => f.IsMainFunction);

            var partCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in functions)
                foreach (var ar in f.ArticleReferences)
                {
                    if (string.IsNullOrWhiteSpace(ar.PartNr))
                        continue;

                    int count;
                    if (!partCounts.TryGetValue(ar.PartNr, out count))
                        partCounts[ar.PartNr] = 1;
                    else
                        partCounts[ar.PartNr] = count + 1;
                }

            if (partCounts.Count == 0)
            {
                MessageBox.Show("Brak artykułów w funkcjach projektu.", "Analyze BOM");
                return false;
            }

            // 3. wczytanie CSV
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
                    var headers = parser.ReadFields();
                    for (int i = 0; i < headers.Length; i++)
                        headerIndices[headers[i].Trim()] = i;
                }
                while (!parser.EndOfData)
                {
                    var row = parser.ReadFields();
                    var key = GetValue(row, "Part Number");
                    if (!string.IsNullOrWhiteSpace(key))
                        data[key] = row;
                }
            }

            string GetValue(string[] row, string col)
            {
                int idx;
                if (!headerIndices.TryGetValue(col, out idx))
                    return "";
                if (idx < 0 || idx >= row.Length)
                    return "";
                return row[idx].Trim();
            }

            // 4. budowa formularza
            var form = new Form
            {
                Text = "Analiza BOM – Dostępność i Ceny",
                Width = 1200,
                Height = 600,
                StartPosition = FormStartPosition.CenterScreen
            };
            form.KeyPreview = true;
            form.KeyDown += (object sender, KeyEventArgs e) => {
                if (e.KeyCode == Keys.Escape)
                    form.Close();
            };

            // 4a. panel kryteriów
            var panelCrit = new Panel { Dock = DockStyle.Top, Height = 40 };
            var dtp = new DateTimePicker { Value = DateTime.Today.AddDays(7), Width = 120 };
            var tbBudget = new TextBox { Text = "0", Width = 80 };
            var btnRecalc = new Button { Text = "Przelicz", AutoSize = true };

            var lbl1 = new System.Windows.Forms.Label { Text = "Deadline:", AutoSize = true, Left = 10, Top = 12 };
            dtp.Left = lbl1.Right + 5;
            var lbl2 = new System.Windows.Forms.Label { Text = "Budżet PLN:", AutoSize = true, Left = dtp.Right + 20, Top = 12 };
            tbBudget.Left = lbl2.Right + 5;
            btnRecalc.Left = tbBudget.Right + 20;

            panelCrit.Controls.AddRange(new Control[] { lbl1, dtp, lbl2, tbBudget, btnRecalc });
            form.Controls.Add(panelCrit);

            // 4b. DataGridView
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True }
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
            form.Controls.Add(dgv);

            // 4c. przycisk eksportu
            var exportBtn = new Button
            {
                Text = "📤 Eksportuj do Excela",
                Dock = DockStyle.Bottom,
                Height = 40
            };
            form.Controls.Add(exportBtn);

            // 5. lista do eksportu
            var exportList = new List<Tuple<string, int, int, string, double, string>>();

            // 6. metoda przeliczająca
            void Recalculate()
            {
                dgv.Rows.Clear();
                exportList.Clear();

                // dni do deadline
                int daysToDeadline = (dtp.Value.Date - DateTime.Today).Days;
                if (daysToDeadline < 0) daysToDeadline = 0;

                // parsowanie budżetu
                double budget;
                if (!double.TryParse(tbBudget.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out budget))
                    budget = double.MaxValue;

                foreach (var kvp in partCounts)
                {
                    string partNr = kvp.Key;
                    int qty = kvp.Value;
                    if (!data.TryGetValue(partNr, out var row))
                    {
                        dgv.Rows.Add(partNr, qty, "?", "Brak w CSV");
                        continue;
                    }

                    int stock = int.TryParse(GetValue(row, "Stock Qty"), out var st) ? st : 0;
                    double lastPln = double.TryParse(GetValue(row, "Last Purchase Price PLN"),
                                                     NumberStyles.Any, CultureInfo.InvariantCulture, out var lp) ? lp : 0;
                    string lastSupp = GetValue(row, "Last Supplier");

                    double bestPrice = double.MaxValue;
                    string bestSupp = "";
                    string bestDeliv = "";

                    var rowVals = new List<string>{
                        partNr,
                        qty.ToString(),
                        stock.ToString(),
                        lastPln.ToString("0.00"),
                        lastSupp
                    };

                    for (int i = 1; i <= 4; i++)
                    {
                        string whName = GetValue(row, $"WH{i}_Name");
                        if (string.IsNullOrWhiteSpace(whName))
                        {
                            rowVals.AddRange(new[] { "", "", "" });
                            continue;
                        }

                        double price = double.TryParse(GetValue(row, $"WH{i}_Price"),
                                                       NumberStyles.Any, CultureInfo.InvariantCulture, out var pr) ? pr : 0;
                        double disc = double.TryParse(GetValue(row, $"WH{i}_Discount"),
                                                       NumberStyles.Any, CultureInfo.InvariantCulture, out var di) ? di : 0;
                        string delivStr = GetValue(row, $"WH{i}_Delivery");
                        int delivDays;
                        if (!int.TryParse(delivStr, out delivDays))
                            delivDays = int.MaxValue;

                        double final = price * (1 - disc);
                        bool okDeadline = delivDays <= daysToDeadline;
                        bool okBudget = (final * qty) <= budget;

                        if (okDeadline && okBudget && final < bestPrice)
                        {
                            bestPrice = final;
                            bestSupp = whName;
                            bestDeliv = delivStr;
                        }

                        rowVals.AddRange(new[]{
                            whName,
                            final.ToString("0.00"),
                            delivStr
                        });
                    }

                    dgv.Rows.Add(rowVals.ToArray());
                    exportList.Add(Tuple.Create(partNr, qty, stock, bestSupp, bestPrice, bestDeliv));
                }
            }

            // podłączenie przycisków
            btnRecalc.Click += (object sender, EventArgs e) => Recalculate();
            exportBtn.Click += (object sender, EventArgs e) =>
            {
                try
                {
                    string fn = $"BOM_{Path.GetFileNameWithoutExtension(project.ProjectFullName)}_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx";
                    string outPath = Path.Combine(@"C:\EplanData", fn);
                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("BOM");
                        ws.Cell(1, 1).InsertTable(exportList.Select(x => new {
                            NrArtykulu = x.Item1,
                            Zapotrzebowanie = x.Item2,
                            Magazyn = x.Item3,
                            Dostawca = x.Item4,
                            CenaPLN = x.Item5.ToString("0.00"),
                            Dostawa = x.Item6
                        }));
                        ws.Columns().AdjustToContents();
                        wb.SaveAs(outPath);
                    }
                    MessageBox.Show($"Zapisano plik:\n{outPath}", "Eksport zakończony");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Błąd eksportu: " + ex.Message);
                }
            };

            // pierwsze wypełnienie
            Recalculate();
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
