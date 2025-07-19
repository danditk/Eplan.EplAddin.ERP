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
using System.Drawing;

// Alias dla Label z WinForms, żeby uniknąć konfliktu z Eplan Label
using WinFormsLabel = System.Windows.Forms.Label;

namespace Eplan.EplAddin.ERP
{
    public class AnalyzeBomWithPrices : IEplAction, IEplActionEnable
    {
        private const string CsvPath = @"C:\\EplanData\\parts_supplier_full.csv";

        private Form _form;
        private WinFormsLabel _lblInfo;
        private WinFormsLabel _lblMinBudgetInfo;
        private WinFormsLabel _lblMinDeadlineInfo;
        private TextBox _txtBudget;
        private DateTimePicker _dtpDeadline;
        private Button _btnRecalc;
        private DataGridView _dgv;

        private Dictionary<string, PartData> _partsData;

        private class PartData
        {
            public int Stock;
            public decimal LastPurchasePrice;
            public string LastSupplier;
            public List<SupplierOffer> Offers = new List<SupplierOffer>();
        }

        private class SupplierOffer
        {
            public string Name;
            public decimal Price;
            public int DeliveryDays;
        }

        public bool Execute(ActionCallingContext ctx)
        {
            Project project = new SelectionSet().GetCurrentProject(true);
            if (project == null)
            {
                MessageBox.Show("Brak aktywnego projektu.", "Analyze BOM");
                return false;
            }

            var projectPartCounts = GetProjectPartCounts(project);
            if (projectPartCounts.Count == 0)
            {
                MessageBox.Show("Brak artykułów w funkcjach projektu.", "Analyze BOM");
                return false;
            }

            if (!File.Exists(CsvPath))
            {
                MessageBox.Show($"Nie znaleziono pliku CSV:\n{CsvPath}", "Analyze BOM");
                return false;
            }

            _partsData = LoadPartsData();

            InitializeForm();

            RecalculateSuppliers(0m, DateTime.Now.Date.AddMonths(1)); // Startowa kalkulacja bez limitów

            _form.ShowDialog();

            return true;
        }

        private Dictionary<string, PartData> LoadPartsData()
        {
            var data = new Dictionary<string, PartData>(StringComparer.OrdinalIgnoreCase);
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
                    string partNr = GetValue(row, "Part Number", headerIndices);
                    if (string.IsNullOrWhiteSpace(partNr))
                        continue;

                    var partData = new PartData();

                    partData.Stock = int.TryParse(GetValue(row, "Stock", headerIndices), out var stock) ? stock : 0;
                    partData.LastPurchasePrice = decimal.TryParse(GetValue(row, "Last Purchase Price PLN", headerIndices), out var lpp) ? lpp : 0m;
                    partData.LastSupplier = GetValue(row, "Last Supplier", headerIndices);

                    // Wczytujemy oferty hurtowników H1-H4
                    for (int i = 1; i <= 4; i++)
                    {
                        string supName = GetValue(row, $"H{i}_Name", headerIndices);
                        string priceStr = GetValue(row, $"H{i}_FinalPrice PLN", headerIndices);
                        string deliveryStr = GetValue(row, $"H{i}_Delivery (days)", headerIndices);

                        if (!string.IsNullOrWhiteSpace(supName) &&
                            decimal.TryParse(priceStr, out decimal price) &&
                            int.TryParse(deliveryStr, out int delivery))
                        {
                            partData.Offers.Add(new SupplierOffer
                            {
                                Name = supName,
                                Price = price,
                                DeliveryDays = delivery
                            });
                        }
                    }

                    data[partNr] = partData;
                }
            }

            string GetValue(string[] row, string col, Dictionary<string, int> hdr)
            {
                if (!hdr.TryGetValue(col, out int idx)) return "";
                if (idx < 0 || idx >= row.Length) return "";
                return row[idx].Trim();
            }

            return data;
        }

        private Dictionary<string, int> GetProjectPartCounts(Project project)
        {
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

            return partCounts;
        }

        private void InitializeForm()
        {
            _form = new Form
            {
                Text = "Analiza BOM – Dostępność i Ceny",
                Width = 1400,
                Height = 600,
                StartPosition = FormStartPosition.CenterScreen,
                KeyPreview = true
            };

            _form.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Enter)
                    _form.Close();
            };

            _lblInfo = new WinFormsLabel
            {
                Text = "Minimalny budżet i deadline zostaną wyliczone na podstawie danych.",
                Dock = DockStyle.Top,
                Height = 20,
                ForeColor = Color.DarkRed,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _lblMinBudgetInfo = new WinFormsLabel
            {
                Text = "Minimalny budżet: 0 PLN",
                Dock = DockStyle.Top,
                Height = 20,
                ForeColor = Color.DarkBlue,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _lblMinDeadlineInfo = new WinFormsLabel
            {
                Text = "Minimalny deadline: --",
                Dock = DockStyle.Top,
                Height = 20,
                ForeColor = Color.DarkBlue,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.LightGray
            };

            var lblBudget = new WinFormsLabel { Text = "Budżet (PLN):", Location = new Point(10, 10), AutoSize = true };
            _txtBudget = new TextBox { Location = new Point(100, 8), Width = 100, Text = "000" };
            var lblDeadline = new WinFormsLabel { Text = "Deadline:", Location = new Point(220, 10), AutoSize = true };
            _dtpDeadline = new DateTimePicker { Location = new Point(290, 8), Width = 120, Value = DateTime.Now.AddDays(30) };
            _btnRecalc = new Button { Text = "PRZELICZ", Location = new Point(430, 6), Size = new Size(100, 25) };

            _btnRecalc.Click += BtnRecalc_Click;

            panel.Controls.Add(lblBudget);
            panel.Controls.Add(_txtBudget);
            panel.Controls.Add(lblDeadline);
            panel.Controls.Add(_dtpDeadline);
            panel.Controls.Add(_btnRecalc);

            _dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True },
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            };

            _dgv.Columns.Add("PartNr", "Nr artykułu");
            _dgv.Columns.Add("Qty", "Zapotrzebowanie");
            _dgv.Columns.Add("Stock", "Magazyn");
            _dgv.Columns.Add("LastPurchasePrice", "Ostatnia cena PLN");
            _dgv.Columns.Add("LastSupplier", "Ostatni dostawca");

            for (int i = 1; i <= 4; i++)
            {
                _dgv.Columns.Add($"WH{i}_Name", $"H{i} Dostawca");
                _dgv.Columns.Add($"WH{i}_FinalPrice", $"H{i} Cena PLN");
                _dgv.Columns.Add($"WH{i}_Delivery", $"H{i} Dostawa (dni)");
            }

            // Dodajemy kontrolki do formularza w odpowiedniej kolejności
            _form.Controls.Add(_dgv);
            _form.Controls.Add(panel);
            _form.Controls.Add(_lblMinBudgetInfo);
            _form.Controls.Add(_lblMinDeadlineInfo);
            _form.Controls.Add(_lblInfo);
        }

        private void BtnRecalc_Click(object sender, EventArgs e)
        {
            int maxDeliveryDays = _partsData.Values.SelectMany(p => p.Offers).Select(o => o.DeliveryDays).DefaultIfEmpty(0).Max();
            DateTime minAllowedDeadline = DateTime.Now.Date.AddDays(maxDeliveryDays);

            DateTime selectedDeadline = _dtpDeadline.Value.Date;
            if (selectedDeadline < minAllowedDeadline)
            {
                MessageBox.Show($"Deadline nie może być wcześniejszy niż {minAllowedDeadline:yyyy-MM-dd}. Uwzględnij dni robocze!", "Błąd daty");
                return;
            }

            if (!decimal.TryParse(_txtBudget.Text.Trim(), out decimal budget))
            {
                MessageBox.Show("Niepoprawna wartość budżetu.", "Błąd");
                return;
            }
            if (budget < 0) budget = 0;

            var projectPartCounts = GetProjectPartCounts(new SelectionSet().GetCurrentProject(true));
            decimal minBudgetNeeded = 0m;
            foreach (var kvp in projectPartCounts)
            {
                if (_partsData.TryGetValue(kvp.Key, out var partInfo))
                {
                    minBudgetNeeded += partInfo.LastPurchasePrice * kvp.Value;
                }
            }

            _lblMinBudgetInfo.Text = $"Minimalny budżet: {minBudgetNeeded:F2} PLN";
            _lblMinDeadlineInfo.Text = $"Minimalny deadline: {minAllowedDeadline:yyyy-MM-dd}";

            if (budget > 0 && budget < minBudgetNeeded)
            {
                MessageBox.Show($"Budżet nie może być mniejszy niż minimalny koszt zamówienia: {minBudgetNeeded:F2} PLN.", "Błąd budżetu");
                return;
            }

            RecalculateSuppliers(budget, selectedDeadline);
        }

        private void RecalculateSuppliers(decimal budget, DateTime deadline)
        {
            _dgv.Rows.Clear();

            var projectPartCounts = GetProjectPartCounts(new SelectionSet().GetCurrentProject(true));

            foreach (var kvp in projectPartCounts)
            {
                string partNr = kvp.Key;
                int qty = kvp.Value;

                if (!_partsData.TryGetValue(partNr, out var partInfo))
                {
                    _dgv.Rows.Add(partNr, qty, "BRAK", "-", "-");
                    for (int i = 1; i <= 4; i++)
                        _dgv.Rows[_dgv.Rows.Count - 1].Cells[$"WH{i}_Name"].Value = "";
                    continue;
                }

                string stockDisplay = partInfo.Stock >= qty ? partInfo.Stock.ToString() : $"{partInfo.Stock} (NIEDOSTATECZNY)";
                string lastPurchasePrice = partInfo.LastPurchasePrice > 0 ? partInfo.LastPurchasePrice.ToString("F2") : "-";
                string lastSupplier = string.IsNullOrWhiteSpace(partInfo.LastSupplier) ? "-" : partInfo.LastSupplier;

                SupplierOffer bestOffer = null;

                if (partInfo.Stock >= qty)
                {
                    // Bierzemy z magazynu
                    bestOffer = null;
                }
                else
                {
                    var suitableOffers = partInfo.Offers
                        .Where(o => o.DeliveryDays <= (deadline - DateTime.Now.Date).Days)
                        .OrderBy(o => o.Price)
                        .ToList();

                    foreach (var offer in suitableOffers)
                    {
                        decimal totalPrice = offer.Price * qty;
                        if (budget == 0 || totalPrice <= budget)
                        {
                            bestOffer = offer;
                            break;
                        }
                    }

                    if (bestOffer == null && suitableOffers.Count > 0)
                        bestOffer = suitableOffers.First();
                }

                var rowIndex = _dgv.Rows.Add();
                var row = _dgv.Rows[rowIndex];

                row.Cells["PartNr"].Value = partNr;
                row.Cells["Qty"].Value = qty;
                row.Cells["Stock"].Value = stockDisplay;
                row.Cells["LastPurchasePrice"].Value = lastPurchasePrice;
                row.Cells["LastSupplier"].Value = lastSupplier;

                for (int i = 1; i <= 4; i++)
                {
                    var offerCsv = partInfo.Offers.ElementAtOrDefault(i - 1);
                    if (offerCsv != null)
                    {
                        row.Cells[$"WH{i}_Name"].Value = offerCsv.Name;
                        row.Cells[$"WH{i}_FinalPrice"].Value = offerCsv.Price.ToString("F2");
                        row.Cells[$"WH{i}_Delivery"].Value = offerCsv.DeliveryDays;
                    }
                    else
                    {
                        row.Cells[$"WH{i}_Name"].Value = "";
                        row.Cells[$"WH{i}_FinalPrice"].Value = "";
                        row.Cells[$"WH{i}_Delivery"].Value = "";
                    }
                }

                if (bestOffer == null)
                {
                    row.Cells["Stock"].Style.BackColor = Color.LightGreen;
                }
                else
                {
                    for (int i = 1; i <= 4; i++)
                    {
                        var supName = row.Cells[$"WH{i}_Name"].Value?.ToString();
                        if (supName == bestOffer.Name)
                        {
                            row.Cells[$"WH{i}_Name"].Style.BackColor = Color.LightGreen;
                            row.Cells[$"WH{i}_FinalPrice"].Style.BackColor = Color.LightGreen;
                            row.Cells[$"WH{i}_Delivery"].Style.BackColor = Color.LightGreen;
                        }
                        else
                        {
                            row.Cells[$"WH{i}_Name"].Style.BackColor = Color.White;
                            row.Cells[$"WH{i}_FinalPrice"].Style.BackColor = Color.White;
                            row.Cells[$"WH{i}_Delivery"].Style.BackColor = Color.White;
                        }
                    }
                }
            }
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
