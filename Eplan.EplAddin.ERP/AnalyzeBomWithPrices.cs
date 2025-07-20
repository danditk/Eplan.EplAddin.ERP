using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;

// alias dla Label z WinForms
using WinFormsLabel = System.Windows.Forms.Label;

public class AnalyzeBomWithPrices : IEplAction
{
    private DateTimePicker _dtpDeadline;
    private TextBox _txtBudget;
    private DataGridView _dgv;
    private List<ExportRow> _exportData;

    private class ExportRow
    {
        public string PartNo;
        public int Qty;
        public int InternalStock;
        public string Supplier;
        public decimal Price;
        public int DeliveryDays;
    }

    public bool Execute(ActionCallingContext ctx)
    {
        string csvPath = @"C:\EplanData\test_database_bom_szlifierka_extended.csv";
        if (!File.Exists(csvPath))
        {
            MessageBox.Show($"Nie znaleziono pliku CSV: {csvPath}", "Błąd");
            return false;
        }

        // Wczytanie CSV
        var lines = File.ReadAllLines(csvPath);
        var header = lines[0].Split(',');
        var rows = lines.Skip(1)
                        .Select(l => l.Split(','))
                        .Where(r => r.Length == header.Length)
                        .ToList();

        // Mapa PartNo → row[]
        var partRows = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        int idxPartNo = Array.IndexOf(header, "PartNo");
        foreach (var row in rows)
        {
            var pn = row[idxPartNo];
            if (!string.IsNullOrWhiteSpace(pn))
                partRows[pn] = row;
        }

        // Pobranie BOM z projektu
        var project = new SelectionSet().GetCurrentProject(true);
        if (project == null)
        {
            MessageBox.Show("Brak aktywnego projektu.", "Błąd");
            return false;
        }
        var bom = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in project.Pages)
        {
            foreach (Function func in page.Functions)
            {
                foreach (var ar in func.ArticleReferences)
                {
                    var pn = ar.PartNr;
                    if (string.IsNullOrWhiteSpace(pn)) continue;
                    if (bom.ContainsKey(pn)) bom[pn]++; else bom[pn] = 1;
                }
            }
        }
        if (bom.Count == 0)
        {
            MessageBox.Show("Brak artykułów w BOMie.", "Info");
            return false;
        }

        // GUI
        var form = new Form
        {
            Text = "Analiza BOM – wybór dostawcy",
            Width = 1000,
            Height = 600,
            StartPosition = FormStartPosition.CenterScreen
        };

        var panel = new Panel { Dock = DockStyle.Top, Height = 50 };
        form.Controls.Add(panel);

        // Deadline
        var lblDeadline = new WinFormsLabel { Text = "Deadline:", Location = new Point(10, 15), AutoSize = true };
        panel.Controls.Add(lblDeadline);
        _dtpDeadline = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd",
            Value = DateTime.Now.AddDays(7),
            Location = new Point(70, 12)
        };
        panel.Controls.Add(_dtpDeadline);

        // Budget
        var lblBudget = new WinFormsLabel { Text = "Budżet [PLN]:", Location = new Point(240, 15), AutoSize = true };
        panel.Controls.Add(lblBudget);
        _txtBudget = new TextBox { Text = "0", Location = new Point(320, 12), Width = 100 };
        panel.Controls.Add(_txtBudget);

        // Buttons
        var btnRecalc = new Button { Text = "PRZELICZ", Location = new Point(440, 10), Width = 100 };
        panel.Controls.Add(btnRecalc);
        var btnExport = new Button { Text = "Eksportuj do Excela", Location = new Point(560, 10), Width = 150 };
        panel.Controls.Add(btnExport);

        // DataGridView
        _dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        };
        form.Controls.Add(_dgv);

        // Eventy
        btnRecalc.Click += (s, e) =>
        {
            if (!decimal.TryParse(_txtBudget.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var budget))
            {
                MessageBox.Show("Nieprawidłowy budżet.", "Błąd");
                return;
            }
            Recalculate(bom, partRows, budget, _dtpDeadline.Value.Date, header);
        };
        btnExport.Click += (s, e) =>
        {
            ExportToExcel(project, _exportData, _dtpDeadline.Value.Date, decimal.TryParse(_txtBudget.Text, out var b) ? b : 0m);
        };

        // Pierwsze przeliczenie
        btnRecalc.PerformClick();

        form.ShowDialog();
        return true;
    }

    private void Recalculate(
        Dictionary<string, int> bom,
        Dictionary<string, string[]> partRows,
        decimal budget,
        DateTime deadline,
        string[] header)
    {
        _exportData = new List<ExportRow>();
        var dt = new DataTable();
        dt.Columns.Add("Nr artykułu");
        dt.Columns.Add("Zapotrzebowanie");
        dt.Columns.Add("Magazyn");
        dt.Columns.Add("Dostawca");
        dt.Columns.Add("Cena PLN");
        dt.Columns.Add("Dostawa (dni)");

        int daysAllowed = (deadline - DateTime.Now.Date).Days;

        foreach (var kv in bom)
        {
            string pn = kv.Key;
            int qty = kv.Value;
            if (!partRows.TryGetValue(pn, out var row)) continue;

            // Indeksy
            int idxInternal = Array.IndexOf(header, "InternalStock");
            int idxLastPrice = Array.IndexOf(header, "LastPrice");

            int internalStock = int.TryParse(row[idxInternal], out var s) ? s : 0;
            decimal lastPrice = decimal.TryParse(row[idxLastPrice], NumberStyles.Any, CultureInfo.InvariantCulture, out var lp) ? lp : 0m;

            // Magazyn
            if (internalStock >= qty)
            {
                _exportData.Add(new ExportRow
                {
                    PartNo = pn,
                    Qty = qty,
                    InternalStock = internalStock,
                    Supplier = "Magazyn",
                    Price = lastPrice,
                    DeliveryDays = 0
                });
                continue;
            }

            // Hurtownie
            var offers = new List<ExportRow>();
            foreach (var sup in new[] { "TME", "RS", "Farnell", "Conrad", "Elfa" })
            {
                int ip = Array.IndexOf(header, $"{sup}_Price");
                int is2 = Array.IndexOf(header, $"{sup}_Stock");
                if (!decimal.TryParse(row[ip], NumberStyles.Any, CultureInfo.InvariantCulture, out var supPrice)) continue;
                if (!int.TryParse(row[is2], out var supStock) || supStock < qty) continue;
                int del = sup == "TME" ? 2 : sup == "RS" ? 3 : sup == "Farnell" ? 5 : sup == "Conrad" ? 6 : 4;
                if (del > daysAllowed) continue;
                decimal total = supPrice * qty;
                if (budget > 0 && total > budget) continue;
                offers.Add(new ExportRow
                {
                    PartNo = pn,
                    Qty = qty,
                    InternalStock = internalStock,
                    Supplier = sup,
                    Price = supPrice,
                    DeliveryDays = del
                });
            }

            var best = offers.OrderBy(o => o.Price).FirstOrDefault();
            if (best != null)
                _exportData.Add(best);
        }

        // Fill grid
        var table = new DataTable();
        table.Columns.Add("Nr artykułu");
        table.Columns.Add("Zapotrzebowanie");
        table.Columns.Add("Magazyn");
        table.Columns.Add("Dostawca");
        table.Columns.Add("Cena PLN");
        table.Columns.Add("Dostawa (dni)");

        foreach (var r in _exportData)
        {
            table.Rows.Add(r.PartNo, r.Qty, r.InternalStock, r.Supplier, r.Price.ToString("0.00"), r.DeliveryDays);
        }

        _dgv.DataSource = table;
    }

    private void ExportToExcel(
        Project project,
        List<ExportRow> data,
        DateTime deadline,
        decimal budget)
    {
        string projName = project?.ProjectName ?? "Projekt";
        string ts = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
        string fileName = $@"C:\EplanData\BOM_{projName}_{ts}.xlsx";

        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("BOM");

            // Podsumowanie
            ws.Cell(1, 1).Value = "Deadline:";
            ws.Cell(1, 2).Value = deadline.ToString("yyyy-MM-dd");
            ws.Cell(2, 1).Value = "Budget:";
            ws.Cell(2, 2).Value = budget;

            // Headers
            var headers = new[] { "Nr artykułu", "Zapotrzebowanie", "Magazyn", "Dostawca", "Cena PLN", "Dostawa (dni)" };
            for (int i = 0; i < headers.Length; i++)
                ws.Cell(4, i + 1).Value = headers[i];

            // Data
            for (int r = 0; r < data.Count; r++)
            {
                ws.Cell(5 + r, 1).Value = data[r].PartNo;
                ws.Cell(5 + r, 2).Value = data[r].Qty;
                ws.Cell(5 + r, 3).Value = data[r].InternalStock;
                ws.Cell(5 + r, 4).Value = data[r].Supplier;
                ws.Cell(5 + r, 5).Value = data[r].Price;
                ws.Cell(5 + r, 6).Value = data[r].DeliveryDays;
            }

            wb.SaveAs(fileName);
        }

        MessageBox.Show($"Zapisano plik:\n{fileName}", "Eksport OK");
    }

    public bool OnRegister(ref string Name, ref int Ordinal)
    {
        Name = "AnalyzeBomWithPrices";
        Ordinal = 50;
        return true;
    }

    public void GetActionProperties(ref ActionProperties properties)
    {
        // pusta implementacja
    }
}
