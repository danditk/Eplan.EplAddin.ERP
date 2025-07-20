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

// alias by uniknąć konfliktu z Label z EPLAN
using WinFormsLabel = System.Windows.Forms.Label;

public class AnalyzeBomWithPrices : IEplAction
{
    private DateTimePicker _dtpDeadline;
    private TextBox _txtBudget;
    private RadioButton _rbDeadlineFirst;
    private RadioButton _rbBudgetFirst;
    private DataGridView _dgv;
    private List<SelectedOffer> _offers;

    // reprezentuje wybraną ofertę dla artykułu
    private class SelectedOffer
    {
        public string PartNo;
        public int Qty;
        public int InternalStock;
        public decimal LastPurchasePrice;
        public string Supplier;
        public decimal Price;            // po rabacie
        public int DeliveryDays;
    }

    public bool Execute(ActionCallingContext ctx)
    {
        // 1) Wczytanie CSV
        string csvPath = @"C:\EplanData\test_database_bom_szlifierka_extended.csv";
        if (!File.Exists(csvPath))
        {
            MessageBox.Show($"Nie znaleziono pliku CSV:\n{csvPath}", "Błąd");
            return false;
        }
        var lines = File.ReadAllLines(csvPath);
        var header = lines[0].Split(',');
        var csvRows = lines.Skip(1)
                          .Select(r => r.Split(','))
                          .Where(r => r.Length == header.Length)
                          .ToDictionary(r => r[Array.IndexOf(header, "PartNo")], StringComparer.OrdinalIgnoreCase);

        // 2) Zliczenie BOM z projektu
        var project = new SelectionSet().GetCurrentProject(true);
        if (project == null)
        {
            MessageBox.Show("Brak aktywnego projektu.", "Błąd");
            return false;
        }
        var bom = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in project.Pages)
            foreach (Function f in page.Functions)
                foreach (var ar in f.ArticleReferences)
                    if (!string.IsNullOrWhiteSpace(ar.PartNr))
                        bom[ar.PartNr] = bom.ContainsKey(ar.PartNr) ? bom[ar.PartNr] + 1 : 1;

        if (bom.Count == 0)
        {
            MessageBox.Show("Brak artykułów w BOMie.", "Info");
            return false;
        }

        // 3) Budowa GUI
        var form = new Form
        {
            Text = "Analiza BOM – wybór dostawcy",
            Width = 1200,
            Height = 650,
            StartPosition = FormStartPosition.CenterScreen
        };

        var panel = new Panel { Dock = DockStyle.Top, Height = 80, BackColor = Color.WhiteSmoke };
        form.Controls.Add(panel);

        // Deadline
        panel.Controls.Add(new WinFormsLabel { Text = "Deadline:", Location = new Point(10, 10), AutoSize = true });
        _dtpDeadline = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd",
            Value = DateTime.Now.AddYears(1),
            Location = new Point(10, 30)
        };
        panel.Controls.Add(_dtpDeadline);

        // Budget
        panel.Controls.Add(new WinFormsLabel { Text = "Budget [PLN]:", Location = new Point(200, 10), AutoSize = true });
        _txtBudget = new TextBox { Text = "0", Location = new Point(200, 30), Width = 100 };
        panel.Controls.Add(_txtBudget);

        // Priorytet
        _rbDeadlineFirst = new RadioButton { Text = "Priorytet: Deadline", Location = new Point(350, 10), Checked = true };
        _rbBudgetFirst = new RadioButton { Text = "Priorytet: Budget", Location = new Point(350, 30) };
        panel.Controls.Add(_rbDeadlineFirst);
        panel.Controls.Add(_rbBudgetFirst);

        // Buttons
        var btnRecalc = new Button { Text = "PRZELICZ", Location = new Point(540, 25), Width = 100 };
        panel.Controls.Add(btnRecalc);
        var btnExport = new Button { Text = "Eksport do Excela", Location = new Point(660, 25), Width = 140 };
        panel.Controls.Add(btnExport);

        // DataGridView
        _dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            AllowUserToAddRows = false
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
            Recalculate(bom, csvRows, header, budget, _dtpDeadline.Value.Date);
        };
        btnExport.Click += (s, e) =>
        {
            ExportToExcel(project.ProjectName, _offers, _dtpDeadline.Value.Date,
                          decimal.TryParse(_txtBudget.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var b) ? b : 0m,
                          _rbDeadlineFirst.Checked ? "Deadline" : "Budget");
        };

        // Pierwsze przeliczenie
        btnRecalc.PerformClick();

        form.ShowDialog();
        return true;
    }

    private void Recalculate(
        Dictionary<string, int> bom,
        Dictionary<string, string[]> csvRows,
        string[] header,
        decimal budget,
        DateTime deadline)
    {
        _offers = new List<SelectedOffer>();
        var dt = new DataTable();
        dt.Columns.Add("Nr artykułu");
        dt.Columns.Add("Zapotrzebowanie");
        dt.Columns.Add("Magazyn");
        dt.Columns.Add("Ostatnia cena zakupu");
        dt.Columns.Add("Dostawca");
        dt.Columns.Add("Cena po rabacie");
        dt.Columns.Add("Dostawa (dni)");

        int daysAllowed = (deadline - DateTime.Now.Date).Days;

        foreach (var kv in bom)
        {
            string pn = kv.Key;
            int qty = kv.Value;
            if (!csvRows.TryGetValue(pn, out var row)) continue;

            // Indeksy
            int iStock = Array.IndexOf(header, "InternalStock");
            int iLastPrice = Array.IndexOf(header, "LastPrice");

            int internalStock = int.TryParse(row[iStock], out var s) ? s : 0;
            decimal lastPurchase = decimal.TryParse(row[iLastPrice], NumberStyles.Any, CultureInfo.InvariantCulture, out var lp) ? lp : 0m;

            // Lista ofert (magazyn + hurtownie)
            var allOffers = new List<SelectedOffer>
            {
                new SelectedOffer
                {
                    PartNo = pn, Qty = qty, InternalStock = internalStock,
                    LastPurchasePrice = lastPurchase,
                    Supplier = "Magazyn", Price = lastPurchase, DeliveryDays = 0
                }
            };

            foreach (var sup in new[] { "TME", "RS", "Farnell", "Conrad", "Elfa" })
            {
                int ip = Array.IndexOf(header, $"{sup}_Price");
                int is2 = Array.IndexOf(header, $"{sup}_Stock");
                decimal price = decimal.TryParse(row[ip], NumberStyles.Any, CultureInfo.InvariantCulture, out var p2) ? p2 : -1;
                if (price < 0) continue;
                int stock = int.TryParse(row[is2], out var s2) ? s2 : 0;
                if (stock < qty) continue;
                int del = sup == "TME" ? 2 : sup == "RS" ? 3 : sup == "Farnell" ? 5 : sup == "Conrad" ? 6 : 4;
                if (del > daysAllowed) continue;
                // ilość * cena – nie sprawdzamy budżetu per pozycja, bo budżet globalny
                allOffers.Add(new SelectedOffer
                {
                    PartNo = pn,
                    Qty = qty,
                    InternalStock = internalStock,
                    LastPurchasePrice = lastPurchase,
                    Supplier = sup,
                    Price = price,
                    DeliveryDays = del
                });
            }

            // Wyznacz min/max tylko w GUI, nie w Excel
            // Wybór oferty wg priorytetu
            IEnumerable<SelectedOffer> cand = allOffers
                .Where(o => true); // tu ewentualne filtry budżetu globalnego
            SelectedOffer chosen;
            if (_rbDeadlineFirst.Checked)
                chosen = cand.OrderBy(o => o.DeliveryDays).ThenBy(o => o.Price).First();
            else
                chosen = cand.OrderBy(o => o.Price).ThenBy(o => o.DeliveryDays).First();

            _offers.Add(chosen);

            // Wiersz GUI
            var dr = dt.NewRow();
            dr["Nr artykułu"] = chosen.PartNo;
            dr["Zapotrzebowanie"] = chosen.Qty;
            dr["Magazyn"] = chosen.InternalStock;
            dr["Ostatnia cena zakupu"] = chosen.LastPurchasePrice.ToString("0.00");
            dr["Dostawca"] = chosen.Supplier;
            dr["Cena po rabacie"] = chosen.Price.ToString("0.00");
            dr["Dostawa (dni)"] = chosen.DeliveryDays;
            dt.Rows.Add(dr);
        }

        // Podświetlający wybór
        _dgv.DataSource = dt;
        foreach (DataGridViewRow row in _dgv.Rows)
        {
            if ((string)row.Cells["Dostawca"].Value == _offers
                .First(o => o.PartNo == (string)row.Cells["Nr artykułu"].Value).Supplier)
            {
                row.DefaultCellStyle.Font = new Font(_dgv.Font, FontStyle.Bold);
            }
        }
    }

    private void ExportToExcel(
        string projectName,
        List<SelectedOffer> data,
        DateTime deadline,
        decimal budget,
        string priority)
    {
        string ts = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
        string fileName = $@"C:\EplanData\BOM_{projectName}_{ts}.xlsx";

        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("BOM");

            // Podsumowanie wiersze 1–2
            ws.Cell(1, 1).Value = "Deadline"; ws.Cell(1, 2).Value = deadline.ToString("yyyy-MM-dd");
            ws.Cell(2, 1).Value = "Budget [PLN]"; ws.Cell(2, 2).Value = budget;
            ws.Cell(1, 4).Value = "Priority"; ws.Cell(1, 5).Value = priority;

            // Terminy pierwszej i ostatniej dostawy
            int earliest = data.Min(o => o.DeliveryDays);
            int latest = data.Max(o => o.DeliveryDays);
            ws.Cell(2, 4).Value = "EarliestFinish"; ws.Cell(2, 5).Value = DateTime.Now.AddDays(earliest).ToString("yyyy-MM-dd");
            ws.Cell(1, 7).Value = "LatestFinish"; ws.Cell(1, 8).Value = DateTime.Now.AddDays(latest).ToString("yyyy-MM-dd");

            // Nagłówki row 4
            var headers = new[] { "Nr artykułu", "Zapotrzebowanie", "Magazyn", "Dostawca", "Cena po rabacie", "Dostawa(dni)" };
            for (int i = 0; i < headers.Length; i++)
                ws.Cell(4, i + 1).Value = headers[i];

            // Dane od wiersza 5
            for (int r = 0; r < data.Count; r++)
            {
                var o = data[r];
                ws.Cell(5 + r, 1).Value = o.PartNo;
                ws.Cell(5 + r, 2).Value = o.Qty;
                ws.Cell(5 + r, 3).Value = o.InternalStock;
                ws.Cell(5 + r, 4).Value = o.Supplier;
                ws.Cell(5 + r, 5).Value = o.Price;
                ws.Cell(5 + r, 6).Value = o.DeliveryDays;
            }
            // Total cost row
            decimal total = data.Sum(o => o.Price * o.Qty);
            int totalRow = 5 + data.Count;
            ws.Cell(totalRow, 4).Value = "Total";
            ws.Cell(totalRow, 5).Value = total;
            ws.Cell(totalRow, 5).Style.Font.SetBold();

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
    public void GetActionProperties(ref ActionProperties properties) { }
}
