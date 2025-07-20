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

// Alias dla WinForms Label, by uniknąć konfliktu z EPLAN
using WinFormsLabel = System.Windows.Forms.Label;

public class AnalyzeBomWithPrices : IEplAction
{
    private DateTimePicker _dtpDeadline;
    private TextBox _txtBudget;
    private RadioButton _rbDeadlineFirst, _rbBudgetFirst;
    private DataGridView _dgv;
    private WinFormsLabel _lblMinCost, _lblMaxCost, _lblEarliest, _lblLatest, _lblTotal, _lblFinalLatest;
    private Button _btnRecalc, _btnExport;
    private List<RowOffer> _selectedOffers;

    private class RowOffer
    {
        public string CatalogNumber, EAN, ERP, Category;
        public int Qty, InternalStock;
        public decimal LastPurchasePrice;
        public string ChosenSupplier;
        public decimal ChosenPrice;
        public int ChosenDelivery;
    }

    public bool Execute(ActionCallingContext ctx)
    {
        // 1) Wczytanie CSV
        string csvFile = @"C:\EplanData\test_database_bom_szlifierka_extended.csv";
        if (!File.Exists(csvFile))
        {
            MessageBox.Show($"Nie znaleziono pliku:\n{csvFile}", "Błąd");
            return false;
        }
        var allLines = File.ReadAllLines(csvFile)
                           .Select(l => l.Split(','))
                           .ToList();
        var header = allLines[0];
        var csvRows = allLines.Skip(1)
            .Where(r => r.Length == header.Length)
            .Select(r => header.Zip(r, (h, v) => (h, v))
                               .ToDictionary(x => x.h, x => x.v, StringComparer.OrdinalIgnoreCase))
            .ToList();

        // 2) Zliczenie BOM z projektu
        var project = new SelectionSet().GetCurrentProject(true);
        if (project == null)
        {
            MessageBox.Show("Brak aktywnego projektu.", "Błąd");
            return false;
        }
        var bom = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var pg in project.Pages)
            foreach (Function f in pg.Functions)
                foreach (var ar in f.ArticleReferences)
                    if (!string.IsNullOrWhiteSpace(ar.PartNr))
                        bom[ar.PartNr] = bom.ContainsKey(ar.PartNr) ? bom[ar.PartNr] + 1 : 1;
        if (bom.Count == 0)
        {
            MessageBox.Show("BOM jest pusty.", "Info");
            return false;
        }

        // 3) Budowa GUI
        var form = new Form
        {
            Text = "Analiza BOM",
            Width = 1200,
            Height = 700,
            StartPosition = FormStartPosition.CenterScreen
        };

        var panel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 150,
            BackColor = Color.WhiteSmoke
        };
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
        panel.Controls.Add(new WinFormsLabel { Text = "Budget [PLN]:", Location = new Point(240, 10), AutoSize = true });
        _txtBudget = new TextBox { Text = "0", Location = new Point(240, 30), Width = 120 };
        panel.Controls.Add(_txtBudget);

        // Priorytet
        _rbDeadlineFirst = new RadioButton { Text = "Priorytet Deadline", Location = new Point(400, 10), AutoSize = true, Checked = true };
        _rbBudgetFirst = new RadioButton { Text = "Priorytet Budget", Location = new Point(400, 35), AutoSize = true };
        panel.Controls.Add(_rbDeadlineFirst);
        panel.Controls.Add(_rbBudgetFirst);

        // Przyciki
        _btnRecalc = new Button { Text = "PRZELICZ", Location = new Point(580, 30), Width = 100 };
        _btnExport = new Button { Text = "Eksport Excel", Location = new Point(700, 30), Width = 120 };
        panel.Controls.Add(_btnRecalc);
        panel.Controls.Add(_btnExport);

        // Podsumowanie
        _lblMinCost = new WinFormsLabel { Location = new Point(240, 60), AutoSize = true };
        _lblMaxCost = new WinFormsLabel { Location = new Point(240, 85), AutoSize = true };
        _lblEarliest = new WinFormsLabel { Location = new Point(10, 60), AutoSize = true };
        _lblLatest = new WinFormsLabel { Location = new Point(10, 85), AutoSize = true };
        _lblTotal = new WinFormsLabel { Location = new Point(900, 60), AutoSize = true, Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold) };
        _lblFinalLatest = new WinFormsLabel { Location = new Point(900, 90), AutoSize = true, Font = new Font(SystemFonts.DefaultFont, FontStyle.Italic) };
        panel.Controls.AddRange(new Control[] { _lblMinCost, _lblMaxCost, _lblEarliest, _lblLatest, _lblTotal, _lblFinalLatest });

        // DataGridView – poniżej panelu
        _dgv = new DataGridView
        {
            Location = new Point(0, panel.Height),
            Width = form.ClientSize.Width,
            Height = form.ClientSize.Height - panel.Height,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            AllowUserToAddRows = false
        };
        form.Controls.Add(_dgv);

        // Eventy
        _btnRecalc.Click += (s, e) =>
        {
            if (!decimal.TryParse(_txtBudget.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var b))
            {
                MessageBox.Show("Błędny budget", "Błąd");
                return;
            }
            Recalculate(bom, csvRows, header, b, _dtpDeadline.Value.Date);
        };
        _btnExport.Click += (s, e) =>
        {
            ExportToExcel(
                project.ProjectName,
                _selectedOffers,
                _dtpDeadline.Value.Date,
                decimal.TryParse(_txtBudget.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var b2) ? b2 : 0m,
                _rbDeadlineFirst.Checked ? "Deadline" : "Budget");
        };

        // Automatyczne pierwsze przeliczenie po wczytaniu formy
        form.Load += (s, e) => _btnRecalc.PerformClick();

        form.ShowDialog();
        return true;
    }

    private void Recalculate(
        Dictionary<string, int> bom,
        List<Dictionary<string, string>> csvRows,
        string[] header,
        decimal budget,
        DateTime deadline)
    {
        _selectedOffers = new List<RowOffer>();
        var dt = new DataTable();

        // GUI: wszystkie hurtownie + kolumny wyborcze
        foreach (var col in new[] { "CatalogNumber", "EAN_Code", "ERP_Code", "Category", "Qty", "Stock", "LastPurchasePrice" })
            dt.Columns.Add(col);
        foreach (var sup in new[] { "TME", "RS", "Farnell", "Conrad", "Elfa" })
        {
            dt.Columns.Add($"{sup}_Price");
            dt.Columns.Add($"{sup}_Discount");
            dt.Columns.Add($"{sup}_Stock");
            dt.Columns.Add($"{sup}_Delivery");
        }
        dt.Columns.Add("ChosenSupplier");
        dt.Columns.Add("ChosenPrice");
        dt.Columns.Add("ChosenDelivery");

        int daysAllowed = (deadline - DateTime.Now.Date).Days;
        decimal sumMin = 0, sumMax = 0, sumTotal = 0;
        int earliest = 0, latest = 0;

        foreach (var kv in bom)
        {
            string pn = kv.Key;
            int qty = kv.Value;
            var rowCsv = csvRows.FirstOrDefault(r => r["PartNo"] == pn);
            if (rowCsv == null) continue;

            // Zbierz oferty
            var offers = new List<RowOffer>();
            int stock = int.TryParse(rowCsv["InternalStock"], out var st) ? st : 0;
            decimal lp = decimal.TryParse(rowCsv["LastPrice"], NumberStyles.Any, CultureInfo.InvariantCulture, out var lpVal) ? lpVal : 0m;

            // Magazyn
            offers.Add(new RowOffer
            {
                CatalogNumber = rowCsv["CatalogNumber"],
                EAN = rowCsv["EAN_Code"],
                ERP = rowCsv["ERP_Code"],
                Category = rowCsv["Category"],
                Qty = qty,
                InternalStock = stock,
                LastPurchasePrice = lp,
                ChosenSupplier = "Magazyn",
                ChosenPrice = lp,
                ChosenDelivery = 0
            });

            // Hurtownie
            foreach (var sup in new[] { "TME", "RS", "Farnell", "Conrad", "Elfa" })
            {
                if (!decimal.TryParse(rowCsv[$"{sup}_Price"], NumberStyles.Any, CultureInfo.InvariantCulture, out var price)) continue;
                if (!decimal.TryParse(rowCsv[$"{sup}_Discount"], NumberStyles.Any, CultureInfo.InvariantCulture, out var disc)) disc = 0m;
                if (!int.TryParse(rowCsv[$"{sup}_Stock"], out var supStock) || supStock < qty) continue;
                int del = sup == "TME" ? 2 : sup == "RS" ? 3 : sup == "Farnell" ? 5 : sup == "Conrad" ? 6 : 4;
                if (del > daysAllowed) continue;

                decimal fp = Math.Round(price * (1 - disc / 100m), 2);
                offers.Add(new RowOffer
                {
                    CatalogNumber = rowCsv["CatalogNumber"],
                    EAN = rowCsv["EAN_Code"],
                    ERP = rowCsv["ERP_Code"],
                    Category = rowCsv["Category"],
                    Qty = qty,
                    InternalStock = stock,
                    LastPurchasePrice = lp,
                    ChosenSupplier = sup,
                    ChosenPrice = fp,
                    ChosenDelivery = del
                });
            }

            // Zakresy
            sumMin += offers.Min(o => o.ChosenPrice * qty);
            sumMax += offers.Max(o => o.ChosenPrice * qty);
            earliest = Math.Max(earliest, offers.Min(o => o.ChosenDelivery));
            latest = Math.Max(latest, offers.Max(o => o.ChosenDelivery));

            // Wybór wg priorytetu
            var chosen = _rbDeadlineFirst.Checked
                ? offers.OrderBy(o => o.ChosenDelivery).ThenBy(o => o.ChosenPrice).First()
                : offers.OrderBy(o => o.ChosenPrice).ThenBy(o => o.ChosenDelivery).First();

            _selectedOffers.Add(chosen);
            sumTotal += chosen.ChosenPrice * qty;

            // Dodaj wiersz do GUI
            var dr = dt.NewRow();
            dr["CatalogNumber"] = chosen.CatalogNumber;
            dr["EAN_Code"] = chosen.EAN;
            dr["ERP_Code"] = chosen.ERP;
            dr["Category"] = chosen.Category;
            dr["Qty"] = chosen.Qty;
            dr["Stock"] = chosen.InternalStock;
            dr["LastPurchasePrice"] = chosen.LastPurchasePrice.ToString("0.00");
            foreach (var sup in new[] { "TME", "RS", "Farnell", "Conrad", "Elfa" })
            {
                dr[$"{sup}_Price"] = rowCsv[$"{sup}_Price"];
                dr[$"{sup}_Discount"] = rowCsv[$"{sup}_Discount"] + "%";
                dr[$"{sup}_Stock"] = rowCsv[$"{sup}_Stock"];
                dr[$"{sup}_Delivery"] = new[] { 2, 3, 5, 6, 4 }[Array.IndexOf(new[] { "TME", "RS", "Farnell", "Conrad", "Elfa" }, sup)];
            }
            dr["ChosenSupplier"] = chosen.ChosenSupplier;
            dr["ChosenPrice"] = chosen.ChosenPrice.ToString("0.00");
            dr["ChosenDelivery"] = chosen.ChosenDelivery;
            dt.Rows.Add(dr);
        }

        // Aktualizuj summary
        _lblMinCost.Text = $"MinCost:  {sumMin:0.00}";
        _lblMaxCost.Text = $"MaxCost:  {sumMax:0.00}";
        _lblEarliest.Text = $"Earliest: {DateTime.Now.AddDays(earliest):yyyy-MM-dd} (+{earliest}d)";
        _lblLatest.Text = $"Latest:   {DateTime.Now.AddDays(latest):yyyy-MM-dd} (+{latest}d)";
        _lblTotal.Text = $"Total:    {sumTotal:0.00}";

        // FinalLatest
        if (_rbDeadlineFirst.Checked && sumTotal > sumMin)
        {
            int newEarliest = _selectedOffers.Min(o => o.ChosenDelivery);
            if (newEarliest < latest)
                _lblFinalLatest.Text = $"FinalLatest: {DateTime.Now.AddDays(newEarliest):yyyy-MM-dd} (+{newEarliest}d)";
            else
                _lblFinalLatest.Text = "";
        }
        else
            _lblFinalLatest.Text = "";

        // Bind i pogrubienie
        _dgv.DataSource = dt;
        foreach (DataGridViewRow rw in _dgv.Rows)
        {
            var sup = rw.Cells["ChosenSupplier"].Value?.ToString();
            if (!string.IsNullOrEmpty(sup))
            {
                rw.Cells["ChosenSupplier"].Style.Font = new Font(_dgv.Font, FontStyle.Bold);
                rw.Cells["ChosenPrice"].Style.Font = new Font(_dgv.Font, FontStyle.Bold);
                rw.Cells["ChosenDelivery"].Style.Font = new Font(_dgv.Font, FontStyle.Bold);
            }
        }
    }

    private void ExportToExcel(
        string projectName,
        List<RowOffer> data,
        DateTime deadline,
        decimal budget,
        string priority)
    {
        string ts = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
        string file = $@"C:\EplanData\BOM_{projectName}_{ts}.xlsx";
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("BOM");
            // Podsumowanie 1–3
            ws.Cell(1, 1).Value = "Deadline"; ws.Cell(1, 2).Value = deadline.ToString("yyyy-MM-dd");
            ws.Cell(2, 1).Value = "Budget"; ws.Cell(2, 2).Value = budget;
            ws.Cell(1, 4).Value = "Priority"; ws.Cell(1, 5).Value = priority;
            ws.Cell(1, 7).Value = "Total"; ws.Cell(1, 8).Value = data.Sum(o => o.ChosenPrice * o.Qty);
            int e = data.Min(o => o.ChosenDelivery), l = data.Max(o => o.ChosenDelivery);
            ws.Cell(2, 4).Value = "EarliestFinish"; ws.Cell(2, 5).Value = $"{DateTime.Now.AddDays(e):yyyy-MM-dd} (+{e}d)";
            ws.Cell(2, 7).Value = "LatestFinish"; ws.Cell(2, 8).Value = $"{DateTime.Now.AddDays(l):yyyy-MM-dd} (+{l}d)";

            // Nagłówki row5
            var hdr = new[] { "CatalogNumber","EAN_Code","ERP_Code","Category","Qty","Stock","LastPurchasePrice",
                              "ChosenSupplier","ChosenPrice","ChosenDelivery" };
            for (int i = 0; i < hdr.Length; i++) ws.Cell(5, i + 1).Value = hdr[i];

            // Dane row6+
            for (int r = 0; r < data.Count; r++)
            {
                var o = data[r];
                ws.Cell(6 + r, 1).Value = o.CatalogNumber;
                ws.Cell(6 + r, 2).Value = o.EAN;
                ws.Cell(6 + r, 3).Value = o.ERP;
                ws.Cell(6 + r, 4).Value = o.Category;
                ws.Cell(6 + r, 5).Value = o.Qty;
                ws.Cell(6 + r, 6).Value = o.InternalStock;
                ws.Cell(6 + r, 7).Value = o.LastPurchasePrice;
                ws.Cell(6 + r, 8).Value = o.ChosenSupplier;
                ws.Cell(6 + r, 9).Value = o.ChosenPrice;
                ws.Cell(6 + r, 10).Value = o.ChosenDelivery;
            }

            wb.SaveAs(file);
        }
        MessageBox.Show($"Zapisano: {file}", "OK");
    }

    public bool OnRegister(ref string Name, ref int Ordinal)
    {
        Name = "AnalyzeBomWithPrices";
        Ordinal = 50;
        return true;
    }
    public void GetActionProperties(ref ActionProperties props) { }
}
