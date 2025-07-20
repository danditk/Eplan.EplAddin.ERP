using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Eplan.EplApi.Base;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;

public class AnalyzeBomWithPrices : IEplAction
{
    private TextBox deadlineBox;
    private TextBox budgetBox;

    public bool Execute(ActionCallingContext ctx)
    {
        // 1) Wczytanie CSV
        string csvPath = @"C:\EplanData\test_database_bom_szlifierka_extended.csv";
        if (!File.Exists(csvPath))
        {
            MessageBox.Show($"Nie znaleziono pliku: {csvPath}");
            return false;
        }
        var lines = File.ReadAllLines(csvPath);
        var header = lines[0].Split(',');
        var rows = lines.Skip(1)
                        .Select(l => l.Split(','))
                        .Where(c => c.Length == header.Length)
                        .ToList();

        // 2) Mapa PartNo → row[]
        var partRows = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        int idxPartNo = Array.IndexOf(header, "PartNo");
        foreach (var row in rows)
        {
            var pn = row[idxPartNo];
            if (!string.IsNullOrWhiteSpace(pn))
                partRows[pn] = row;
        }

        // 3) Pobranie projektu i BOM
        var project = new SelectionSet().GetCurrentProject(true);
        if (project == null)
        {
            MessageBox.Show("Brak aktywnego projektu.");
            return false;
        }
        var bom = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in project.Pages)
        {
            foreach (Function func in page.Functions)
            {
                foreach (var ar in func.ArticleReferences)
                {
                    string pn = ar.PartNr;
                    if (string.IsNullOrWhiteSpace(pn))
                        continue;
                    if (bom.ContainsKey(pn)) bom[pn]++;
                    else bom[pn] = 1;
                }
            }
        }
        if (bom.Count == 0)
        {
            MessageBox.Show("Brak artykułów w BOMie.");
            return false;
        }

        // 4) Przygotowanie DataTable
        var dt = new DataTable();
        dt.Columns.Add("PartNo");
        dt.Columns.Add("Qty");
        dt.Columns.Add("InternalStock");
        dt.Columns.Add("LastSupplier");
        dt.Columns.Add("LastPrice");

        string[] suppliers = { "TME", "RS", "Farnell", "Conrad", "Elfa" };
        foreach (var sup in suppliers)
        {
            dt.Columns.Add($"{sup}_Price");
            dt.Columns.Add($"{sup}_Discount");
            dt.Columns.Add($"{sup}_Stock");
        }

        // 5) Wypełnianie tabeli
        foreach (var kv in bom)
        {
            var pn = kv.Key;
            var qty = kv.Value;
            if (!partRows.TryGetValue(pn, out var row))
                continue;

            var dr = dt.NewRow();
            dr["PartNo"] = pn;
            dr["Qty"] = qty;

            int idxInternal = Array.IndexOf(header, "InternalStock");
            int idxLastSup = Array.IndexOf(header, "LastSupplier");
            int idxLastPrice = Array.IndexOf(header, "LastPrice");
            dr["InternalStock"] = row[idxInternal];
            dr["LastSupplier"] = row[idxLastSup];
            dr["LastPrice"] = row[idxLastPrice];

            foreach (var sup in suppliers)
            {
                int idxPrice = Array.IndexOf(header, $"{sup}_Price");
                int idxDiscount = Array.IndexOf(header, $"{sup}_Discount");
                int idxStock = Array.IndexOf(header, $"{sup}_Stock");

                decimal price = decimal.TryParse(row[idxPrice], NumberStyles.Any, CultureInfo.InvariantCulture, out var p) ? p : 0m;
                decimal disc = decimal.TryParse(row[idxDiscount], NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0m;
                int st = int.TryParse(row[idxStock], out var s2) ? s2 : 0;

                dr[$"{sup}_Price"] = price.ToString("0.00");
                dr[$"{sup}_Discount"] = disc.ToString("0.##") + " %";
                dr[$"{sup}_Stock"] = st;
            }

            dt.Rows.Add(dr);
        }

        // 6) GUI
        var form = new Form { Text = "Analiza BOM z hurtowniami", Width = 1200, Height = 600 };
        var grid = new DataGridView
        {
            DataSource = dt,
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        };
        form.Controls.Add(grid);

        deadlineBox = new TextBox { Text = "Deadline YYYY-MM-DD", ForeColor = Color.Gray, Top = 5, Left = 5, Width = 200 };
        budgetBox = new TextBox { Text = "Budżet [PLN]", ForeColor = Color.Gray, Top = 35, Left = 5, Width = 200 };
        form.Controls.Add(deadlineBox);
        form.Controls.Add(budgetBox);

        deadlineBox.GotFocus += (s, e) =>
        {
            if (deadlineBox.ForeColor == Color.Gray)
            {
                deadlineBox.Text = "";
                deadlineBox.ForeColor = Color.Black;
            }
        };
        budgetBox.GotFocus += (s, e) =>
        {
            if (budgetBox.ForeColor == Color.Gray)
            {
                budgetBox.Text = "";
                budgetBox.ForeColor = Color.Black;
            }
        };

        var exportBtn = new Button { Text = "Eksportuj do Excela", Top = 65, Left = 5, Width = 200 };
        form.Controls.Add(exportBtn);

        // 7) Eksport z ClosedXML – konwertujemy wartości na stringi
        exportBtn.Click += (s, e) =>
        {
            string deadline = deadlineBox.Text.Trim();
            decimal.TryParse(budgetBox.Text.Trim(), out var budget);

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("BOM");
                ws.Cell(1, 1).Value = "Deadline:";
                ws.Cell(1, 2).Value = deadline;
                ws.Cell(2, 1).Value = "Budget:";
                ws.Cell(2, 2).Value = budget;

                // nagłówki
                for (int c = 0; c < dt.Columns.Count; c++)
                    ws.Cell(4, c + 1).Value = dt.Columns[c].ColumnName;

                // dane jako string
                for (int r = 0; r < dt.Rows.Count; r++)
                    for (int c = 0; c < dt.Columns.Count; c++)
                        ws.Cell(r + 5, c + 1).Value = dt.Rows[r][c]?.ToString() ?? "";

                string fileName = $@"C:\EplanData\BOM_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx";
                wb.SaveAs(fileName);
                MessageBox.Show($"Zapisano: {fileName}");
            }
        };

        form.ShowDialog();
        return true;
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
