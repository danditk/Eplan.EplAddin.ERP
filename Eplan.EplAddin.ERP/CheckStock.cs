// CheckAvailability.cs  – EPLAN API 2025  /  C# 7.3
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Eplan.EplApi.ApplicationFramework;  // IEplAction, ActionCallingContext
using Eplan.EplApi.DataModel;            // Function, ArticleReference
using Eplan.EplApi.HEServices;           // SelectionSet

namespace Eplan.EplAddin.ERP
{
    public class CheckAvailability : IEplAction
    {
        private const string CsvPath =
            @"C:\EplanData\parts_stock_full.csv";   // zmień na swoją ścieżkę

        // ---------------------- Execute ----------------------
        public bool Execute(ActionCallingContext ctx)
        {
            // 1. Pobierz zaznaczenie funkcji (Devices lub na stronie)
            StorableObject[] selected = new SelectionSet().Selection;
            if (selected == null || selected.Length == 0)
            {
                MessageBox.Show("Zaznacz funkcje (zasoby) w projekcie.", "CheckAvailability");
                return false;
            }

            // 2. Zlicz wystąpienia artykułów
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var so in selected)
            {
                if (so is Function f)
                {
                    foreach (ArticleReference ar in f.ArticleReferences)
                    {
                        if (string.IsNullOrWhiteSpace(ar.PartNr)) continue;
                        if (!counts.ContainsKey(ar.PartNr)) counts[ar.PartNr] = 0;
                        counts[ar.PartNr] += 1;          // Count w API 2025 jest zawsze >=1
                    }
                }
            }

            if (counts.Count == 0)
            {
                MessageBox.Show("Zaznaczone funkcje nie mają przypisanych artykułów.", "CheckAvailability");
                return false;
            }

            // 3. Wczytaj CSV
            if (!File.Exists(CsvPath))
            {
                MessageBox.Show($"Nie znaleziono pliku CSV:\n{CsvPath}", "CheckAvailability");
                return false;
            }

            var csv = File.ReadAllLines(CsvPath)
                          .Skip(1)
                          .Select(l => l.Split(','))     // jeśli masz średniki, zamień na Split(';')
                          .Where(c => c.Length > 3)
                          .ToDictionary(c => c[0].Trim(), c => c[3].Trim()); // kol.3 = magazyn

            // 4. Przygotuj okno z tabelą
            var form = new Form
            {
                Text = "Dostępność materiałów",
                Width = 600,
                Height = 350,
                StartPosition = FormStartPosition.CenterParent,
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
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgv.Columns.Add("PartNr", "Nr artykułu");
            dgv.Columns.Add("Needed", "Potrzebne");
            dgv.Columns.Add("Stock", "Magazyn");
            dgv.Columns.Add("Status", "✅ / ❌");

            foreach (var kv in counts)
            {
                string part = kv.Key;
                int needed = kv.Value;
                int inStock = csv.TryGetValue(part, out var stockStr) && int.TryParse(stockStr, out var s) ? s : 0;
                string status = inStock >= needed ? "✅" : "❌";

                dgv.Rows.Add(part, needed, inStock, status);
            }

            form.Controls.Add(dgv);
            form.ShowDialog();
            return true;
        }

        // ---------------------- Rejestracja -------------------
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "CheckAvailability";
            Ordinal = 30;
            return true;
        }

        // ---------------------- Właściwości akcji -------------
        public void GetActionProperties(ref ActionProperties props) { }
    }
}
