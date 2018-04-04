/*
    Import Export Excel Test
    Requirement Npoi.Mapper
    [Convention-based mapper between strong typed object and Excel data via NPOI.]
    https://github.com/donnytian/Npoi.Mapper
*/

using Npoi.Mapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace ImportExportExcel {
    public partial class Demo : Form {
        public Demo() {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e) {
            try {
                var importPath = @"E:\TestExcel.xlsx";
                var mapper = new Mapper(importPath);
                var excelData = mapper.Take<ExcelModel>("TK_Sheet")
                                      .Select(rowInfo => rowInfo.Value).ToList();

                //excelData object is ExcelModel list object. You can do as you like.
                dataGridView1.DataSource = excelData;
                dataGridView1.Refresh();
                MessageBox.Show(importPath, "Import successful.");
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnExport_Click(object sender, EventArgs e) {
            try {
                var excelTampData = new List<ExcelModel>();
                excelTampData = Enumerable.Range(1, 10)
                                          .Select(r => new ExcelModel {
                                              ID = r,
                                              Name = "Enum-" + r,
                                              Description = "Enumerable Index " + r,
                                              Active = r / 2 == 0
                                          }).ToList();

                //excelTampData object is ExcelModel list object. You can do as you like.
                var exportPath = string.Format(@"E:\{0}.xlsx", "MyExport_"+DateTime.Now.ToString("hh_mm_sss"));
                var mapper = new Mapper();
                mapper.Save(exportPath, excelTampData, "TK_ExportSheet", overwrite: true);

                MessageBox.Show(exportPath, "Export successful.");
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
