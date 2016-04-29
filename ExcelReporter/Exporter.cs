using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;


namespace ExcelReporter
{
    public class Exporter
    {
        private IWorkbook _workbook;
        private IFont _hfont=null;
        private IFont _rfont=null;
        private ICellStyle _hEstilo = null;
        private ICellStyle _rEstilo = null;

        public Exporter()
        {
            this._workbook= new XSSFWorkbook();
        }

        private IFont headerFont()
        {
            if (_hfont == null)
            {
                IFont f = this._workbook.CreateFont();

                f.FontName = "Calibri";
                f.FontHeightInPoints = 11;
                f.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                f.Color = IndexedColors.White.Index;
                return f;
            }
            else
            {
                return _hfont;
            }
            
        }
        private IFont rowFont()
        {
            if (_rfont == null)
            {
                IFont f = this._workbook.CreateFont();

                f.FontName = "Calibri";
                f.FontHeightInPoints = 11;
                //f.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

                return f;
            }
            else
            {
                return _rfont;
            }
        }

        private ICell crearCeldaHeader(IRow row,int n)
        {
            ICell cell = row.CreateCell(n);

            if (this._hEstilo == null)
            {
                this._hEstilo = this._workbook.CreateCellStyle();
                this._hEstilo.SetFont(headerFont());
                this._hEstilo.Alignment = HorizontalAlignment.Center;
                this._hEstilo.FillForegroundColor = IndexedColors.Black.Index;
                this._hEstilo.FillBackgroundColor = IndexedColors.White.Index;
                this._hEstilo.BorderBottom = BorderStyle.Thin;
                this._hEstilo.BorderTop = BorderStyle.Thin;
                this._hEstilo.BorderLeft = BorderStyle.Thin;
                this._hEstilo.BorderRight = BorderStyle.Thin;

                this._hEstilo.FillPattern = FillPattern.SolidForeground;
            }
           
            cell.CellStyle = this._hEstilo;
            return cell;
        }
        private ICell crearCeldaRow(IRow row,int n)
        {
            ICell cell = row.CreateCell(n);

            if (this._rEstilo == null)
            {
                this._rEstilo = this._workbook.CreateCellStyle();

                this._rEstilo.SetFont(rowFont());

                this._rEstilo.Alignment = HorizontalAlignment.Center;

                this._rEstilo.BorderBottom = BorderStyle.Thin;
                this._rEstilo.BorderTop = BorderStyle.Thin;
                this._rEstilo.BorderLeft = BorderStyle.Thin;
                this._rEstilo.BorderRight = BorderStyle.Thin;
            }
            cell.CellStyle = this._rEstilo;
            return cell;
        }
        private void addHoja(DataTable dt,string nombre)
        {
            ISheet sheet = _workbook.CreateSheet(nombre);
            
            /*columnas*/
            IRow row = sheet.CreateRow(0);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = crearCeldaHeader(row, j);
                cell.SetCellValue(dt.Columns[j].ColumnName);
            }

            /*rows*/
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                row = sheet.CreateRow(i+1);
                
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell= crearCeldaRow(row,j);
                    
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }
        }

        public void AgregarWorkSheet(DataTable dt, string nombre)
        {
            addHoja(dt, nombre);
        }

        public void CrearArchivoExcelDialog()
        {
            System.Windows.Forms.SaveFileDialog ofd = new System.Windows.Forms.SaveFileDialog();
            //ofd.Filter = "Excel 97-2003 (*.xls)|*.xls|Excel 2007+(*.xlsx)|*.xlsx";
            ofd.Filter = "Excel 2007+(*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                CrearArchivoExcel(ofd.FileName);
            }
        }

        /*public void CrearArchivoExcel(string file)
        {
            using (FileStream stream = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(stream);
            }
        }*/

        public void CrearArchivoExcel(string file)
        {
            using (FileStream stream = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(stream);
            }
        }

        public void setFuenteEncabezado(FuenteCelda fuente)
        {
            this._hfont = fuente.getCellStyle(this._workbook);
        }
        public void setFuenteFilas(FuenteCelda fuente)
        {
            this._rfont = fuente.getCellStyle(this._workbook);
        }

        public void setEstiloEncabezado(EstiloCelda estilo)
        {
            this._hEstilo = estilo.getCellStyle(this._workbook);
        }
        public void setEstiloFilas(EstiloCelda estilo)
        {
            this._rEstilo = estilo.getCellStyle(this._workbook);
        }

    }
}
