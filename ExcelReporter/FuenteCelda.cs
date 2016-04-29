using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace ExcelReporter
{
    public class FuenteCelda
    {
        public string Nombre_Fuente { get;set; }
        public short Size { get; set; }
        public bool Negrita { get; set; }
        public Color color { get; set; }
        public IFont getCellStyle(IWorkbook wb)
        {
            IFont f = wb.CreateFont();
            if (Nombre_Fuente == null)
            {
                f.FontName = "Calibri";
            }
            else
            {
                f.FontName = Nombre_Fuente;
            }

            if(Negrita == true)
                f.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

            if(Size == 0)
                f.FontHeightInPoints = 11;
            else
                f.FontHeightInPoints = Size;

            if (color != null)
                f.Color = (short)(color.R + color.G + color.B);

            return f;
        }
    }
}
