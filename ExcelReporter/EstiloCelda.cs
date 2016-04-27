using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace ExcelReporter
{
    class EstiloCelda
    {
        public enum Alineacion_Horizontal_Texto {
            centrado,
            izquierda,
            derecha,
            justificado
        }
        public enum Alineacion_Vertical_Texto
        {
            centro,
            arriba,
            abajo
        }
        public enum Tipo_Relleno_Celda
        { 
            solido,
            cuadrados
        }

        public bool cuadricula { get; set; }
        public Alineacion_Horizontal_Texto Alineacion_TextoH { get; set; }
        public Alineacion_Vertical_Texto Alineacion_TextoV { get; set; }
        public Tipo_Relleno_Celda Tipo_Relleno { get; set; }

        public Color color_fondo { get; set; } 


        public ICellStyle getCellStyle(IWorkbook wb)
        {
            ICellStyle style = wb.CreateCellStyle();

            switch (Alineacion_TextoH)
            { 
                case Alineacion_Horizontal_Texto.centrado:
                    style.Alignment = HorizontalAlignment.Center;
                break;
                case Alineacion_Horizontal_Texto.izquierda:
                    style.Alignment = HorizontalAlignment.Left;
                break;
                case Alineacion_Horizontal_Texto.derecha:
                    style.Alignment = HorizontalAlignment.Right;
                break;
                case Alineacion_Horizontal_Texto.justificado:
                    style.Alignment = HorizontalAlignment.Justify;
                break;
            }
            switch (Alineacion_TextoV)
            { 
                case Alineacion_Vertical_Texto.centro:
                    style.VerticalAlignment = VerticalAlignment.Center;
                break;
                case Alineacion_Vertical_Texto.arriba:
                    style.VerticalAlignment = VerticalAlignment.Top;
                break;
                case Alineacion_Vertical_Texto.abajo:
                    style.VerticalAlignment = VerticalAlignment.Bottom;
                break;
            }
            if (color_fondo != null)
            {
                style.FillForegroundColor = (short)(color_fondo.R+color_fondo.G+color_fondo.B);//IndexedColors.Black.Index;
                //style.FillBackgroundColor = IndexedColors.White.Index;
            }

            if (cuadricula == true)
            {
                style.BorderBottom = BorderStyle.Thin;
                style.BorderTop = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
            }
            switch (Tipo_Relleno)
            { 
                case Tipo_Relleno_Celda.solido:
                    style.FillPattern = FillPattern.SolidForeground;
                break;
                case Tipo_Relleno_Celda.cuadrados:
                    style.FillPattern = FillPattern.Squares;
                break;
            }
        
            return style;
        }
    }
}
