using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Playground.Text_to_excel_converter
{
    class TextToExcelConverter
    {
        public TextToExcelConverter()
        {
            Application.EnableVisualStyles();
            Application.Run(new TxtToExcel_Form());
        }
    }
}
