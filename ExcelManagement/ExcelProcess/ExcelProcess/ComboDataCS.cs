using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProcess
{
    public class ComboDataCS
    {
        private object _comboText;
        private object _comboValue;
        public object ComboText
        {
            get { return _comboText; }
            set { _comboText = value; }
        }

        public object ComboValue
        {
            get { return _comboValue; }
            set { _comboValue = value; }
        }
        public ComboDataCS()
        {

        }
        public ComboDataCS(object text, object value)
        {
            _comboText = text;
            _comboValue = value;
        }
    }
}
