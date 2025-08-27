using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TitlePageGenerator
{
    public class ParagraphSetting
    {
        public string Text { get; set; }
        public string Font { get; set; }
        public float FontSize { get; set; }
        public string FontStyle { get; set; }
        public float LeftIndent { get; set; }

        public override string ToString()
        {
            return $"{Text};{Font};{FontSize};{FontStyle};{LeftIndent}";
        }

        public static ParagraphSetting FromString(string line)
        {
            string[] parts = line.Split(';');
            ParagraphSetting result;

            if (parts.Length < 5)
            {
                result = null;
            }
            else
            {
                result = new ParagraphSetting
                {
                    Text = parts[0],
                    Font = parts[1],
                    FontSize = float.Parse(parts[2]),
                    FontStyle = parts[3],
                    LeftIndent = float.Parse(parts[4])
                };
            }

            return result;
        }
    }
}