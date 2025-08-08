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
            var parts = line.Split(';');
            if (parts.Length < 5) return null;
            return new ParagraphSetting
            {
                Text = parts[0],
                Font = parts[1],
                FontSize = float.Parse(parts[2]),
                FontStyle = parts[3],
                LeftIndent = float.Parse(parts[4])
            };
        }
    }
}
