namespace PvdCalculator.Models
{
    public class ListItemModel
    {
        public ListItemModel(string text = default(string), string value = default(string))
        {
            Text = text;
            Value = value;
        }

        public string Text { get; set; }
        public string Value { get; set; }
    }
}