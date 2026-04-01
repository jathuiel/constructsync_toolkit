using System.ComponentModel;

namespace SetAtributesToolkit
{
    /// <summary>
    /// Item genérico com nome e estado de seleção (checkbox).
    /// Compartilhado entre SelectionInspector, MultiSetByProperty e qualquer
    /// outra view que precise de uma lista com checkboxes.
    /// </summary>
    public class CheckableItem : INotifyPropertyChanged
    {
        private string _name;
        private bool _isChecked;

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        public bool IsChecked
        {
            get => _isChecked;
            set { _isChecked = value; OnPropertyChanged(nameof(IsChecked)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
