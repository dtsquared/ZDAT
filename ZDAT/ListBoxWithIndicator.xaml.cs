using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ZDAT
{
    /// <summary>
    /// Interaction logic for ListBoxWithIndicator.xaml
    /// </summary>
    public partial class ListBoxWithIndicator : UserControl
    {
        #region Data

        const int INDICATOR_HEIGHT = 16;
        readonly ObservableCollection<double> _indicatorOffsets;

        #endregion // Data

        #region Constructor
        public ListBoxWithIndicator()
        {
            InitializeComponent();

            // Set up the list of selection indicator offsets
            // as the data source for the ItemsControl.
            _indicatorOffsets = new ObservableCollection<double>();
            _indicatorList.ItemsSource = _indicatorOffsets;
            
            // Move the indicators when the set of 
            // selected items is modified.
            _listBox.SelectionChanged += delegate
            {
                this.UpdateIndicators();
            };

            // Move the indicators when the ListBox's
            // ScrollViewer is scrolled.
            _listBox.AddHandler(
                ScrollViewer.ScrollChangedEvent,
                new ScrollChangedEventHandler(delegate
                {
                    this.UpdateIndicators();
                }));
        }
        #endregion // Constructor

        #region Public Properties

        #region IndicatorBrush

        /// <summary>
        /// Gets/sets the Brush used to fill the selection indicator.
        /// This is a dependency property.
        /// </summary>
        public Brush IndicatorBrush
        {
            get { return (Brush)GetValue(IndicatorBrushProperty); }
            set { SetValue(IndicatorBrushProperty, value); }
        }

        /// <summary>
        /// Represents the IndicatorBrush property.
        /// This field is read-only.
        /// </summary>
        public static readonly DependencyProperty IndicatorBrushProperty =
            DependencyProperty.Register(
            "IndicatorBrush",
            typeof(Brush),
            typeof(ListBoxWithIndicator),
            new PropertyMetadata(SystemColors.HighlightBrush));

        #endregion // IndicatorBrush

        #region ListBox

        /// <summary>
        /// Returns the ListBox.
        /// </summary>
        public ListBox ListBox
        {
            get { return _listBox; }
        }

        #endregion // ListBox

        #region ListBoxStyle

        /// <summary>
        /// Gets/sets the Style applied to the ListBox.
        /// This is a dependency property.
        /// </summary>
        public Style ListBoxStyle
        {
            get { return (Style)GetValue(ListBoxStyleProperty); }
            set { SetValue(ListBoxStyleProperty, value); }
        }

        /// <summary>
        /// Represents the ListBoxStyle property.
        /// This field is read-only.
        /// </summary>
        public static readonly DependencyProperty ListBoxStyleProperty =
            DependencyProperty.Register(
            "ListBoxStyle",
            typeof(Style),
            typeof(ListBoxWithIndicator));

        #endregion // ListBoxStyle

        #endregion // Public Properties

        #region Private Helpers

        void UpdateIndicators()
        {
            if (_indicatorOffsets.Count > 0)
                _indicatorOffsets.Clear();

            if (_listBox.SelectedItems.Count == 0)
                return;

            // Make sure the ListBoxItems have been generated.
            ItemContainerGenerator gen = _listBox.ItemContainerGenerator;
            if (gen.Status != GeneratorStatus.ContainersGenerated)
                return;

            // Determine the selection indicator vertical offset for
            // every selected item in the ListBox.
            foreach (object selectedItem in _listBox.SelectedItems)
            {
                // Get the selected ListBoxItem.
                ListBoxItem lbItem = gen.ContainerFromItem(selectedItem) as ListBoxItem;

                // Just in case...
                if (lbItem == null)
                    continue;

                // Determine the selected item's location relative to the 
                // ListBox in which it is contained.
                GeneralTransform trans = lbItem.TransformToAncestor(_listBox);
                Point location = trans.Transform(new Point(0, 0));

                // Calculate and store the selection indicator offset.
                // NOTE: Adding a value to '_indicatorOffsets' automatically
                // notifies the UI because it is an ObservableCollection<double>.
                double offset = location.Y + (lbItem.ActualHeight / 2) - (INDICATOR_HEIGHT / 2);
                _indicatorOffsets.Add(offset);
            }
        }

        #endregion // Private Helpers
    }
}
