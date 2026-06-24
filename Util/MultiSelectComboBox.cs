using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace SQLManage.Util
{
    /// <summary>
    /// 多选下拉控件 - 支持多选工站
    /// </summary>
    public class MultiSelectComboBox : Control
    {
        private Border _rootBorder;
        private TextBlock _displayTextBlock;
        private ToggleButton _toggleButton;
        private Popup _popup;
        private ListBox _listBox;

        #region SelectedItems 依赖属性

        public static readonly DependencyProperty SelectedItemsProperty =
            DependencyProperty.Register(nameof(SelectedItems), typeof(ObservableCollection<string>),
                typeof(MultiSelectComboBox), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnSelectedItemsChanged));

        public ObservableCollection<string> SelectedItems
        {
            get => (ObservableCollection<string>)GetValue(SelectedItemsProperty);
            set => SetValue(SelectedItemsProperty, value);
        }

        #endregion

        #region ItemsSource 依赖属性

        public static readonly DependencyProperty ItemsSourceProperty =
            DependencyProperty.Register(nameof(ItemsSource), typeof(IEnumerable),
                typeof(MultiSelectComboBox), new PropertyMetadata(null, OnItemsSourceChanged));

        public IEnumerable ItemsSource
        {
            get => (IEnumerable)GetValue(ItemsSourceProperty);
            set => SetValue(ItemsSourceProperty, value);
        }

        #endregion

        public MultiSelectComboBox()
        {
            var col = new ObservableCollection<string>();
            col.CollectionChanged += OnSelectedCollectionChanged;
            SetCurrentValue(SelectedItemsProperty, col);

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // 显示文本
            _displayTextBlock = new TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(8, 3, 28, 3),
                TextTrimming = TextTrimming.CharacterEllipsis,
                Text = "全部",
                FontSize = 13
            };

            // 下拉箭头
            var arrow = new TextBlock
            {
                Text = "▼",
                FontSize = 9,
                Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 6, 0)
            };

            var grid = new Grid();
            grid.Children.Add(_displayTextBlock);
            grid.Children.Add(arrow);

            // ToggleButton 包裹整个显示区域，控制 Popup 开关
            _toggleButton = new ToggleButton
            {
                Content = grid,
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xBB, 0xBB, 0xBB)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(0),
                Focusable = true,
                Cursor = Cursors.Hand,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                VerticalContentAlignment = VerticalAlignment.Stretch,
                SnapsToDevicePixels = true
            };

            // 外层 Border 用于布局
            _rootBorder = new Border
            {
                Child = _toggleButton,
                SnapsToDevicePixels = true
            };

            // Popup 下拉面板
            _popup = new Popup
            {
                PlacementTarget = _rootBorder,
                Placement = PlacementMode.Bottom,
                StaysOpen = false,
                AllowsTransparency = true,
                PopupAnimation = PopupAnimation.Fade,
                Focusable = false
            };

            // 双向绑定 ToggleButton.IsChecked <-> Popup.IsOpen
            var isOpenBinding = new Binding(nameof(ToggleButton.IsChecked))
            {
                Source = _toggleButton,
                Mode = BindingMode.TwoWay
            };
            _popup.SetBinding(Popup.IsOpenProperty, isOpenBinding);

            var popupBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xAA, 0xAA, 0xAA)),
                BorderThickness = new Thickness(1),
                MaxHeight = 200,
                SnapsToDevicePixels = true,
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    BlurRadius = 4,
                    ShadowDepth = 1,
                    Opacity = 0.3,
                    Color = Colors.Black
                }
            };

            _listBox = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Margin = new Thickness(4)
            };
            _listBox.ItemContainerGenerator.StatusChanged += OnListBoxGenerated;

            // 阻止 Popup 内点击冒泡到外部导致 Popup 关闭
            // 但要让 CheckBox 自身的点击正常工作
            _listBox.AddHandler(
                ListBox.PreviewMouseLeftButtonDownEvent,
                new MouseButtonEventHandler((s, ev) =>
                {
                    // 检查点击源是否是 CheckBox 自身或其内部
                    var source = ev.OriginalSource as DependencyObject;
                    bool isCheckBox = false;
                    while (source != null)
                    {
                        if (source is CheckBox)
                        {
                            isCheckBox = true;
                            break;
                        }
                        source = VisualTreeHelper.GetParent(source);
                    }

                    // 如果点击的不是 CheckBox，阻止冒泡防止关闭 Popup
                    if (!isCheckBox)
                    {
                        ev.Handled = true;
                    }
                }),
                true); // handledEventsToo = true 确保即使事件已被处理也能收到

            var scrollViewer = new ScrollViewer
            {
                Content = _listBox,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
            popupBorder.Child = scrollViewer;
            _popup.Child = popupBorder;

            // ListBox 的 ItemsSource 绑定
            _listBox.SetBinding(ListBox.ItemsSourceProperty, new Binding(nameof(ItemsSource))
            {
                Source = this,
                Mode = BindingMode.OneWay
            });

            // Popup 打开时同步宽度和 CheckBox 状态
            _popup.Opened += (s, e) =>
            {
                _popup.Width = this.ActualWidth > 0 ? this.ActualWidth : 120;
                UpdateCheckBoxStates();
            };

            // 将 Border 添加为控件自身的可视化子元素
            AddLogicalChild(_rootBorder);
            AddVisualChild(_rootBorder);
        }

        private void OnListBoxGenerated(object sender, EventArgs e)
        {
            if (_listBox.ItemContainerGenerator.Status == GeneratorStatus.ContainersGenerated)
            {
                UpdateCheckBoxStates();
            }
        }

        private void UpdateCheckBoxStates()
        {
            if (_listBox == null || SelectedItems == null) return;

            for (int i = 0; i < _listBox.Items.Count; i++)
            {
                var container = _listBox.ItemContainerGenerator.ContainerFromIndex(i) as ListBoxItem;
                if (container == null) continue;

                var checkBox = container.Content as CheckBox;
                if (checkBox == null)
                {
                    checkBox = CreateCheckBox(_listBox.Items[i]?.ToString());
                    container.Content = checkBox;
                }

                string itemStr = _listBox.Items[i]?.ToString();
                checkBox.IsChecked = SelectedItems.Contains(itemStr);
            }
        }

        private CheckBox CreateCheckBox(string item)
        {
            var cb = new CheckBox
            {
                Content = item,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 3, 4, 3),
                Tag = item,
                FontSize = 13
            };

            cb.Checked += OnCheckBoxChanged;
            cb.Unchecked += OnCheckBoxChanged;

            return cb;
        }

        private void OnCheckBoxChanged(object sender, RoutedEventArgs e)
        {
            var cb = sender as CheckBox;
            if (cb == null) return;

            string item = cb.Tag?.ToString();
            if (item == null) return;

            if (_updatingCheckBoxes) return;

            _updatingCheckBoxes = true;
            try
            {
                if (item == "全部")
                {
                    if (cb.IsChecked == true)
                    {
                        SelectedItems.Clear();
                        SelectedItems.Add("全部");
                    }
                    else
                    {
                        SelectedItems.Remove("全部");
                    }
                }
                else
                {
                    if (cb.IsChecked == true)
                    {
                        SelectedItems.Remove("全部");
                        if (!SelectedItems.Contains(item))
                            SelectedItems.Add(item);
                    }
                    else
                    {
                        SelectedItems.Remove(item);
                    }

                    if (SelectedItems.Count == 0)
                    {
                        SelectedItems.Add("全部");
                    }
                }

                UpdateDisplayText();
                UpdateCheckBoxStates();
            }
            finally
            {
                _updatingCheckBoxes = false;
            }
        }

        private bool _updatingCheckBoxes;

        private void UpdateDisplayText()
        {
            if (_displayTextBlock == null) return;

            if (SelectedItems == null || SelectedItems.Count == 0)
                _displayTextBlock.Text = "全部";
            else if (SelectedItems.Contains("全部"))
                _displayTextBlock.Text = "全部";
            else if (SelectedItems.Count == 1)
                _displayTextBlock.Text = SelectedItems[0];
            else
                _displayTextBlock.Text = $"已选 {SelectedItems.Count} 项";
        }

        #region 依赖属性回调

        private static void OnItemsSourceChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (MultiSelectComboBox)d;
            if (control._listBox != null)
            {
                control._listBox.ItemContainerGenerator.StatusChanged -= control.OnListBoxGenerated;
                control._listBox.ItemContainerGenerator.StatusChanged += control.OnListBoxGenerated;
            }
            if (control.SelectedItems != null && control.SelectedItems.Count == 0)
            {
                control.SelectedItems.Add("全部");
            }
            control.UpdateDisplayText();
        }

        private static void OnSelectedItemsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (MultiSelectComboBox)d;
            control.UpdateDisplayText();

            if (e.OldValue is ObservableCollection<string> oldCol)
                oldCol.CollectionChanged -= control.OnSelectedCollectionChanged;
            if (e.NewValue is ObservableCollection<string> newCol)
                newCol.CollectionChanged += control.OnSelectedCollectionChanged;
        }

        private void OnSelectedCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            UpdateDisplayText();
            if (_popup != null && _popup.IsOpen)
                UpdateCheckBoxStates();
        }

        #endregion

        #region 布局重写

        protected override int VisualChildrenCount => 1;

        protected override Visual GetVisualChild(int index)
        {
            if (index != 0) throw new ArgumentOutOfRangeException(nameof(index));
            return _rootBorder;
        }

        protected override Size MeasureOverride(Size constraint)
        {
            double w = constraint.Width;
            if (double.IsInfinity(w) || double.IsNaN(w) || w < 100) w = 120;
            double h = constraint.Height;
            if (double.IsInfinity(h) || double.IsNaN(h) || h < 26) h = 30;

            _rootBorder?.Measure(new Size(w, h));
            return new Size(w, h);
        }

        protected override Size ArrangeOverride(Size arrangeBounds)
        {
            _rootBorder?.Arrange(new Rect(0, 0, arrangeBounds.Width, arrangeBounds.Height));

            if (_popup != null)
            {
                _popup.Width = arrangeBounds.Width;
            }

            return arrangeBounds;
        }

        #endregion
    }
}
