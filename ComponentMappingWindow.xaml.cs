using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Documents;
using System.Windows.Media.Animation;

namespace SchBom_Convert
{
    public partial class ComponentMappingWindow : Window
    {
        // DataGrid 的列檢視模型：關鍵字以逗號分隔字串呈現／編輯
        public class MappingRow
        {
            public string Name { get; set; } = "";
            public string Category { get; set; } = "";
            public string KeywordsText { get; set; } = "";
        }

        private readonly ObservableCollection<MappingRow> _rows = new();

        // 拖曳排序狀態
        private MappingRow? _dragItem;
        private Point _dragStartPoint;
        private AdornerLayer? _adornerLayer;
        private InsertionAdorner? _insertionAdorner;
        private DragGhostAdorner? _ghostAdorner;
        private int _lastDropRow = -2;
        private bool _lastDropAbove;

        public ComponentMappingWindow()
        {
            InitializeComponent();
            MappingGrid.ItemsSource = _rows;
            LoadFromService();
        }

        // 從服務載入現有規則（保序）
        private void LoadFromService()
        {
            _rows.Clear();
            foreach (var m in ComponentMappingService.Instance.Mappings)
            {
                _rows.Add(new MappingRow
                {
                    Name = m.Name,
                    Category = m.Category,
                    KeywordsText = string.Join(", ", m.Keywords ?? new List<string>())
                });
            }
        }

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            var row = new MappingRow { Category = "SMD" };
            int sel = MappingGrid.SelectedIndex;
            if (sel >= 0) _rows.Insert(sel + 1, row);   // 插在選取列之後，不再跑到最下面
            else _rows.Add(row);
            MappingGrid.SelectedItem = row;
            MappingGrid.ScrollIntoView(row);
        }

        private void DeleteRow_Click(object sender, RoutedEventArgs e)
        {
            if (MappingGrid.SelectedItem is MappingRow row)
            {
                int idx = _rows.IndexOf(row);
                _rows.Remove(row);
                if (_rows.Count > 0)
                    MappingGrid.SelectedIndex = Math.Min(idx, _rows.Count - 1);
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            int i = MappingGrid.SelectedIndex;
            if (i > 0)
            {
                _rows.Move(i, i - 1);
                MappingGrid.SelectedIndex = i - 1;
                MappingGrid.ScrollIntoView(MappingGrid.SelectedItem);
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            int i = MappingGrid.SelectedIndex;
            if (i >= 0 && i < _rows.Count - 1)
            {
                _rows.Move(i, i + 1);
                MappingGrid.SelectedIndex = i + 1;
                MappingGrid.ScrollIntoView(MappingGrid.SelectedItem);
            }
        }

        private void ResetDefaults_Click(object sender, RoutedEventArgs e)
        {
            var r = MessageBox.Show(
                "確定要還原成內建預設規則嗎？目前所有自訂規則將被覆蓋。",
                "還原預設", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (r == MessageBoxResult.Yes)
            {
                ComponentMappingService.Instance.ResetToDefaults();
                LoadFromService();
                MessageBox.Show("已還原為預設規則。", "完成",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // 先提交 DataGrid 正在編輯中的儲存格／列
            MappingGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            MappingGrid.CommitEdit(DataGridEditingUnit.Row, true);

            var list = new List<ComponentMapping>();
            foreach (var row in _rows)
            {
                string name = (row.Name ?? "").Trim();
                var keywords = SplitKeywords(row.KeywordsText);

                // 略過名稱與關鍵字皆空的列
                if (string.IsNullOrEmpty(name) && keywords.Count == 0)
                    continue;

                list.Add(new ComponentMapping
                {
                    Name = name,
                    Category = (row.Category ?? "").Trim(),
                    Keywords = keywords
                });
            }

            ComponentMappingService.Instance.Save(list);
            MessageBox.Show($"已儲存 {list.Count} 筆分類規則。", "完成",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // 只用逗號／分號（含全形）切，保留含空白的多字關鍵字（如 "ISO RS485"、"Y2 CAP"）
        private static List<string> SplitKeywords(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return new List<string>();
            return text.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries)
                       .Select(s => s.Trim())
                       .Where(s => s.Length > 0)
                       .ToList();
        }

        // ===== 拖曳排序 =====
        // 只從列首握把（≡）開始拖曳，避免干擾儲存格文字編輯
        private void MappingGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragItem = null;
            // 只有點在「≡」握把欄才視為拖曳起點，避免干擾其他儲存格編輯
            var cell = FindAncestor<DataGridCell>(e.OriginalSource as DependencyObject);
            if (cell != null && cell.Column == HandleColumn)
            {
                _dragStartPoint = e.GetPosition(null);
                _dragItem = FindAncestor<DataGridRow>(cell)?.Item as MappingRow;
            }
        }

        private void MappingGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (_dragItem == null || e.LeftButton != MouseButtonState.Pressed) return;

            var pos = e.GetPosition(null);
            if (Math.Abs(pos.X - _dragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(pos.Y - _dragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            ShowInsertionAdorner();
            try
            {
                DragDrop.DoDragDrop(MappingGrid, _dragItem, DragDropEffects.Move);
            }
            finally
            {
                RemoveInsertionAdorner();
                _dragItem = null;
            }
        }

        private void MappingGrid_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Move;
            e.Handled = true;

            _ghostAdorner?.SetMouse(e.GetPosition(MappingGrid));

            if (_insertionAdorner == null) return;
            if (TryGetDropTarget(e, out int rowIndex, out bool above, out double y))
            {
                if (rowIndex != _lastDropRow || above != _lastDropAbove)
                {
                    _lastDropRow = rowIndex;
                    _lastDropAbove = above;
                    _insertionAdorner.MoveLineTo(y);
                }
            }
        }

        private void MappingGrid_Drop(object sender, DragEventArgs e)
        {
            if (_dragItem == null) return;

            int oldIndex = _rows.IndexOf(_dragItem);
            if (oldIndex < 0) return;

            if (TryGetDropTarget(e, out int rowIndex, out bool above, out _))
            {
                int insertIndex = above ? rowIndex : rowIndex + 1;
                if (oldIndex < insertIndex) insertIndex--;          // 移除來源後索引前移
                insertIndex = Math.Max(0, Math.Min(insertIndex, _rows.Count - 1));

                if (insertIndex != oldIndex)
                {
                    _rows.Move(oldIndex, insertIndex);
                    MappingGrid.SelectedIndex = insertIndex;
                }
            }
            // _dragItem 由 MouseMove 的 finally 清除
        }

        // 計算放置目標：哪一列、插在上方或下方、以及指示線的 Y 座標
        private bool TryGetDropTarget(DragEventArgs e, out int rowIndex, out bool insertAbove, out double yLine)
        {
            rowIndex = -1; insertAbove = true; yLine = 0;
            if (_rows.Count == 0) return false;

            var row = FindAncestor<DataGridRow>(e.OriginalSource as DependencyObject);
            if (row?.Item is MappingRow item)
            {
                rowIndex = _rows.IndexOf(item);
                var b = RowBounds(row);
                double my = e.GetPosition(MappingGrid).Y;
                insertAbove = my < b.Top + b.Height / 2;
                yLine = insertAbove ? b.Top : b.Bottom;
                return true;
            }

            // 不在任何列上（多半拖到清單下方空白）→ 視為最後一列之後
            rowIndex = _rows.Count - 1;
            insertAbove = false;
            if (MappingGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex) is DataGridRow lastRow)
            {
                yLine = RowBounds(lastRow).Bottom;
                return true;
            }
            return false;
        }

        private Rect RowBounds(DataGridRow row)
        {
            var t = row.TransformToAncestor(MappingGrid);
            return t.TransformBounds(new Rect(new Point(0, 0), row.RenderSize));
        }

        private void ShowInsertionAdorner()
        {
            _adornerLayer = AdornerLayer.GetAdornerLayer(MappingGrid);
            if (_adornerLayer == null) return;
            _insertionAdorner = new InsertionAdorner(MappingGrid);
            _adornerLayer.Add(_insertionAdorner);
            _lastDropRow = -2;

            // 拖曳殘影：被拖那一列的半透明複本，跟著游標走
            if (_dragItem != null &&
                MappingGrid.ItemContainerGenerator.ContainerFromItem(_dragItem) is DataGridRow row)
            {
                _ghostAdorner = new DragGhostAdorner(MappingGrid, row, row.RenderSize);
                _adornerLayer.Add(_ghostAdorner);
                _ghostAdorner.SetMouse(Mouse.GetPosition(MappingGrid));
            }
        }

        private void RemoveInsertionAdorner()
        {
            if (_adornerLayer != null)
            {
                if (_insertionAdorner != null) _adornerLayer.Remove(_insertionAdorner);
                if (_ghostAdorner != null) _adornerLayer.Remove(_ghostAdorner);
            }
            _insertionAdorner = null;
            _ghostAdorner = null;
            _adornerLayer = null;
        }

        private static T? FindAncestor<T>(DependencyObject? obj) where T : DependencyObject
        {
            while (obj != null && obj is not T)
                obj = VisualTreeHelper.GetParent(obj);
            return obj as T;
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }

    // DataGrid 拖曳排序時顯示的插入位置指示線（含淡入與平滑滑動動畫）
    internal sealed class InsertionAdorner : Adorner
    {
        private static readonly Brush LineBrush = MakeBrush();
        private static readonly Pen LinePen = MakePen();
        private bool _initialized;

        public static readonly DependencyProperty LineYProperty =
            DependencyProperty.Register(nameof(LineY), typeof(double), typeof(InsertionAdorner),
                new FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.AffectsRender));

        public double LineY
        {
            get => (double)GetValue(LineYProperty);
            set => SetValue(LineYProperty, value);
        }

        public InsertionAdorner(UIElement adornedElement) : base(adornedElement)
        {
            IsHitTestVisible = false;
            Opacity = 0;
            BeginAnimation(OpacityProperty, new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(120)));
        }

        // 平滑滑動到新的插入位置；第一次直接定位避免從頂端飛入
        public void MoveLineTo(double y)
        {
            if (!_initialized)
            {
                _initialized = true;
                BeginAnimation(LineYProperty, null);
                LineY = y;
                return;
            }
            var anim = new DoubleAnimation(y, TimeSpan.FromMilliseconds(120))
            {
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };
            BeginAnimation(LineYProperty, anim);
        }

        protected override void OnRender(DrawingContext dc)
        {
            double width = AdornedElement.RenderSize.Width;
            double y = LineY;
            dc.DrawLine(LinePen, new Point(0, y), new Point(width, y));

            const double s = 5;
            dc.DrawGeometry(LineBrush, null, Triangle(new Point(0, y), s, true));
            dc.DrawGeometry(LineBrush, null, Triangle(new Point(width, y), s, false));
        }

        private static StreamGeometry Triangle(Point tip, double size, bool pointRight)
        {
            double dir = pointRight ? 1 : -1;
            var g = new StreamGeometry();
            using (var c = g.Open())
            {
                c.BeginFigure(new Point(tip.X, tip.Y - size), true, true);
                c.LineTo(new Point(tip.X, tip.Y + size), true, false);
                c.LineTo(new Point(tip.X + dir * size, tip.Y), true, false);
            }
            g.Freeze();
            return g;
        }

        private static Brush MakeBrush()
        {
            var b = new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50));
            b.Freeze();
            return b;
        }

        private static Pen MakePen()
        {
            var p = new Pen(MakeBrush(), 2);
            p.Freeze();
            return p;
        }
    }

    // 拖曳時跟著游標走的半透明殘影（被拖那一列的複本）
    internal sealed class DragGhostAdorner : Adorner
    {
        private readonly Brush _brush;
        private readonly Pen _border;
        private readonly Size _size;
        private Point _topLeft;

        public DragGhostAdorner(UIElement adornedElement, Visual rowVisual, Size size) : base(adornedElement)
        {
            IsHitTestVisible = false;
            _size = size;
            _brush = new VisualBrush(rowVisual) { Stretch = Stretch.Fill };

            var bp = new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50));
            bp.Freeze();
            _border = new Pen(bp, 1);
            _border.Freeze();

            Opacity = 0;
            BeginAnimation(OpacityProperty, new DoubleAnimation(0, 0.6, TimeSpan.FromMilliseconds(100)));
        }

        // 以目前游標（DataGrid 座標）垂直置中放置殘影
        public void SetMouse(Point gridPoint)
        {
            _topLeft = new Point(0, gridPoint.Y - _size.Height / 2);
            InvalidateVisual();
        }

        protected override void OnRender(DrawingContext dc)
        {
            dc.DrawRectangle(_brush, _border, new Rect(_topLeft, _size));
        }
    }
}
