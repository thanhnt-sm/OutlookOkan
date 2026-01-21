using OutlookOkan.Types;
using OutlookOkan.ViewModels;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace OutlookOkan.Views
{
    public partial class ConfirmationWindow : Window
    {
        private readonly dynamic _item;
        private readonly string _tempFilePath;

        public ConfirmationWindow(CheckList checkList, dynamic item)
        {
            DataContext = new ConfirmationWindowViewModel(checkList);

            _item = item;
            _tempFilePath = checkList.TempFilePath;

            InitializeComponent();

            // Nhập thời gian trễ gửi vào hộp hiển thị (thiết lập).
            DeferredDeliveryMinutesBox.Text = checkList.DeferredMinutes.ToString();

            //縦方向の最大サイズを制限h thước tối đa theo chiều dọc
            MaxHeight = SystemParameters.WorkArea.Height;

            // Tải kích thước cửa sổ
            if (Properties.Settings.Default.ConfirmationWindowWidth != 0)
            {
                Width = Properties.Settings.Default.ConfirmationWindowWidth;
            }

            if (Properties.Settings.Default.ConfirmationWindowHeight != 0)
            {
                Height = Properties.Settings.Default.ConfirmationWindowHeight;
            }
        }

        /// <summary>
        /// Vì việc Bind DialogResult là khó, nên thực hiện ở code-behind.
        /// </summary>
        private void SendButton_OnClick(object sender, RoutedEventArgs e)
        {
            // Thiết lập thời gian gửi
            _ = int.TryParse(DeferredDeliveryMinutesBox.Text, out var deferredDeliveryMinutes);

            if (deferredDeliveryMinutes != 0)
            {
                if (_item.DeferredDeliveryTime == new DateTime(4501, 1, 1, 0, 0, 0))
                {
                    // Trường hợp chỉ có thời gian hoãn được thiết lập bởi tính năng của add-in
                    _item.DeferredDeliveryTime = DateTime.Now.AddMinutes(deferredDeliveryMinutes);
                }
                else
                {
                    // Trường hợp thời gian hoãn (thời gian gửi) được thiết lập bởi cả tính năng của add-in và tính năng chuẩn của Outlook
                    if (DateTime.Now.AddMinutes(deferredDeliveryMinutes) > _item.DeferredDeliveryTime.AddMinutes(deferredDeliveryMinutes))
                    {
                        // Do [Ngày giờ gửi đã thiết lập + Thời gian hoãn bởi add-in] là trước [Ngày giờ hiện tại + Thời gian hoãn bởi add-in], nên chọn cái sau.
                        _item.DeferredDeliveryTime = DateTime.Now.AddMinutes(deferredDeliveryMinutes);
                    }
                    else
                    {
                        _item.DeferredDeliveryTime = _item.DeferredDeliveryTime.AddMinutes(deferredDeliveryMinutes);
                    }
                }
            }

            DialogResult = true;
        }

        /// <summary>
        /// Vì việc Bind DialogResult là khó, nên thực hiện ở code-behind.
        /// </summary>
        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        /// <summary>
        /// Vì việc xử lý sự kiện của checkbox là khó, nên gọi phương thức của ViewModel từ phía code-behind.
        /// </summary>
        private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        /// <summary>
        /// Vì việc xử lý sự kiện của checkbox là khó, nên gọi phương thức của ViewModel từ phía code-behind.
        /// </summary>
        private void ToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        /// <summary>
        /// Giới hạn hộp nhập liệu thời gian trễ gửi chỉ được nhập số.
        /// </summary>
        private void DeferredDeliveryMinutesBox_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var regex = new Regex("[^0-9]+$");
            if (!regex.IsMatch(DeferredDeliveryMinutesBox.Text + e.Text)) return;

            DeferredDeliveryMinutesBox.Text = "0";
            e.Handled = true;
        }

        /// <summary>
        /// Bỏ qua việc dán vào hộp nhập liệu thời gian trễ gửi. (Vì có nguy cơ số toàn giác được dán vào)
        /// </summary>
        private void DeferredDeliveryMinutesBox_OnPreviewExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Command == ApplicationCommands.Paste)
            {
                e.Handled = true;
            }
        }

        #region MouseUpEvent_OnHandler

        private void AlertGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            // Bỏ qua nếu không phải chuột trái. (Vì CurrentItem có thể bị lệch)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Alert)AlertGrid.CurrentItem;
            currentItem.IsChecked = !currentItem.IsChecked;
            AlertGrid.Items.Refresh();

            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        private void ToGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            // Bỏ qua nếu không phải chuột trái. (Vì CurrentItem có thể bị lệch)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Address)ToGrid.CurrentItem;
            currentItem.IsChecked = !currentItem.IsChecked;
            ToGrid.Items.Refresh();

            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        private void CcGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            // Bỏ qua nếu không phải chuột trái. (Vì CurrentItem có thể bị lệch)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Address)CcGrid.CurrentItem;
            currentItem.IsChecked = !currentItem.IsChecked;
            CcGrid.Items.Refresh();

            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        private void BccGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            // Bỏ qua nếu không phải chuột trái. (Vì CurrentItem có thể bị lệch)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Address)BccGrid.CurrentItem;
            currentItem.IsChecked = !currentItem.IsChecked;
            BccGrid.Items.Refresh();

            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        private void AttachmentGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            // Bỏ qua nếu không phải chuột trái. (Vì CurrentItem có thể bị lệch)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Attachment)AttachmentGrid.CurrentItem;
            var cell = GetDataGridObject<DataGridCell>(AttachmentGrid, e.GetPosition(AttachmentGrid));
            if (cell is null) return;
            var columnIndex = cell.Column.DisplayIndex;

            if (columnIndex == 1 && currentItem.IsCanOpen)
            {
                var result = MessageBox.Show(Properties.Resources.OpenTheAttachedFile + " (" + currentItem.FileName + ")" + Environment.NewLine + Properties.Resources.ChangesInTheFileWillNotBeSaved, Properties.Resources.OpenTheAttachedFile, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.ServiceNotification);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        var process = new ProcessStartInfo
                        {
                            UseShellExecute = true,
                            FileName = currentItem.FilePath,
                        };
                        Process.Start(process);
                    }
                    catch (Exception ex)
                    {
                        // Log error for debugging purposes
                        Debug.WriteLine($"[OutlookOkan] Failed to open attachment: {ex.Message}");
                    }
                    finally
                    {
                        currentItem.IsChecked = true;
                        AttachmentGrid.Items.Refresh();
                        var viewModel = DataContext as ConfirmationWindowViewModel;
                        viewModel?.ToggleSendButton();
                    }
                }

            }
            else
            {
                if (!currentItem.IsNotMustOpenBeforeCheck) return;

                currentItem.IsChecked = !currentItem.IsChecked;
                AttachmentGrid.Items.Refresh();

                var viewModel = DataContext as ConfirmationWindowViewModel;
                viewModel?.ToggleSendButton();
            }
        }

        private T GetDataGridObject<T>(Visual dataGrid, Point point)
        {
            var result = default(T);
            var hitResultTest = VisualTreeHelper.HitTest(dataGrid, point);
            if (hitResultTest == null) return result;
            var visualHit = hitResultTest.VisualHit;
            while (visualHit != null)
            {
                if (visualHit is T)
                {
                    result = (T)(object)visualHit;
                    break;
                }
                visualHit = VisualTreeHelper.GetParent(visualHit);
            }
            return result;
        }

        #endregion

        private void ConfirmationWindow_OnClosing(object sender, CancelEventArgs e)
        {
            // Lưu kích thước cửa sổ
            Properties.Settings.Default.ConfirmationWindowWidth = Width;
            Properties.Settings.Default.ConfirmationWindowHeight = Height;
            Properties.Settings.Default.Save();

            if (string.IsNullOrEmpty(_tempFilePath)) return;

            try
            {
                File.Delete(_tempFilePath);
            }
            catch (Exception ex)
            {
                // Log error for debugging - temp file cleanup is non-critical
                Debug.WriteLine($"[OutlookOkan] Failed to delete temp file: {ex.Message}");
            }
        }
    }
}