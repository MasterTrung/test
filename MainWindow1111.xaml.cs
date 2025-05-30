using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DragDropTreeApp.Models;

// Sử dụng alias để tránh xung đột namespace
using WinForms = System.Windows.Forms;
using WinControls = System.Windows.Controls;
using WinDocuments = System.Windows.Documents;
using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;
using OpenXmlWordprocessing = DocumentFormat.OpenXml.Wordprocessing;

// Định nghĩa lại các tên xung đột
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;


using System.Collections.Generic;

namespace DragDropTreeApp
{
    public partial class MainWindow : Window
    {
        private ObservableCollection<DocNode> _docNodes;

        // Biến cho việc kéo thả
        private Point _startPoint;
        private TreeViewItem _draggedItem;
        private DocNode _draggedNode;
        private bool _isDragging = false;

        // Thêm các biến cho chức năng Undo
        private Stack<UndoAction> _undoStack = new Stack<UndoAction>();
        private bool _isUndoing = false;
        // Khai báo thêm biến
        private ObservableCollection<DocNode> _detailNodes;



        public MainWindow()
        {
            InitializeComponent();

            _docNodes = new ObservableCollection<DocNode>();
            _detailNodes = new ObservableCollection<DocNode>();


            TreeDoc.ItemsSource = _docNodes;
            TreeDocDetail.ItemsSource = _detailNodes;

            UpdateEmptyState();

            // Đăng ký phím tắt Ctrl+Z
            CommandBindings.Add(new CommandBinding(ApplicationCommands.Undo, UndoCommand_Execute, UndoCommand_CanExecute));
        }


        // Thêm hàm xử lý cho button Undo
        private void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            if (_undoStack.Count > 0)
            {
                UndoCommand_Execute(this, null);
            }
        }


        // Thêm các hàm xử lý cho TreeViewDetail
        private void TreeDocDetail_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            // Xử lý tương tự như TreeDoc_MouseRightButtonUp nhưng làm việc với _detailNodes
        }

        private void TreeDocDetail_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(null);
        }

        private void TreeDocDetail_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            // Kiểm tra nếu nút trái được nhấn và đang di chuyển
            if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
            {
                Point currentPosition = e.GetPosition(null);

                // Kiểm tra khoảng cách để xác định có phải đang kéo hay không
                if (Math.Abs(currentPosition.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(currentPosition.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    // Xác định TreeViewItem được kéo
                    _draggedItem = FindTreeViewItemUnderMouse(e.GetPosition(TreeDocDetail), TreeDocDetail);

                    if (_draggedItem != null)
                    {
                        _draggedNode = _draggedItem.Header as DocNode;

                        if (_draggedNode != null)
                        {
                            // Bắt đầu kéo
                            _isDragging = true;

                            // Đánh dấu nguồn là panel phải
                            DataObject dragData = new DataObject();
                            dragData.SetData("DocNodeFormat", _draggedNode);
                            dragData.SetData("SourcePanel", "Detail");

                            DragDrop.DoDragDrop(_draggedItem, dragData, DragDropEffects.Move | DragDropEffects.Copy);

                            // Reset trạng thái kéo
                            _isDragging = false;
                            _draggedItem = null;
                            _draggedNode = null;

                            // Xóa chỉ báo drop target
                            ClearDropTargetIndicators();
                        }
                    }
                }
            }
        }

        private void TreeDocDetail_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("DocNodeFormat"))
            {
                // Kiểm tra source panel
                string sourcePanel = (string)e.Data.GetData("SourcePanel");

                // Nếu từ panel trái, cho phép sao chép
                if (sourcePanel == "Main")
                {
                    e.Effects = DragDropEffects.Copy;
                }
                // Nếu từ panel phải, cho phép di chuyển
                else
                {
                    e.Effects = DragDropEffects.Move;
                }

                // Xóa tất cả chỉ báo trước đó
                ClearDropTargetIndicators(TreeDocDetail);

                // Hiển thị chỉ báo vị trí thả
                Point position = e.GetPosition(TreeDocDetail);
                TreeViewItem targetItem = FindTreeViewItemUnderMouse(position, TreeDocDetail);

                if (targetItem != null && targetItem != _draggedItem)
                {
                    targetItem.Tag = "DropTarget";
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void TreeDocDetail_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("DocNodeFormat"))
            {
                DocNode draggedNode = (DocNode)e.Data.GetData("DocNodeFormat");
                string sourcePanel = (string)e.Data.GetData("SourcePanel");

                // Tìm item đích để thả
                Point position = e.GetPosition(TreeDocDetail);
                TreeViewItem targetItem = FindTreeViewItemUnderMouse(position, TreeDocDetail);
                int targetIndex = -1;

                if (targetItem != null)
                {
                    DocNode targetNode = targetItem.Header as DocNode;
                    if (targetNode != null)
                    {
                        targetIndex = _detailNodes.IndexOf(targetNode);
                    }
                }

                if (targetIndex == -1)
                {
                    targetIndex = _detailNodes.Count;
                }

                // Nếu từ panel trái, sao chép sang phải
                if (sourcePanel == "Main")
                {
                    var nodeCopy = CloneDocNode(draggedNode);
                    _detailNodes.Insert(targetIndex, nodeCopy);

                    StatusTextBlock.Text = "Đã sao chép phần tử sang panel Chi tiết";
                }
                // Nếu từ panel phải, di chuyển trong panel phải
                else
                {
                    int sourceIndex = _detailNodes.IndexOf(draggedNode);
                    if (sourceIndex != -1 && sourceIndex != targetIndex)
                    {
                        var movedNode = _detailNodes[sourceIndex];
                        movedNode.IsAnimating = true;
                        _detailNodes.RemoveAt(sourceIndex);

                        if (targetIndex > sourceIndex)
                            targetIndex--;

                        _detailNodes.Insert(targetIndex, movedNode);

                        Task.Delay(500).ContinueWith(_ => {
                            Application.Current.Dispatcher.Invoke(() => {
                                movedNode.IsAnimating = false;
                            });
                        });

                        StatusTextBlock.Text = "Đã di chuyển phần tử trong panel Chi tiết";
                    }
                }

                // Cập nhật overlay
                UpdateDetailEmptyState();

                // Xóa chỉ báo drop target
                ClearDropTargetIndicators(TreeDocDetail);
            }
        }

        private TreeViewItem FindTreeViewItemUnderMouse(Point point, TreeView treeView)
        {
            HitTestResult result = VisualTreeHelper.HitTest(treeView, point);
            if (result != null)
            {
                DependencyObject obj = result.VisualHit;
                while (obj != null && !(obj is TreeViewItem))
                {
                    obj = VisualTreeHelper.GetParent(obj);
                }

                return obj as TreeViewItem;
            }

            return null;
        }

        private void ClearDropTargetIndicators(TreeView treeView)
        {
            var collection = treeView == TreeDoc ? _docNodes : _detailNodes;

            foreach (var node in collection)
            {
                int index = collection.IndexOf(node);
                TreeViewItem item = treeView.ItemContainerGenerator.ContainerFromIndex(index) as TreeViewItem;
                if (item != null)
                {
                    item.Tag = null;
                }
            }
        }

        //// Cập nhật lại phương thức TreeDoc_PreviewMouseMove
        //private void TreeDoc_PreviewMouseMove(object sender, MouseEventArgs e)
        //{
        //    // Kiểm tra nếu nút trái được nhấn và đang di chuyển
        //    if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
        //    {
        //        Point currentPosition = e.GetPosition(null);

        //        // Kiểm tra khoảng cách để xác định có phải đang kéo hay không
        //        if (Math.Abs(currentPosition.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
        //            Math.Abs(currentPosition.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
        //        {
        //            // Xác định TreeViewItem được kéo
        //            _draggedItem = FindTreeViewItemUnderMouse(e.GetPosition(TreeDoc), TreeDoc);

        //            if (_draggedItem != null)
        //            {
        //                _draggedNode = _draggedItem.Header as DocNode;

        //                if (_draggedNode != null)
        //                {
        //                    // Bắt đầu kéo
        //                    _isDragging = true;

        //                    // Đánh dấu nguồn là panel trái
        //                    DataObject dragData = new DataObject();
        //                    dragData.SetData("DocNodeFormat", _draggedNode);
        //                    dragData.SetData("SourcePanel", "Main");

        //                    DragDrop.DoDragDrop(_draggedItem, dragData, DragDropEffects.Move | DragDropEffects.Copy);

        //                    // Reset trạng thái kéo
        //                    _isDragging = false;
        //                    _draggedItem = null;
        //                    _draggedNode = null;

        //                    // Xóa chỉ báo drop target
        //                    ClearDropTargetIndicators(TreeDoc);
        //                }
        //            }
        //        }
        //    }
        //}

        // 2. Gộp logic lại làm một phương thức duy nhất:
        private void TreeDoc_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            // Kiểm tra nếu nút trái được nhấn và đang di chuyển
            if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
            {
                Point currentPosition = e.GetPosition(null);

                // Kiểm tra khoảng cách để xác định có phải đang kéo hay không
                if (Math.Abs(currentPosition.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(currentPosition.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    // Xác định nguồn (Panel nào)
                    TreeView sourceTreeView = sender as TreeView;
                    string sourcePanel = sourceTreeView == TreeDoc ? "Main" : "Detail";

                    // Xác định TreeViewItem được kéo
                    _draggedItem = FindTreeViewItemUnderMouse(e.GetPosition(sourceTreeView), sourceTreeView);

                    if (_draggedItem != null)
                    {
                        _draggedNode = _draggedItem.Header as DocNode;

                        if (_draggedNode != null)
                        {
                            // Bắt đầu kéo
                            _isDragging = true;

                            // Đánh dấu nguồn (Main hoặc Detail)
                            DataObject dragData = new DataObject();
                            dragData.SetData("DocNodeFormat", _draggedNode);
                            dragData.SetData("SourcePanel", sourcePanel);

                            DragDrop.DoDragDrop(_draggedItem, dragData, DragDropEffects.Move | DragDropEffects.Copy);

                            // Reset trạng thái kéo
                            _isDragging = false;
                            _draggedItem = null;
                            _draggedNode = null;

                            // Xóa chỉ báo drop target
                            ClearDropTargetIndicators(sourcePanel == "Main" ? TreeDoc : TreeDocDetail);
                        }
                    }
                }
            }
        }

        // Cập nhật lại phương thức TreeDoc_DragOver và TreeDoc_Drop tương tự
        private void TreeDoc_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("DocNodeFormat"))
            {
                // Kiểm tra source panel
                string sourcePanel = (string)e.Data.GetData("SourcePanel");

                if (sourcePanel == "Detail")
                {
                    e.Effects = DragDropEffects.Copy;
                }
                else
                {
                    e.Effects = DragDropEffects.Move;
                }

                // Xóa tất cả chỉ báo trước đó
                ClearDropTargetIndicators(TreeDoc);

                // Hiển thị chỉ báo vị trí thả
                Point position = e.GetPosition(TreeDoc);
                TreeViewItem targetItem = FindTreeViewItemUnderMouse(position, TreeDoc);

                if (targetItem != null && targetItem != _draggedItem)
                {
                    targetItem.Tag = "DropTarget";
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void TreeDoc_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("DocNodeFormat"))
            {
                DocNode draggedNode = (DocNode)e.Data.GetData("DocNodeFormat");
                string sourcePanel = (string)e.Data.GetData("SourcePanel");

                // Tìm item đích để thả
                Point position = e.GetPosition(TreeDoc);
                TreeViewItem targetItem = FindTreeViewItemUnderMouse(position, TreeDoc);
                int targetIndex = -1;

                if (targetItem != null)
                {
                    DocNode targetNode = targetItem.Header as DocNode;
                    if (targetNode != null)
                    {
                        targetIndex = _docNodes.IndexOf(targetNode);
                    }
                }

                if (targetIndex == -1)
                {
                    targetIndex = _docNodes.Count;
                }

                // Nếu từ panel phải, sao chép sang trái
                if (sourcePanel == "Detail")
                {
                    var nodeCopy = CloneDocNode(draggedNode);
                    _docNodes.Insert(targetIndex, nodeCopy);

                    StatusTextBlock.Text = "Đã sao chép phần tử từ panel Chi tiết";
                }
                // Nếu từ panel trái, di chuyển trong panel trái
                else
                {
                    int sourceIndex = _docNodes.IndexOf(draggedNode);
                    if (sourceIndex != -1 && sourceIndex != targetIndex)
                    {
                        MoveItemWithAnimation(sourceIndex, targetIndex);
                    }
                }

                // Xóa chỉ báo drop target
                ClearDropTargetIndicators(TreeDoc);
            }
        }

        private DocNode CloneDocNode(DocNode source)
        {
            // Tạo bản sao của DocNode
            var clone = new DocNode
            {
                Content = source.Content,
                NodeType = source.NodeType,
                ImagePath = source.ImagePath,
                IsMixable = source.IsMixable
            };

            // Sao chép dữ liệu bảng nếu có
            if (source.TableData != null && source.TableData.Count > 0)
            {
                clone.TableData = new ObservableCollection<ObservableCollection<string>>();
                foreach (var row in source.TableData)
                {
                    var newRow = new ObservableCollection<string>();
                    foreach (var cell in row)
                    {
                        newRow.Add(cell);
                    }
                    clone.TableData.Add(newRow);
                }
            }

            // Sao chép thông tin ảnh trong bảng nếu có
            if (source.ImagesInTable != null && source.ImagesInTable.Count > 0)
            {
                clone.ImagesInTable = new List<TableImageInfo>();
                foreach (var img in source.ImagesInTable)
                {
                    clone.ImagesInTable.Add(new TableImageInfo
                    {
                        ImagePath = img.ImagePath,
                        ImageData = img.ImageData,
                        Row = img.Row,
                        Column = img.Column
                    });
                }
            }

            return clone;
        }

        private void UpdateDetailEmptyState()
        {
            if (_detailNodes != null && _detailNodes.Count > 0)
            {
                EmptyStateDetailOverlay.Visibility = Visibility.Collapsed;
            }
            else
            {
                EmptyStateDetailOverlay.Visibility = Visibility.Visible;
            }
        }

        // Bổ sung code cho menu chuột phải
        private void TreeDoc_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            // Lấy danh sách các mục được chọn
            var selectedItems = GetSelectedItems();

            // Nếu không có mục nào được chọn, không làm gì
            if (selectedItems.Count == 0)
                return;

            // Tạo menu chuột phải
            ContextMenu contextMenu = new ContextMenu();

            // Thêm các MenuItem
            var mergeItem = new MenuItem { Header = "Gộp thành một khối" };
            mergeItem.Click += (s, args) => MergeSelectedNodes(selectedItems);
            contextMenu.Items.Add(mergeItem);

            var splitItem = new MenuItem { Header = "Tách khỏi khối" };
            splitItem.Click += (s, args) => SplitSelectedNodes(selectedItems);
            contextMenu.Items.Add(splitItem);

            contextMenu.Items.Add(new Separator());

            var moveUpItem = new MenuItem { Header = "Di chuyển lên trên" };
            moveUpItem.Click += (s, args) => MoveSelectedNodesUp(selectedItems);
            contextMenu.Items.Add(moveUpItem);

            var moveDownItem = new MenuItem { Header = "Di chuyển xuống dưới" };
            moveDownItem.Click += (s, args) => MoveSelectedNodesDown(selectedItems);
            contextMenu.Items.Add(moveDownItem);

            contextMenu.Items.Add(new Separator());

            var setNotMixableItem = new MenuItem { Header = "Đặt là Không trộn" };
            setNotMixableItem.Click += (s, args) => SetSelectedNodesMixable(selectedItems, false);
            contextMenu.Items.Add(setNotMixableItem);

            var setMixableItem = new MenuItem { Header = "Đặt là Có trộn" };
            setMixableItem.Click += (s, args) => SetSelectedNodesMixable(selectedItems, true);
            contextMenu.Items.Add(setMixableItem);

            // Hiển thị menu
            contextMenu.IsOpen = true;
        }

        private List<DocNode> GetSelectedItems()
        {
            var selectedItems = new List<DocNode>();

            for (int i = 0; i < _docNodes.Count; i++)
            {
                var item = TreeDoc.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                if (item != null && item.IsSelected)
                {
                    selectedItems.Add(_docNodes[i]);
                }
            }

            return selectedItems;
        }

        private void MergeSelectedNodes(List<DocNode> selectedNodes)
        {
            if (selectedNodes.Count <= 1)
                return;

            // Tìm vị trí đầu tiên trong danh sách
            int startIndex = int.MaxValue;
            foreach (var node in selectedNodes)
            {
                int index = _docNodes.IndexOf(node);
                startIndex = Math.Min(startIndex, index);
            }

            // Kiểm tra xem các mục có liên tiếp không
            bool areConsecutive = true;
            for (int i = 0; i < selectedNodes.Count; i++)
            {
                if (_docNodes.IndexOf(selectedNodes[i]) != startIndex + i)
                {
                    areConsecutive = false;
                    break;
                }
            }

            if (!areConsecutive)
            {
                MessageBox.Show("Chỉ có thể gộp các mục liên tiếp!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Lưu lại cho Undo
            var undoAction = new UndoAction(UndoAction.ActionType.Merge)
            {
                SourceIndex = startIndex,
                TargetIndex = startIndex,
                NodeList = new List<DocNode>(selectedNodes)
            };
            _undoStack.Push(undoAction);

            // Gộp nội dung (ví dụ với text)
            string mergedContent = string.Join("\n", selectedNodes.Select(n => n.Content));

            // Tạo node mới với nội dung gộp
            var mergedNode = new DocNode
            {
                NodeType = DocNodeType.Text,
                Content = mergedContent,
                IsMixable = selectedNodes.All(n => n.IsMixable)
            };

            // Xóa các node cũ
            for (int i = 0; i < selectedNodes.Count; i++)
            {
                _docNodes.RemoveAt(startIndex);
            }

            // Thêm node mới
            _docNodes.Insert(startIndex, mergedNode);

            // Cập nhật giao diện
            SelectItemByIndex(startIndex);
            StatusTextBlock.Text = "Đã gộp các phần tử thành một khối";
        }

        private void SplitSelectedNodes(List<DocNode> selectedNodes)
        {
            if (selectedNodes.Count != 1 || selectedNodes[0].NodeType != DocNodeType.Text)
            {
                MessageBox.Show("Chỉ có thể tách một khối văn bản!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DocNode nodeToSplit = selectedNodes[0];
            int index = _docNodes.IndexOf(nodeToSplit);

            // Tách nội dung theo dòng
            string[] lines = nodeToSplit.Content.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            if (lines.Length <= 1)
            {
                MessageBox.Show("Không thể tách nội dung này!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Lưu lại cho Undo
            var undoAction = new UndoAction(UndoAction.ActionType.Split)
            {
                SourceIndex = index,
                Node = nodeToSplit
            };

            // Xóa node cũ
            _docNodes.RemoveAt(index);

            // Thêm các node mới
            List<DocNode> newNodes = new List<DocNode>();
            for (int i = 0; i < lines.Length; i++)
            {
                var newNode = new DocNode
                {
                    NodeType = DocNodeType.Text,
                    Content = lines[i],
                    IsMixable = nodeToSplit.IsMixable
                };

                _docNodes.Insert(index + i, newNode);
                newNodes.Add(newNode);
            }

            // Lưu danh sách node mới cho Undo
            undoAction.NodeList = newNodes;
            _undoStack.Push(undoAction);

            // Cập nhật giao diện
            SelectItemByIndex(index);
            StatusTextBlock.Text = "Đã tách thành " + lines.Length + " phần tử";
        }

        private void MoveSelectedNodesUp(List<DocNode> selectedNodes)
        {
            if (selectedNodes.Count == 0)
                return;

            // Tìm chỉ số nhỏ nhất (cao nhất trong danh sách)
            int minIndex = int.MaxValue;
            foreach (var node in selectedNodes)
            {
                int index = _docNodes.IndexOf(node);
                minIndex = Math.Min(minIndex, index);
            }

            // Không thể di chuyển lên nếu đã ở trên cùng
            if (minIndex <= 0)
                return;

            // Lưu lại cho Undo
            var undoAction = new UndoAction(UndoAction.ActionType.Move);

            // Di chuyển lên từng mục một
            foreach (var node in selectedNodes.OrderBy(n => _docNodes.IndexOf(n)))
            {
                int index = _docNodes.IndexOf(node);
                if (index > 0 && !selectedNodes.Contains(_docNodes[index - 1]))
                {
                    // Lưu thông tin di chuyển
                    undoAction.NodeList.Add(node);

                    // Di chuyển lên
                    _docNodes.RemoveAt(index);
                    _docNodes.Insert(index - 1, node);
                }
            }

            if (undoAction.NodeList.Count > 0)
                _undoStack.Push(undoAction);

            StatusTextBlock.Text = "Đã di chuyển các mục lên trên";
        }

        private void MoveSelectedNodesDown(List<DocNode> selectedNodes)
        {
            if (selectedNodes.Count == 0)
                return;

            // Tìm chỉ số lớn nhất (thấp nhất trong danh sách)
            int maxIndex = -1;
            foreach (var node in selectedNodes)
            {
                int index = _docNodes.IndexOf(node);
                maxIndex = Math.Max(maxIndex, index);
            }

            // Không thể di chuyển xuống nếu đã ở dưới cùng
            if (maxIndex >= _docNodes.Count - 1)
                return;

            // Lưu lại cho Undo
            var undoAction = new UndoAction(UndoAction.ActionType.Move);

            // Di chuyển xuống từng mục một (từ dưới lên để tránh xung đột index)
            foreach (var node in selectedNodes.OrderByDescending(n => _docNodes.IndexOf(n)))
            {
                int index = _docNodes.IndexOf(node);
                if (index < _docNodes.Count - 1 && !selectedNodes.Contains(_docNodes[index + 1]))
                {
                    // Lưu thông tin di chuyển
                    undoAction.NodeList.Add(node);

                    // Di chuyển xuống
                    _docNodes.RemoveAt(index);
                    _docNodes.Insert(index + 1, node);
                }
            }

            if (undoAction.NodeList.Count > 0)
                _undoStack.Push(undoAction);

            StatusTextBlock.Text = "Đã di chuyển các mục xuống dưới";
        }

        private void SetSelectedNodesMixable(List<DocNode> selectedNodes, bool mixable)
        {
            if (selectedNodes.Count == 0)
                return;

            // Lưu lại trạng thái cũ cho Undo
            var undoAction = new UndoAction(UndoAction.ActionType.SetMixable)
            {
                NodeList = new List<DocNode>(selectedNodes),
                OldMixableState = selectedNodes[0].IsMixable
            };
            _undoStack.Push(undoAction);

            // Đặt trạng thái mới
            foreach (var node in selectedNodes)
            {
                node.IsMixable = mixable;
            }

            StatusTextBlock.Text = mixable ? "Đã đặt các mục là Có trộn" : "Đã đặt các mục là Không trộn";
        }


        // Thêm class UndoAction để lưu hành động
        private class UndoAction
        {
            public enum ActionType
            {
                Move,
                Delete,
                Add,
                Merge,
                Split,
                SetMixable
            }

            public ActionType Type { get; set; }
            public int SourceIndex { get; set; }
            public int TargetIndex { get; set; }
            public DocNode Node { get; set; }
            public List<DocNode> NodeList { get; set; }
            public bool OldMixableState { get; set; }

            public UndoAction(ActionType type)
            {
                Type = type;
                NodeList = new List<DocNode>();
            }
        }

        // Kiểm tra xem có thể thực hiện Undo không
        private void UndoCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = _undoStack.Count > 0;
        }

        // Thực hiện Undo
        private void UndoCommand_Execute(object sender, ExecutedRoutedEventArgs e)
        {
            if (_undoStack.Count == 0)
                return;

            _isUndoing = true;

            var action = _undoStack.Pop();

            switch (action.Type)
            {
                case UndoAction.ActionType.Move:
                    // Hoàn tác di chuyển: Di chuyển từ vị trí hiện tại về vị trí cũ
                    var node = _docNodes[action.TargetIndex];
                    _docNodes.RemoveAt(action.TargetIndex);
                    _docNodes.Insert(action.SourceIndex, node);
                    SelectItemByIndex(action.SourceIndex);
                    break;

                case UndoAction.ActionType.Delete:
                    // Hoàn tác xóa: Thêm lại các mục đã xóa
                    foreach (var nodePair in action.NodeList.Select((node, index) => new { Node = node, Index = action.SourceIndex + index }))
                    {
                        _docNodes.Insert(nodePair.Index, nodePair.Node);
                    }
                    SelectItemByIndex(action.SourceIndex);
                    break;

                case UndoAction.ActionType.Add:
                    // Hoàn tác thêm: Xóa mục đã thêm
                    _docNodes.RemoveAt(action.TargetIndex);
                    break;

                case UndoAction.ActionType.Merge:
                    // Hoàn tác gộp: Tách lại các mục đã gộp
                    _docNodes.RemoveAt(action.TargetIndex);
                    foreach (var nodePair in action.NodeList.Select((node, index) => new { Node = node, Index = action.SourceIndex + index }))
                    {
                        _docNodes.Insert(nodePair.Index, nodePair.Node);
                    }
                    break;

                case UndoAction.ActionType.Split:
                    // Hoàn tác tách: Gộp lại các mục đã tách
                    for (int i = 0; i < action.NodeList.Count; i++)
                    {
                        _docNodes.RemoveAt(action.SourceIndex);
                    }
                    _docNodes.Insert(action.SourceIndex, action.Node);
                    break;

                case UndoAction.ActionType.SetMixable:
                    // Hoàn tác thay đổi trạng thái trộn
                    foreach (var selectedNode in action.NodeList)
                    {
                        selectedNode.IsMixable = action.OldMixableState;
                    }
                    break;
            }

            StatusTextBlock.Text = "Đã hoàn tác thao tác";
            _isUndoing = false;
        }

        // Ghi lại hành động di chuyển vào ngăn xếp Undo
        private void RecordMoveAction(int sourceIndex, int targetIndex)
        {
            if (_isUndoing) return;

            var action = new UndoAction(UndoAction.ActionType.Move)
            {
                SourceIndex = sourceIndex,
                TargetIndex = targetIndex,
                Node = _docNodes[sourceIndex]
            };

            _undoStack.Push(action);
        }

        //// Cập nhật MoveItemWithAnimation để ghi lại hành động
        //private void MoveItemWithAnimation(int sourceIndex, int targetIndex)
        //{
        //    // Ghi lại hành động cho Undo
        //    RecordMoveAction(sourceIndex, targetIndex);

        //    // Phần code còn lại giữ nguyên...
        //    var movedNode = _docNodes[sourceIndex];
        //    movedNode.IsAnimating = true;
        //    _docNodes.RemoveAt(sourceIndex);
        //    _docNodes.Insert(targetIndex, movedNode);

        //    Task.Delay(500).ContinueWith(_ => {
        //        Application.Current.Dispatcher.Invoke(() => {
        //            movedNode.IsAnimating = false;
        //        });
        //    });

        //    TreeDoc.UpdateLayout();
        //    HighlightItem(targetIndex);
        //    SelectItemByIndex(targetIndex);

        //    StatusTextBlock.Text = "Đã chuyển vị trí phần tử";
        //}


        // Gộp hai phương thức MoveItemWithAnimation thành một
        private void MoveItemWithAnimation(int sourceIndex, int targetIndex, bool isInDetailPanel = false)
        {
            // Chọn ObservableCollection phù hợp
            var nodes = isInDetailPanel ? _detailNodes : _docNodes;

            // Chỉ ghi lại hành động Undo cho panel chính
            if (!isInDetailPanel && !_isUndoing)
            {
                var action = new UndoAction(UndoAction.ActionType.Move)
                {
                    SourceIndex = sourceIndex,
                    TargetIndex = targetIndex,
                    Node = nodes[sourceIndex]
                };

                _undoStack.Push(action);
            }

            // Tạm thời tạo một bản sao của item được kéo
            var movedNode = nodes[sourceIndex];

            // Đánh dấu để TreeView tạo animation khi load lại item
            movedNode.IsAnimating = true;

            // Xóa khỏi vị trí cũ
            nodes.RemoveAt(sourceIndex);

            // Thêm vào vị trí mới
            nodes.Insert(targetIndex, movedNode);

            // Sau một khoảng thời gian, tắt cờ animation
            Task.Delay(500).ContinueWith(_ => {
                Application.Current.Dispatcher.Invoke(() => {
                    movedNode.IsAnimating = false;
                });
            });

            // Cập nhật lựa chọn hiện tại
            var treeView = isInDetailPanel ? TreeDocDetail : TreeDoc;
            treeView.UpdateLayout();
            HighlightItem(targetIndex, isInDetailPanel);
            SelectItemByIndex(targetIndex, isInDetailPanel);

            // Cập nhật status text
            StatusTextBlock.Text = "Đã chuyển vị trí phần tử";
        }




        #region File I/O Methods

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Word files (*.docx)|*.docx"
            };

            if (ofd.ShowDialog() == true)
            {
                try
                {
                    StatusTextBlock.Text = "Đang tải file...";
                    _docNodes = LoadFromWordDocument(ofd.FileName);
                    TreeDoc.ItemsSource = _docNodes;
                    StatusTextBlock.Text = $"Đã tải: {System.IO.Path.GetFileName(ofd.FileName)}";

                    UpdateEmptyState();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi đọc file Word: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusTextBlock.Text = "Lỗi tải file";
                }
            }
        }

        private void ExportFile_Click(object sender, RoutedEventArgs e)
        {
            if (_docNodes.Count == 0)
            {
                MessageBox.Show("Không có nội dung để xuất file!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Word Files (*.docx)|*.docx",
                DefaultExt = ".docx"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    StatusTextBlock.Text = "Đang xuất file...";
                    ExportToWord(sfd.FileName, _docNodes);
                    MessageBox.Show("Xuất file thành công!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    StatusTextBlock.Text = $"Đã xuất: {System.IO.Path.GetFileName(sfd.FileName)}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi xuất file Word: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusTextBlock.Text = "Lỗi xuất file";
                }
            }
        }

        private ObservableCollection<DocNode> LoadFromWordDocument(string filePath)
        {
            var result = new ObservableCollection<DocNode>();
            int tableCount = 1;
            int imageCount = 1;

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                var mainPart = doc.MainDocumentPart;
                var body = mainPart.Document.Body;

                foreach (var element in body.Elements())
                {
                    // Xử lý bảng
                    if (element is Table table)
                    {
                        var tableNode = new DocNode
                        {
                            NodeType = DocNodeType.Table,
                            Content = $"Bảng {tableCount}",
                            TableData = new ObservableCollection<ObservableCollection<string>>(),
                            ImagesInTable = new List<TableImageInfo>()
                        };

                        int r = 0;
                        foreach (var row in table.Elements<TableRow>())
                        {
                            var rowData = new ObservableCollection<string>();
                            int c = 0;
                            foreach (var cell in row.Elements<TableCell>())
                            {
                                // Lấy text trong cell
                                string cellText = string.Join(" ", cell.Descendants<Text>().Select(t => t.Text));

                                // Kiểm tra ảnh trong cell
                                var drawings = cell.Descendants<Drawing>().ToList();
                                foreach (var drawing in drawings)
                                {
                                    // Tìm ID của blip trong drawing
                                    var blips = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList();
                                    foreach (var blip in blips)
                                    {
                                        var embedId = blip.Embed?.Value;
                                        if (!string.IsNullOrEmpty(embedId))
                                        {
                                            var imagePart = (ImagePart)mainPart.GetPartById(embedId);
                                            using (var stream = imagePart.GetStream())
                                            using (var ms = new MemoryStream())
                                            {
                                                stream.CopyTo(ms);
                                                byte[] imageData = ms.ToArray();

                                                string extension = GetExtensionFromContentType(imagePart.ContentType);
                                                string imagePath = Path.Combine(Path.GetTempPath(), $"table_img_{imageCount}{extension}");
                                                File.WriteAllBytes(imagePath, imageData);

                                                cellText += $" [Ảnh {imageCount}]";
                                                tableNode.ImagesInTable.Add(new TableImageInfo
                                                {
                                                    ImagePath = imagePath,
                                                    ImageData = imageData,
                                                    Row = r,
                                                    Column = c
                                                });

                                                imageCount++;
                                            }
                                        }
                                    }
                                }

                                rowData.Add(cellText);
                                c++;
                            }
                            tableNode.TableData.Add(rowData);
                            r++;
                        }

                        result.Add(tableNode);
                        tableCount++;
                    }
                    // Xử lý đoạn văn bản
                    else if (element is Paragraph para)
                    {
                        // Kiểm tra xem đoạn văn có chứa ảnh không
                        var drawing = para.Descendants<Drawing>().FirstOrDefault();
                        if (drawing != null)
                        {
                            // Tìm Blip (thành phần chứa ID ảnh) trong Drawing
                            var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                            if (blip != null && !string.IsNullOrEmpty(blip.Embed?.Value))
                            {
                                var imagePart = (ImagePart)mainPart.GetPartById(blip.Embed.Value);
                                using (var stream = imagePart.GetStream())
                                using (var ms = new MemoryStream())
                                {
                                    stream.CopyTo(ms);
                                    byte[] imageData = ms.ToArray();

                                    string extension = GetExtensionFromContentType(imagePart.ContentType);
                                    string imagePath = Path.Combine(Path.GetTempPath(), $"img_{imageCount}{extension}");
                                    File.WriteAllBytes(imagePath, imageData);

                                    result.Add(new DocNode
                                    {
                                        NodeType = DocNodeType.Image,
                                        Content = $"Ảnh {imageCount}",
                                        ImagePath = imagePath
                                    });

                                    imageCount++;
                                }
                            }
                        }
                        else
                        {
                            string paraText = string.Join(" ", para.Descendants<Text>().Select(t => t.Text));
                            if (!string.IsNullOrWhiteSpace(paraText))
                            {
                                result.Add(new DocNode
                                {
                                    NodeType = DocNodeType.Text,
                                    Content = paraText
                                });
                            }
                        }
                    }
                }
            }

            return result;
        }

        private void ExportToWord(string filePath, ObservableCollection<DocNode> nodes)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                foreach (var node in nodes)
                {
                    switch (node.NodeType)
                    {
                        case DocNodeType.Table:
                            // Tạo bảng với full border và distribute columns
                            var table = CreateTableWithFormattedBorders();

                            // Thêm dữ liệu và ảnh vào bảng
                            int rowIdx = 0;
                            foreach (var rowData in node.TableData)
                            {
                                var row = new TableRow();
                                int colIdx = 0;

                                foreach (var cellData in rowData)
                                {
                                    var cell = new TableCell();

                                    // Thêm nội dung văn bản
                                    var para = new Paragraph(new Run(new Text(cellData ?? "")));
                                    cell.Append(para);

                                    // Thêm ảnh nếu có trong cell này
                                    var cellImages = node.ImagesInTable.Where(i => i.Row == rowIdx && i.Column == colIdx).ToList();
                                    foreach (var img in cellImages)
                                    {
                                        if (File.Exists(img.ImagePath))
                                        {
                                            // Thêm ảnh vào cell
                                            var imgPara = new Paragraph();
                                            AddImageToParagraph(mainPart, imgPara, img.ImageData);
                                            cell.Append(imgPara);
                                        }
                                    }

                                    // Thêm cell vào row
                                    cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                                    row.Append(cell);
                                    colIdx++;
                                }

                                table.Append(row);
                                rowIdx++;
                            }

                            body.Append(table);
                            body.Append(new Paragraph()); // Dòng trống sau bảng
                            break;

                        case DocNodeType.Image:
                            if (!string.IsNullOrEmpty(node.ImagePath) && File.Exists(node.ImagePath))
                            {
                                var imageData = File.ReadAllBytes(node.ImagePath);
                                var imagePara = new Paragraph();

                                // Thêm ảnh vào paragraph
                                AddImageToParagraph(mainPart, imagePara, imageData);

                                body.Append(imagePara);
                                body.Append(new Paragraph()); // Dòng trống sau ảnh
                            }
                            break;

                        case DocNodeType.Text:
                        default:
                            body.Append(new Paragraph(new Run(new Text(node.Content ?? ""))));
                            break;
                    }
                }

                mainPart.Document.Save();
            }
        }

        private Table CreateTableWithFormattedBorders()
        {
            var table = new Table();

            // Định dạng bảng: full border, width 100%, distribute columns
            var tblProps = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 }
                ),
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }, // 100% width
                new TableLayout { Type = TableLayoutValues.Fixed }  // Fixed width table
            );

            table.AppendChild(tblProps);
            return table;
        }

        private void AddImageToParagraph(MainDocumentPart mainPart, Paragraph para, byte[] imageData)
        {
            // Thêm ImagePart vào document
            var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            using (var ms = new MemoryStream(imageData))
            {
                imagePart.FeedData(ms);
            }

            // Tính toán kích thước ảnh
            int width = 0, height = 0;
            using (var ms = new MemoryStream(imageData))
            using (var img = System.Drawing.Image.FromStream(ms))
            {
                width = img.Width;
                height = img.Height;
            }

            // Tính tỷ lệ và giữ nguyên tỷ lệ
            double maxWidthCm = 15.0;  // Chiều rộng tối đa 15cm
            double aspectRatio = (double)height / width;

            // Tính kích thước mới giữ tỷ lệ
            double widthCm = Math.Min(maxWidthCm, width / 96.0 * 2.54); // chuyển px sang cm
            double heightCm = widthCm * aspectRatio; // giữ tỷ lệ

            // Lấy relationshipId
            string relationshipId = mainPart.GetIdOfPart(imagePart);

            // Tạo drawing element với kích thước tính toán
            var element = new Drawing(
                CreateInlineDrawing(relationshipId, widthCm, heightCm)
            );

            para.Append(new Run(element));
        }

        private DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline CreateInlineDrawing(string relationshipId, double widthCm, double heightCm)
        {
            // Chuyển đổi sang EMU (English Metric Units) - 1cm = 360000 EMU
            int emuWidth = (int)(widthCm * 360000);
            int emuHeight = (int)(heightCm * 360000);

            var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = emuWidth, Cy = emuHeight },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = 1U, Name = "Picture" },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks { NoChangeAspect = true }),
                CreateGraphicElement(relationshipId, emuWidth, emuHeight)
            );

            inline.DistanceFromTop = 0;
            inline.DistanceFromBottom = 0;
            inline.DistanceFromLeft = 0;
            inline.DistanceFromRight = 0;

            return inline;
        }

        private DocumentFormat.OpenXml.Drawing.Graphic CreateGraphicElement(string relationshipId, int width, int height)
        {
            // Tạo Graphic element
            var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
            var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData();
            graphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

            // Picture element
            var picture = new DocumentFormat.OpenXml.Drawing.Pictures.Picture();

            // NonVisualPictureProperties
            var nvPicPr = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties();
            var cNvPr = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties { Id = 1U, Name = "Image" };
            var cNvPicPr = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties();
            nvPicPr.Append(cNvPr);
            nvPicPr.Append(cNvPicPr);
            picture.Append(nvPicPr);

            // BlipFill
            var blipFill = new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill();
            var blip = new DocumentFormat.OpenXml.Drawing.Blip { Embed = relationshipId };
            var stretch = new DocumentFormat.OpenXml.Drawing.Stretch(
                new DocumentFormat.OpenXml.Drawing.FillRectangle());
            blipFill.Append(blip);
            blipFill.Append(stretch);
            picture.Append(blipFill);

            // ShapeProperties
            var spPr = new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties();
            var transform = new DocumentFormat.OpenXml.Drawing.Transform2D();
            var offset = new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 };
            var extents = new DocumentFormat.OpenXml.Drawing.Extents { Cx = width, Cy = height };
            transform.Append(offset);
            transform.Append(extents);
            spPr.Append(transform);

            var presetGeometry = new DocumentFormat.OpenXml.Drawing.PresetGeometry { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle };
            presetGeometry.Append(new DocumentFormat.OpenXml.Drawing.AdjustValueList());
            spPr.Append(presetGeometry);
            picture.Append(spPr);

            graphicData.Append(picture);
            graphic.Append(graphicData);

            return graphic;
        }

        private string GetExtensionFromContentType(string contentType)
        {
            switch (contentType)
            {
                case "image/png": return ".png";
                case "image/jpeg":
                case "image/jpg": return ".jpg";
                case "image/gif": return ".gif";
                case "image/bmp": return ".bmp";
                default: return ".png";
            }
        }

        #endregion

        #region Drag & Drop Implementation

        private void TreeDoc_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Lưu điểm bắt đầu kéo
            _startPoint = e.GetPosition(null);
        }

        //private void TreeDoc_PreviewMouseMove(object sender, MouseEventArgs e)
        //{
        //    // Kiểm tra nếu nút trái được nhấn và đang di chuyển
        //    if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
        //    {
        //        Point currentPosition = e.GetPosition(null);

        //        // Kiểm tra khoảng cách để xác định có phải đang kéo hay không
        //        if (Math.Abs(currentPosition.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
        //            Math.Abs(currentPosition.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
        //        {
        //            // Xác định TreeViewItem được kéo
        //            _draggedItem = FindTreeViewItemUnderMouse(e.GetPosition(TreeDoc));

        //            if (_draggedItem != null)
        //            {
        //                _draggedNode = _draggedItem.Header as DocNode;

        //                if (_draggedNode != null)
        //                {
        //                    // Bắt đầu kéo
        //                    _isDragging = true;

        //                    // Thiết lập dữ liệu kéo
        //                    DataObject dragData = new DataObject("DocNodeFormat", _draggedNode);
        //                    DragDrop.DoDragDrop(_draggedItem, dragData, DragDropEffects.Move);

        //                    // Reset trạng thái kéo
        //                    _isDragging = false;
        //                    _draggedItem = null;
        //                    _draggedNode = null;

        //                    // Xóa chỉ báo drop target trên tất cả items
        //                    ClearDropTargetIndicators();
        //                }
        //            }
        //        }
        //    }
        //}

        private void TreeDoc_DragOver(object sender, DragEventArgs e)
        {
            if (_isDragging && _draggedNode != null)
            {
                e.Effects = DragDropEffects.Move;

                // Xóa tất cả chỉ báo trước đó
                ClearDropTargetIndicators();

                // Hiển thị chỉ báo vị trí thả
                Point position = e.GetPosition(TreeDoc);
                TreeViewItem targetItem = FindTreeViewItemUnderMouse(position);

                if (targetItem != null && targetItem != _draggedItem)
                {
                    targetItem.Tag = "DropTarget";
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void TreeDoc_Drop(object sender, DragEventArgs e)
        {
            if (_isDragging && _draggedNode != null)
            {
                // Tìm item đích để thả
                Point position = e.GetPosition(TreeDoc);
                TreeViewItem targetItem = FindTreeViewItemUnderMouse(position);

                if (targetItem != null && targetItem != _draggedItem)
                {
                    DocNode targetNode = targetItem.Header as DocNode;

                    if (targetNode != null)
                    {
                        // Tìm vị trí của item được kéo và item đích
                        int sourceIndex = _docNodes.IndexOf(_draggedNode);
                        int targetIndex = _docNodes.IndexOf(targetNode);

                        if (sourceIndex != -1 && targetIndex != -1)
                        {
                            // Đảo vị trí các item với animation
                            MoveItemWithAnimation(sourceIndex, targetIndex);
                        }
                    }
                }

                // Xóa các chỉ báo drop target
                ClearDropTargetIndicators();
            }
        }

        //private void MoveItemWithAnimation(int sourceIndex, int targetIndex)
        //{
        //    // Tạm thời tạo một bản sao của item được kéo
        //    var movedNode = _docNodes[sourceIndex];

        //    // Đánh dấu để TreeView tạo animation khi load lại item
        //    movedNode.IsAnimating = true;

        //    // Xóa khỏi vị trí cũ
        //    _docNodes.RemoveAt(sourceIndex);

        //    // Thêm vào vị trí mới
        //    _docNodes.Insert(targetIndex, movedNode);

        //    // Sau một khoảng thời gian, tắt cờ animation
        //    Task.Delay(500).ContinueWith(_ => {
        //        Application.Current.Dispatcher.Invoke(() => {
        //            movedNode.IsAnimating = false;
        //        });
        //    });

        //    // Cập nhật lựa chọn hiện tại
        //    TreeDoc.UpdateLayout();
        //    HighlightItem(targetIndex);
        //    SelectItemByIndex(targetIndex);

        //    // Cập nhật status text
        //    StatusTextBlock.Text = "Đã chuyển vị trí phần tử";
        //}

        private void HighlightItem(int index)
        {
            if (index >= 0 && index < _docNodes.Count)
            {
                TreeViewItem item = TreeDoc.ItemContainerGenerator.ContainerFromIndex(index) as TreeViewItem;
                if (item != null)
                {
                    // Hiệu ứng nhấp nháy với màu sắc
                    ColorAnimation colorAnimation = new ColorAnimation
                    {
                        From = Colors.LightYellow,
                        To = Colors.Transparent,
                        Duration = new Duration(TimeSpan.FromSeconds(1.5)),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                    };

                    SolidColorBrush brush = new SolidColorBrush(Colors.LightYellow);
                    item.Background = brush;

                    brush.BeginAnimation(SolidColorBrush.ColorProperty, colorAnimation);
                }
            }
        }

        private void SelectItemByIndex(int index)
        {
            if (index >= 0 && index < _docNodes.Count)
            {
                TreeViewItem item = TreeDoc.ItemContainerGenerator.ContainerFromIndex(index) as TreeViewItem;
                if (item != null)
                {
                    item.IsSelected = true;
                    item.Focus();
                }
            }
        }

        private TreeViewItem FindTreeViewItemUnderMouse(Point point)
        {
            HitTestResult result = VisualTreeHelper.HitTest(TreeDoc, point);
            if (result != null)
            {
                DependencyObject obj = result.VisualHit;
                while (obj != null && !(obj is TreeViewItem))
                {
                    obj = VisualTreeHelper.GetParent(obj);
                }

                return obj as TreeViewItem;
            }

            return null;
        }

        private void ClearDropTargetIndicators()
        {
            foreach (var node in _docNodes)
            {
                TreeViewItem item = GetTreeViewItemForNode(node);
                if (item != null)
                {
                    item.Tag = null;
                }
            }
        }

        private TreeViewItem GetTreeViewItemForNode(DocNode node)
        {
            int index = _docNodes.IndexOf(node);
            if (index >= 0)
            {
                return TreeDoc.ItemContainerGenerator.ContainerFromIndex(index) as TreeViewItem;
            }
            return null;
        }

        #endregion

        #region UI Helper Methods

        private void UpdateEmptyState()
        {
            if (_docNodes != null && _docNodes.Count > 0)
            {
                EmptyStateOverlay.Visibility = Visibility.Collapsed;
            }
            else
            {
                EmptyStateOverlay.Visibility = Visibility.Visible;
                StatusTextBlock.Text = "Chưa có nội dung";
            }
        }

        #endregion
    }
}