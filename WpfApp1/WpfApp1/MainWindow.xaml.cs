using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using Dicom;          
using Dicom.Imaging;  

// 폴더 선택창 
using Forms = System.Windows.Forms;

namespace WpfApp1
{
    // ListView에 뿌릴 Key/Value 아이템
    public class TagItem
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public partial class MainWindow : Window
    {
        // 태그 표시용 컬렉션 
        public ObservableCollection<TagItem> TagItems { get; } = new ObservableCollection<TagItem>();

        private string[] _files = Array.Empty<string>();
        private int _index = 0;

        // 스크롤바 값 바꿀 때 이벤트 중복 호출 방지
        private bool _suppressScrollEvent = false;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;

            // DataContext를 윈도우 자신으로 잡아 바인딩 가능하게
            DataContext = this;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            string folder = PickDicomFolder();
            if (string.IsNullOrEmpty(folder))
            {
                Close();
                return;
            }

            // 폴더 내 dcm 파일 목록 로드 (파일명 정렬)
            _files = Directory.GetFiles(folder, "*.dcm")
                              .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                              .ToArray();

            if (_files.Length == 0)
            {
                MessageBox.Show("선택한 폴더에 .dcm 파일이 없습니다.", "알림");
                Close();
                return;
            }

            // 중간부터 시작
            _index = _files.Length / 2;

            // 스크롤바 범위 세팅
            SliceScrollBar.Minimum = 0;
            SliceScrollBar.Maximum = _files.Length - 1;

            _suppressScrollEvent = true;
            SliceScrollBar.Value = _index;
            _suppressScrollEvent = false;

            ShowDicom(_files[_index]);
        }

        // 폴더 선택
        private string PickDicomFolder()
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                dlg.Description = "DICOM(.dcm) 파일이 있는 폴더를 선택하세요.";
                dlg.ShowNewFolderButton = false;

                var result = dlg.ShowDialog();
                if (result == Forms.DialogResult.OK && Directory.Exists(dlg.SelectedPath))
                    return dlg.SelectedPath;
            }
            return null;
        }

        // 마우스 휠로 이전/다음 슬라이스
        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (_files.Length == 0) return;

            if (e.Delta > 0) _index--;
            else _index++;

            if (_index < 0) _index = 0;
            if (_index >= _files.Length) _index = _files.Length - 1;

            // 스크롤바도 동기화
            _suppressScrollEvent = true;
            SliceScrollBar.Value = _index;
            _suppressScrollEvent = false;

            ShowDicom(_files[_index]);
        }

        // 스크롤바 드래그로 슬라이스 점프
        private void SliceScrollBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_files.Length == 0) return;
            if (_suppressScrollEvent) return;

            int newIndex = (int)Math.Round(SliceScrollBar.Value);
            if (newIndex < 0) newIndex = 0;
            if (newIndex >= _files.Length) newIndex = _files.Length - 1;

            _index = newIndex;
            ShowDicom(_files[_index]);
        }

        // 이미지 표시 + 태그 업데이트
        private void ShowDicom(string path)
        {
            try
            {
                // 이미지 렌더 
                var dicomImage = new DicomImage(path);
                var rendered = dicomImage.RenderImage(0);

                // fo-dicom 4.x 환경에서 Bitmap으로 변환 -> WPF BitmapSource로 변환
                using (System.Drawing.Bitmap bmp = rendered.As<System.Drawing.Bitmap>())
                {
                    IntPtr hBitmap = bmp.GetHbitmap();
                    try
                    {
                        BitmapSource bs = Imaging.CreateBitmapSourceFromHBitmap(
                            hBitmap,
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());

                        bs.Freeze();
                        DicomImageView.Source = bs;
                    }
                    finally
                    {
                        DeleteObject(hBitmap);
                    }
                }

                // 태그 로드 & ListView 갱신
                UpdateTags(path);

                // 타이틀 표시
                Title = $"{System.IO.Path.GetFileName(path)} ({_index + 1}/{_files.Length})";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "DICOM 표시 실패");
            }
        }

        // DICOM 태그를 Key/Value 형태로 ListView에 표시
        private void UpdateTags(string path)
        {
            TagItems.Clear();

            // DICOM 파일 열기
            var dicomFile = DicomFile.Open(path);
            var ds = dicomFile.Dataset;

            // 너무 많은 태그가 있을 수 있어서 일단 전부 넣되 Value를 적당히 잘라서 표시
            foreach (var item in ds)
            {
                string key = item.Tag.DictionaryEntry != null
                    ? item.Tag.DictionaryEntry.Name
                    : item.Tag.ToString();

                // 값이 너무 길면 UI가 난장판 되니까 길이 제한
                string value;
                try
                {
                    value = ds.GetValueOrDefault(item.Tag, 0, "").ToString();
                }
                catch
                {
                    // 값이 배열이거나 복잡해서 GetValueOrDefault가 안 먹는 경우
                    value = item.ToString();
                }

                if (value == null) value = "";

                // 너무 길면 자르기
                const int maxLen = 120;
                if (value.Length > maxLen)
                    value = value.Substring(0, maxLen) + " ...";

                TagItems.Add(new TagItem
                {
                    Key = key,
                    Value = value
                });
            }
        }

        // GDI 자원 정리
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);
    }
}
