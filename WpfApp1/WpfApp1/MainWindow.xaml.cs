using System;
using System.Collections.Generic;
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
    public partial class MainWindow : Window
    {
        // 현재 폴더의 dcm 파일 목록
        private string[] _files = Array.Empty<string>();

        // 현재 보고 있는 인덱스
        private int _index = 0;

        // 휠/스크롤바 동기화 중 이벤트 중복 방지용
        private bool _suppressScrollEvent = false;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 실행하자마자 폴더 선택
            string folder = PickDicomFolder();
            if (string.IsNullOrEmpty(folder))
            {
                Close(); // 취소하면 종료
                return;
            }

            // 폴더에서 .dcm 파일 읽어오기(이름순)
            _files = Directory.GetFiles(folder, "*.dcm")
                              .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                              .ToArray();

            if (_files.Length == 0)
            {
                MessageBox.Show("선택한 폴더에 .dcm 파일이 없습니다.", "알림");
                Close();
                return;
            }

            // 스크롤바 범위 설정 (0 ~ 파일개수-1)
            SliceScrollBar.Minimum = 0;
            SliceScrollBar.Maximum = _files.Length - 1;
            SliceScrollBar.Value = 0;

            // 첫 장 표시
            _index = 0;
            ShowDicom(_files[_index]);
        }

        // 폴더 선택창 띄우기
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

        // 마우스 휠로 이전/다음 장
        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (_files.Length == 0) return;

            // 휠 위로 = 이전, 아래로 = 다음
            if (e.Delta > 0) _index--;
            else _index++;

            if (_index < 0) _index = 0;
            if (_index >= _files.Length) _index = _files.Length - 1;

            // 스크롤바 값도 같이 이동
            _suppressScrollEvent = true;
            SliceScrollBar.Value = _index;
            _suppressScrollEvent = false;

            ShowDicom(_files[_index]);
        }

        // 스크롤바로 인덱스 이동
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

        // DICOM 한 장 표시 + 태그 목록 갱신
        private void ShowDicom(string path)
        {
            try
            {
                var dicomImage = new DicomImage(path);
                var rendered = dicomImage.RenderImage(0);

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

                // 태그(키-값) 표시
                UpdateTagList(path);

                // 창 제목도 현재 상태로 표시
                Title = $"{System.IO.Path.GetFileName(path)} ({_index + 1}/{_files.Length})";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "DICOM 표시 실패");
            }
        }

        // DICOM 태그를 Key-Value로 뽑아서 ListView에 표시
        private void UpdateTagList(string path)
        {
            // DICOM 파일 읽기
            var dicomFile = DicomFile.Open(path);
            DicomDataset ds = dicomFile.Dataset;

            // 화면에 보여줄 Key-Value 리스트 만들기
            var items = new List<TagItem>();

            // 모든 태그를 돌면서 Key/Value 만들기
            foreach (var element in ds)
            {
                // 예: (0010,0010) PatientName
                string key = $"{element.Tag} {element.Tag.DictionaryEntry.Name}";

                // 값은 너무 길면 보기 힘드니까 적당히 자르기
                string value;
                try
                {
                    // element.Get<string>() 형태는 VR마다 예외가 날 수 있어서 Get<string> 대신 ToString 활용
                    value = ds.GetValueOrDefault(element.Tag, 0, string.Empty);

                    if (string.IsNullOrWhiteSpace(value))
                        value = element.ToString();
                }
                catch
                {
                    value = element.ToString();
                }

                value = MakePrettyValue(value);

                items.Add(new TagItem { Key = key, Value = value });
            }

            // UI에 바인딩
            TagListView.ItemsSource = items;

            TagHeaderText.Text = $"DICOM Tags  ({System.IO.Path.GetFileName(path)})";
        }

        // 값이 너무 길면 잘라서 보기 좋게 만들기
        private string MakePrettyValue(string v)
        {
            if (v == null) return "";

            // 줄바꿈/탭 정리
            v = v.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();

            // 너무 길면 자르기
            const int max = 120;
            if (v.Length > max)
                v = v.Substring(0, max) + " ...";

            return v;
        }

        // ListView에 넣을 한 줄(키/값)
        private class TagItem
        {
            public string Key { get; set; }
            public string Value { get; set; }
        }

        // Windows GDI 자원 해제를 위한 함수 선언
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);
    }
}
