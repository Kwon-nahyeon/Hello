using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;   // Win32 API(DllImport) 사용을 위해 필요
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using Dicom.Imaging;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        // 현재 폴더의 dcm 파일 목록
        private string[] _files = Array.Empty<string>();

        // 현재 보고 있는 인덱스(몇 번째 파일인지)
        private int _index = 0;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
        }

        // 프로그램 시작 시 최초 1회 실행
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 시작 파일
            var firstPath = @"C:\Users\knh03\OneDrive\문서\manifest-1765788746020\Pelvic-Reference-Data\Pelvic-Ref-001\11-24-2012-NA-Infinity-1 TreatTx Plan-96820\47628.000000-NewCBCT-96129\1-01.dcm";

            // 폴더 경로만 뽑기
            var dir = Path.GetDirectoryName(firstPath);

            // 폴더 안의 *.dcm 파일들 가져오기 (1-01, 1-02 ... 순서로 정렬)
            _files = Directory.GetFiles(dir, "*.dcm")
                              .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                              .ToArray();

            // firstPath가 목록에서 몇 번째인지 찾아서 그걸 시작 인덱스로
            _index = Array.FindIndex(_files, f => string.Equals(f, firstPath, StringComparison.OrdinalIgnoreCase));
            if (_index < 0) _index = 0;

            // 첫 장 띄우기
            ShowDicom(_files[_index]);
        }

        // 마우스 휠로 이전/다음 장 넘기기
        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (_files.Length == 0) return;

            // 휠 위로 = 이전장, 휠 아래로 = 다음장
            if (e.Delta > 0) _index--;
            else _index++;

            // 인덱스 범위 보정(끝/처음에서 멈춤)
            if (_index < 0) _index = 0;
            if (_index >= _files.Length) _index = _files.Length - 1;

            ShowDicom(_files[_index]);
        }

        // DICOM 파일 1장 화면에 표시하는 함수
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

                        // 창 제목에 현재 파일명 표시
                        Title = $"{System.IO.Path.GetFileName(path)} ({_index + 1}/{_files.Length})";
                    }
                    finally
                    {
                        DeleteObject(hBitmap);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "DICOM 표시 실패");
            }
        }

        [DllImport("gdi32.dll")]    // Win32 API 사용 선언
        private static extern bool DeleteObject(IntPtr hObject);
    }
}
